from flask import Flask, render_template, send_from_directory, flash, redirect, url_for, g
from flask_migrate import Migrate
from extensions import db  # Импортируем db из extensions.py
import sys
import os
import threading
import signal
import time
import locale
import logging
from sqlalchemy import distinct, select, func, desc, text
from sqlalchemy.exc import IntegrityError, OperationalError

from config.platform import Config, IS_WINDOWS, IS_LINUX

# Импортируем Blueprint после инициализации db
from students import bp as students
from exams import bp as exams
from settings import bp as settings
from teachers import bp as teachers
from events import bp as events
from departments import bp as departments
from method import bp as method

from models import Teacher, Student, Department, Concert, Contest, MethodAssembly, StudentStatus, Region, ExamType, Subject
from forms import MethodAssemblyForm
from utils import get_term, get_academic_year
from migrations import apply_migrations

locale.setlocale(locale.LC_TIME, 'ru_RU.utf-8')
app = Flask(__name__)

logger = logging.getLogger(__name__)

# Конфигурация приложения
app.config['SECRET_KEY'] = Config.SECRET_KEY
app.config['SQLALCHEMY_DATABASE_URI'] = Config.SQLALCHEMY_DATABASE_URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = Config.SQLALCHEMY_TRACK_MODIFICATIONS
app.config['UPLOAD_FOLDER'] = Config.UPLOAD_FOLDER
app.config['DOCS_FOLDER'] = Config.DOCS_FOLDER

# Инициализация расширений
db.init_app(app)
migrate = Migrate(app, db, render_as_batch=True)

# Регистрация роутов
app.register_blueprint(students)
app.register_blueprint(exams)
app.register_blueprint(settings)
app.register_blueprint(teachers)
app.register_blueprint(events)
app.register_blueprint(departments)
app.register_blueprint(method)

# Глобальные переменные для управления сервером
server = None
server_thread = None

# Глобальная переменная для graceful shutdown
shutdown_flag = False

CURRENT_DB_VERSION = 4  # Текущая версия схемы

def setup_database():
    """Инициализация БД и системы версионирования"""
    try:
        # Создаём таблицу версий если её нет
        db.session.execute(text('''
            CREATE TABLE IF NOT EXISTS db_version (
                id INTEGER PRIMARY KEY,
                version INTEGER NOT NULL,
                applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        '''))
        
        # Проверяем текущую версию
        result = db.session.execute(text('SELECT version FROM db_version ORDER BY id DESC LIMIT 1'))
        current_version = result.scalar()
        
        if current_version is None:
            # Первая инициализация - устанавливаем версию 1
            db.session.execute(text('INSERT INTO db_version (version) VALUES (:version)'), 
                             {'version': CURRENT_DB_VERSION})
            current_version = CURRENT_DB_VERSION
        
        db.session.commit()
        return current_version
        
    except Exception as e:
        db.session.rollback()
        print(f"Ошибка инициализации БД: {e}")
        return 0

def check_and_migrate_database():
    """Проверяет и применяет миграции"""
    try:
        current_version = setup_database()
        
        if current_version < CURRENT_DB_VERSION:
            print(f"Применяем миграции БД: {current_version} -> {CURRENT_DB_VERSION}")
            
            # Создаём backup
            create_backup()
            
            # Применяем миграции
            applied = apply_migrations(current_version, CURRENT_DB_VERSION)
            
            if applied:
                print(f"✅ Миграции применены: {applied}")
            else:
                print("❌ Не удалось применить миграции")
                
        else:
            print(f"✅ БД актуальна (версия {current_version})")
            
    except Exception as e:
        print(f"❌ Ошибка миграции: {e}")

def create_backup():
    """Создает резервную копию БД перед миграцией"""
    import shutil
    import datetime
    
    try:
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        db_path = os.path.join(base_dir, 'music_school.db')
        backup_dir = os.path.join(base_dir, 'backups')
        
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(backup_dir, f"backup_{timestamp}.db")
        
        if os.path.exists(db_path):
            shutil.copy2(db_path, backup_path)
            logger.info(f"Backup created: {backup_path}")
            
    except Exception as e:
        logger.error(f"Backup failed: {e}")

def run_server():
    """Запуск сервера в отдельном потоке"""
    global server
    from werkzeug.serving import make_server
    server = make_server('127.0.0.1', 5000, app)
    print("Сервер запущен на http://127.0.0.1:5000")
    server.serve_forever()

def signal_handler(signum, frame):
    """Обработчик сигналов для graceful shutdown"""
    global shutdown_flag
    print(f"Получен сигнал {signum}, завершение работы...")
    shutdown_flag = True
    if server:
        server.shutdown()

# Регистрируем обработчики сигналов
signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

# Специальный роут для статических файлов в frozen-режиме
@app.route('/static/<path:filename>')
def serve_static(filename):
    static_dir = app.config['UPLOAD_FOLDER']
    return send_from_directory(static_dir, filename)


@app.route('/favicon.ico')
def retrieve_favicon():
    return send_from_directory(os.path.join(app.config['UPLOAD_FOLDER'], 'images'), 'favicon.png')

@app.before_request
def get_credentials():
    g.academic_year = get_academic_year()
    g.term = get_term()
    try:
        deps = Department.query.count()
        students = Student.query.count()
        teachers = Teacher.query.count()
        g.d = True if deps else False
        g.s = True if students else False
        g.t = True if teachers else False
        statuses = StudentStatus.query.count()
        regions = Region.query
        if not statuses:
            for status in ["учится", "выпущен(а)", "в академическом отпуске", "отчислен(а)"]:
                st_status = StudentStatus(status=status)
                db.session.add(st_status)
        if not regions.all() or regions.count() != 91:
            from extensions import regions as regions_list
            db.session.execute(text('DELETE FROM regions'))
            db.session.execute(text(regions_list))
        db.session.commit()
    except OperationalError:
        db.create_all()
        statuses = StudentStatus.query.count()
        if not statuses:
            for status in ["учится", "выпущен(а)", "в академическом отпуске", "отчислен(а)"]:
                    st_status = StudentStatus(status=status)
                    db.session.add(st_status)
        regions = Region.query
        if not regions.all() or regions.count() != 91:
            from extensions import regions as regions_list
            db.session.execute(text('DELETE FROM regions'))
            db.session.execute(text(regions_list))
        db.session.commit()
        flash('База данных создана', 'success')
        setup_database()

# Главная страница
@app.route('/')
def index():
    teachers = Teacher.query.count()
    bodies = db.session.execute(select(func.count(distinct(Student.full_name))).filter(Student.status_id.in_([1, 3]))).scalar_one()
    students = Student.query.filter_by(status_id=1).count()
    deps = Department.query.count()
    concerts = Concert.query.filter(Concert.term==get_term()).order_by(Concert.date).all()
    contests = Contest.query.filter(Contest.term==get_term()).order_by(Contest.date).all()
    exam_types = ExamType.query.count()
    subjects = Subject.query.count()
    return render_template('index.html', teachers=teachers, bodies=bodies, students=students, deps=deps, concerts=concerts, contests=contests, exam_types=exam_types, subjects=subjects, title='Главная')


@app.route('/shutdown', methods=['POST'])
def shutdown():
    """Эндпоинт для остановки сервера (работает на всех платформах)"""
    global shutdown_flag
    
    # Устанавливаем флаг завершения
    shutdown_flag = True
    
    # Пытаемся остановить сервер, если он ещё работает
    server_stopped = False
    if server:
        try:
            server.shutdown()
            server_stopped = True
        except:
            pass
    
    # Всегда возвращаем успешный ответ
    if server_stopped:
        return 'Сервер остановлен. Вы можете закрыть это окно.'
    else:
        return 'Сервер уже остановлен. Это окно/вкладку можно закрыть.'

@app.errorhandler(404)
def error404(error):
    return render_template('error.html', e_msg=error, title='Страница не найдена'), 404

@app.errorhandler(500)
def error500(error):
    return render_template('error.html', e_msg=error, title='Ошибка сервера'), 500

@app.errorhandler(403)
def access_forbidden(error):
    err_text = str(error).split(':')[0]
    return render_template('error.html', e_msg=err_text, title='Вам сюда нельзя'), 403

if __name__ == '__main__':
    # Проверяем и применяем миграции БД
    with app.app_context():
        check_and_migrate_database()
    
    # Создаем папку для загрузок если её нет
    if not os.path.exists(app.config['DOCS_FOLDER']):
        os.makedirs(app.config['DOCS_FOLDER'])
    
    # Запуск сервера в отдельном потоке
    server_thread = threading.Thread(target=run_server)
    server_thread.daemon = True
    server_thread.start()
    
    # Открытие браузера (только в Windows или по желанию в Linux)
    if IS_WINDOWS or os.environ.get('AUTO_OPEN_BROWSER', 'false').lower() == 'true':
        Config.open_browser()
    
    print("Сервер запущен. Для остановки нажмите Ctrl+C или используйте кнопку в интерфейсе.")
    
    try:
        # Главный цикл с проверкой флага завершения
        while not shutdown_flag:
            time.sleep(0.5)
        
        print("Завершение работы по запросу пользователя...")
        
        # Даем серверу время на graceful shutdown
        if server_thread.is_alive():
            server_thread.join(timeout=5)
            
    except KeyboardInterrupt:
        print("Получен сигнал прерывания, завершение работы...")
        shutdown_flag = True
        if server:
            server.shutdown()
    
    print("Приложение завершило работу.")
    sys.exit(0)