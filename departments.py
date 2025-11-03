from flask import Blueprint, render_template, redirect, url_for, flash, request, send_file
from models import db, Student, Department, Teacher, Concert, ConcertParticipation, Contest, ContestParticipation, Ensemble, DepartmentReportItem, School, Exam
from forms import DepartmentForm, DepartmentReportForm
from utils import get_academic_year, get_term, generate_dep_report, get_deps_students, fetch_all_deps_report
from sqlalchemy.exc import IntegrityError
from sqlalchemy import desc, select, func

bp = Blueprint('departments', __name__, url_prefix='/departments')

@bp.route('/')
def all():
    # Создаем запрос, который сразу возвращает объекты Department с active_count
    stmt = select(
        Department,
        func.count(Student.id).label('active_count')
    ).select_from(Department)\
     .outerjoin(
        Student, 
        (Department.id == Student.department_id) & (Student.status_id == 1)).group_by(Department).order_by(desc(Department.short_name))  # Сортировка по убыванию
    
    # Выполняем запрос
    all_deps = db.session.execute(stmt).all()
    
    # Подготавливаем данные
    deps, totals = [], 0
    for dept, active_count in all_deps:
        dept.active_count = active_count or 0  # Устанавливаем свойство
        deps.append(dept)
        totals += dept.active_count

    students = db.session.query(Student).order_by(Student.class_level, Student.full_name).all()
    for student in students:
        student.ensemble_participations = []
        for ensemble in student.ensembles:
            participations = ConcertParticipation.query.filter_by(
                ensemble_id=ensemble.id
            ).options(
                db.joinedload(ConcertParticipation.concert)
            ).all()
            student.ensemble_participations.extend(participations)

    reports_list = {d.id: [0, 0, 0, 0, 0] for d in Department.query.all()}
    for d in deps:
        for report in d.reports:
            if report.term in [1, 2, 3, 4, 5]:
                reports_list[d.id][report.term-1] = 1
    
    report_avail = {
        1: False,
        2: False,
        3: False,
        4: False,
        5: False
        }

    for i in range(5):
        res = 0
        for dep in Department.query.all():
            if reports_list[dep.id][i]:
                res += 1
        if res == len(deps):
            report_avail[i+1] = True

    return render_template('departments/list.html', deps=deps, title='Программы и отделения', total=totals, dep_reports=reports_list, is_reportable=report_avail)

@bp.route('/<int:id>', methods=['GET', 'POST'])
def view(id):
    department = Department.query.get_or_404(id)
    students = Student.query.filter(Student.department_id==id, Student.status_id==1).order_by(Student.class_level, Student.short_name).all()
    for student in students:
        student.ensemble_participations = []
        student.contest_ens_participations = []
        for ensemble in student.ensembles:
            participations = ConcertParticipation.query.filter_by(
                ensemble_id=ensemble.id
            ).options(
                db.joinedload(ConcertParticipation.concert)
            ).all()
            ens_participations = ContestParticipation.query.filter_by(
                ensemble_id=ensemble.id
            ).options(
                db.joinedload(ContestParticipation.contest)
            ).all()
            student.ensemble_participations.extend(participations)
            student.contest_ens_participations.extend(ens_participations)
    exams = Exam.query.filter_by(department_id=department.id).all()
    return render_template('departments/view.html', department=department, students=students, title=department.title.capitalize(), exams=exams)

@bp.route('/add', methods=['GET', 'POST'])
def add():
    form = DepartmentForm()

    if form.validate_on_submit():
        try:
            dep = Department(
                full_name=form.full_name.data,
                short_name=form.short_name.data,
                title=form.title.data
            )
            db.session.add(dep)
            db.session.commit()
            flash('Отделение успешно добавлено', 'success')
            return redirect(url_for('departments.all'))
        except IntegrityError:
            flash('Такая программа уже есть в системе', 'warning')
    
    return render_template('departments/add.html', form=form, title='Добавление отделения')


@bp.route('/<int:id>/edit', methods=['GET', 'POST'])
def edit(id):
    dep = Department.query.get_or_404(id)
    form = DepartmentForm(obj=dep)
    
    if form.validate_on_submit():
        form.populate_obj(dep)        
        db.session.commit()
        flash('Данные об отделении успешно обновлены', 'success')
        return redirect(url_for('departments.view', id=dep.id))
        
    return render_template('departments/edit.html', form=form, title=f'Изменение данных об отделении <b>{dep.title}</b>', dep=dep)
    

@bp.route('/<int:id>/delete')
def delete(id):
    dep = Department.query.get_or_404(id)
    try:
        db.session.delete(dep)
        db.session.commit()
        flash('Отделение успешно удалено', 'success')
        return redirect(url_for('departments.all'))
    except IntegrityError:
        flash('Невозможно удалить отделение с закреплёнными учениками', 'warning')
        return redirect(url_for('departments.all'))

@bp.route('/<int:id>/get_students')
def get_students(id):
    dep = Department.query.get_or_404(id)
    file_stream = get_deps_students(id)
    filename = f"Список_учеников_{dep.title}.docx"
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@bp.route('/<int:id>/report', methods=['GET', 'POST'])
def report(id):
    dep = Department.query.get_or_404(id)
    form = DepartmentReportForm()
    form.department_id.choices = [(d.id, d.title) for d in Department.query.all()]
    form.department_id.data = id
    students = Student.query.filter_by(department_id=dep.id, status_id=1).count()
    
    if form.validate_on_submit():
        for field in [form.got_avg, form.got_bad, form.got_best, form.got_good]:
            field.data = 0 if field.data is None else field.data
        if students != form.got_best.data + form.got_good.data + form.got_avg.data + form.got_bad.data:
            flash('Количество учеников отделения не совпадает с введёнными данными', 'warning')
            return redirect(url_for('departments.report', id=id))
        try:
            d_report = DepartmentReportItem(
                academic_year=get_academic_year(),
                term=get_term(),
                department_id=id,
                total=students,
                got_best=form.got_best.data,
                got_good=form.got_good.data,
                got_avg=form.got_avg.data,
                got_bad=form.got_bad.data,
                quantity=round((sum([form.got_best.data, form.got_good.data, form.got_avg.data]) / students) * 100),
                quality=round((sum([form.got_best.data, form.got_good.data]) / students) * 100)
            )
            db.session.add(d_report)
            db.session.commit()
            flash(f'Отчёт по успеваемости отделения <b>{dep.title}</b> успешно сохранён', 'success')
            return redirect(url_for('departments.view', id=id))
        except IntegrityError:
            db.session.rollback()
            flash('Такой отчёт за указанный период уже есть', 'warning')
            return redirect(url_for('departments.report', id=id))
    else:
        print(form.errors)
    
    return render_template('departments/add_report.html', dep=dep, title='Отчёт отделения по успеваемости', form=form, students=students)
            
@bp.route('/get_all_students')
def get_all_students():
    file_stream = get_deps_students()
    filename = "Список_учеников_полный.docx"
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@bp.route('/<int:dep_id>/get_report_term_<int:term>')
def get_dep_report(dep_id, term):
    dep = Department.query.filter(Department.id==dep_id).one()
    file_stream = generate_dep_report(dep_id, term)
	
    filename = f"Отчёт об успеваемости, {dep.title}, {get_academic_year()} учебный год.docx"
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@bp.route('/get_all_deps_report/term_<int:term>')
def get_all_deps_report(term):
    # получить отделения
    # deps = Department.query.all()
    # собрать отчёты по классному руководству
    # teachers = Teacher.query.all()
    # teacher_class_reports = ClassReportItem.query.filter_by(term=term, academic_year=get_academic_year()).all()
    # # собрать отчёты по отделениям
    # dep_reports = DepartmentReportItem.query.filter_by(term=term, academic_year=get_academic_year()).all()
    # # вывести в документ
    # print(fetch_all_deps_report(term))
    if School.query.one_or_none() is None:
        flash('Не заполнены данные о школе &ndash; для этого перейдите в раздел <b>Настройки</b> в главном меню', 'warning')
        return redirect(url_for('departments.all'))
    file_stream = fetch_all_deps_report(term)
    filename = f"Отчёт об успеваемости, {term}_{get_academic_year()}.docx"
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
