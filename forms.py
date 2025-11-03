from flask_wtf import FlaskForm
from flask_wtf.file import FileAllowed
from wtforms import StringField, DateField, SelectField, IntegerField, TextAreaField, BooleanField, SelectMultipleField, SubmitField, FileField
from wtforms.validators import DataRequired, Optional
from datetime import date

class StudentForm(FlaskForm):
    full_name = StringField('ФИО ученика', validators=[DataRequired()])
    birth_date = DateField('Дата рождения', default=date.today, validators=[DataRequired()])
    department_id = SelectField('Отделение', coerce=int, validators=[DataRequired()])
    admission_year = IntegerField('Год поступления', validators=[Optional()])
    class_level = IntegerField('Класс', validators=[Optional()])
    study_years = IntegerField('Срок обучения', validators=[DataRequired()])
    status_id = SelectField('Статус', coerce=int, validators=[Optional()])
    lead_teacher_id = SelectField('Основной преподаватель', coerce=int, validators=[DataRequired()])
    contact_phone = StringField('Телефон ученика')
    
    address = StringField('Адрес проживания', validators=[DataRequired()])
    mother_full_name = StringField('ФИО матери/опекуна', validators=[DataRequired()])
    mother_workplace = StringField('Место работы матери', validators=[Optional()])
    mother_occupation = StringField('Должность матери', validators=[Optional()])
    mother_contact_phone = StringField('Телефон матери', validators=[DataRequired()])
    father_full_name = StringField('ФИО отца', validators=[Optional()])
    father_workplace = StringField('Место работы отца', validators=[Optional()])
    father_occupation = StringField('Должность отца', validators=[Optional()])
    father_contact_phone = StringField('Телефон отца', validators=[Optional()])
    is_deep_level = BooleanField('углубленный уровень', validators=[Optional()])
    submit = SubmitField('Добавить ученика')


class SubjectAddForm(FlaskForm):
    title = StringField('Название предмета', validators=[DataRequired()])
    submit = SubmitField('Добавить предмет')


class SubjectEditForm(FlaskForm):
    title = StringField('Новое название предмета', validators=[DataRequired()])
    submit = SubmitField('Изменить предмет')


class DepartmentForm(FlaskForm):
    full_name = StringField('Полное название программы', validators=[DataRequired()])
    short_name = StringField('Краткое название программы', validators=[DataRequired()])
    title = StringField('Название отделения', validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class DismissionForm(FlaskForm):
    dismission_date = DateField('Дата отчисления', validators=[Optional()])
    dismission_reason = StringField('Причина отчисления', validators=[Optional()])
    cert_no = StringField('№ свидетельства об окончании', validators=[Optional()])
    submit = SubmitField('Сохранить')
    

class GraduationForm(FlaskForm):
    dismission_date = DateField('Дата отчисления', validators=[DataRequired()])
    dismission_reason = StringField('Причина отчисления', validators=[DataRequired()])
    cert_no = StringField('№ свидетельства об окончании', validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class ExamTypeForm(FlaskForm):
    name = StringField('Название вида аттестации', validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class ExamStep1Form(FlaskForm):
    exam_type_id = SelectField('Вид аттестации', coerce=int, validators=[DataRequired()])
    date = DateField('Дата зачёта', default=date.today, validators=[DataRequired()])
    department_id = SelectField('Отделение', coerce=int, validators=[DataRequired()])
    discipline = StringField('Дисциплина', validators=[DataRequired()])
    commission = SelectMultipleField('Состав комиссии', coerce=int, validators=[DataRequired()])
    submit = SubmitField('Далее')


class ExamStep2Form(FlaskForm):
    student_ids = SelectMultipleField('Выберите учеников', coerce=int, validators=[DataRequired()])
    submit = SubmitField('Далее')
    

class ExamStep3Form(FlaskForm):
    submit = SubmitField('Сохранить')


class TeacherForm(FlaskForm):
    full_name = StringField('Полное имя преподавателя', validators=[DataRequired()])
    main_department_id = SelectField('Отделение', coerce=int, choices=[(0, '- нет -')])
    is_combining = BooleanField('совместитель', validators=[Optional()])
    submit = SubmitField('Сохранить')

class TeacherCourseForm(FlaskForm):
    course_type = SelectField('Тип курса', coerce=int, choices=[(1, 'курсы повышения квалификации'), (2, 'курсы профессиональной переподготовки')], validators=[DataRequired()])
    title = StringField('Название курса', validators=[DataRequired()])
    hours = IntegerField('Часов', validators=[DataRequired()])
    start_date = DateField('Начало обучения', validators=[DataRequired()])
    end_date = DateField('Окончание обучения', validators=[DataRequired()])
    place = StringField('Место проведения', description='Для онлайн-формата можно не указывать', validators=[Optional()])
    cert_no = StringField('Документ об окончании', description='Укажите полностью серию и номер', validators=[DataRequired()])
    submit = SubmitField('Сохранить')

class ReportForm(FlaskForm):
    subject_id = SelectField('Предмет', coerce=int, validators=[DataRequired()])
    term = SelectField('Период', choices=[(1, 'I ч.'), (2, 'II ч.'), (3, 'III ч.'), (4, 'IV ч.'), (5, 'год')], coerce=int, validators=[DataRequired()])
    total = IntegerField('Всего учеников', validators=[DataRequired()], render_kw={'min': 1})
    got_best = IntegerField('на 5', validators=[Optional()], render_kw={'min': 0})
    got_good = IntegerField('на 4', validators=[Optional()], render_kw={'min': 0})
    got_avg = IntegerField('на 3', validators=[Optional()], render_kw={'min': 0})
    got_bad = IntegerField('на 2', validators=[Optional()], render_kw={'min': 0})
    submit = SubmitField('Сохранить')


class DepartmentReportForm(FlaskForm):
    department_id = SelectField('Отделение', coerce=int, validators=[DataRequired()])
    got_best = IntegerField('на 5', validators=[Optional()], render_kw={'min': 0})
    got_good = IntegerField('на 4', validators=[Optional()], render_kw={'min': 0})
    got_avg = IntegerField('на 3', validators=[Optional()], render_kw={'min': 0})
    got_bad = IntegerField('на 2', validators=[Optional()], render_kw={'min': 0})
    submit = SubmitField('Сохранить')


class ClassReportForm(FlaskForm):
    term = SelectField('Период', choices=[(1, 'I ч.'), (2, 'II ч.'), (3, 'III ч.'), (4, 'IV ч.'), (5, 'год')], coerce=int, validators=[DataRequired()])
    got_best = IntegerField('на 5', validators=[Optional()], render_kw={'min': 0})
    got_good = IntegerField('на 4', validators=[Optional()], render_kw={'min': 0})
    got_avg = IntegerField('на 3', validators=[Optional()], render_kw={'min': 0})
    got_bad = IntegerField('на 2', validators=[Optional()], render_kw={'min': 0})
    submit = SubmitField('Сохранить')


class LectureForm(FlaskForm):
    date = DateField('Дата', default=date.today, validators=[DataRequired()])
    title = StringField('Тема', validators=[DataRequired()])
    resp_teacher_id = SelectField('Ответственный', coerce=int, validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class OpenLessonForm(FlaskForm):
    # term = SelectField('Период', coerce=int, choices=[(1, 'I ч.'), (2, 'II ч.'), (3, 'III ч.'), (4, 'IV ч.')], validators=[DataRequired()])
    date = DateField('Дата', default=date.today, validators=[DataRequired()])
    title = StringField('Тема', validators=[DataRequired()])
    student_id = SelectField('Ученик', coerce=int, validators=[DataRequired()])
    place = StringField('Место проведения', validators=[Optional()])
    resp_teacher_id = SelectField('Ответственный', coerce=int, validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class MethodAssemblyForm(FlaskForm):
    date = DateField('Дата', default=date.today, validators=[DataRequired()])
    teacher_id = SelectField('Ответственный', coerce=int, validators=[DataRequired()])
    title = StringField('Повестка', validators=[DataRequired()])
    description = TextAreaField('Детали', validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class MethodProtocolForm(FlaskForm):
    # Основная информация о заседании
    title = StringField('Заголовок', description='Используется при формировании плана методических заседаний', validators=[DataRequired()])
    date = DateField('Дата заседания', validators=[DataRequired()], default=date.today)
    # Содержание протокола
    agenda = TextAreaField('Повестка', validators=[DataRequired()])  # Повестка дня
    decisions = TextAreaField('Постановили', validators=[DataRequired()])  # Принятые решения
    # Файл протокола
    protocol_file = FileField('Файл протокола', validators=[Optional(), FileAllowed(['pdf', 'doc', 'docx'])])
    submit = SubmitField('Сохранить')


class MethodProtocolUploadForm(FlaskForm):
    protocol_file = FileField('Файл протокола', validators=[DataRequired(), FileAllowed(['pdf', 'doc', 'docx'])])
    submit = SubmitField('Загрузить')


# сначала добавляется коллектив, только затем участники коллектива
class EnsembleForm(FlaskForm):
    name = StringField('Название коллектива', validators=[DataRequired()])
    teacher_id = SelectField('Руководитель', coerce=int, validators=[DataRequired()])
    submit = SubmitField('Добавить коллектив')


class EnsembleMemberForm(FlaskForm):
    ensemble_id = SelectField('Коллектив', validators=[DataRequired()])
    student_id = SelectField('Ученик', validators=[DataRequired()])
    submit = SubmitField('Добавить ученика в коллектив')


# форма концерта (без участников), поля из EventMixin + Concert
# сначала добавить мероприятие, потом участников
class ConcertForm(FlaskForm):
    date = DateField('Дата', default=date.today, validators=[DataRequired()])
    place = StringField('Место проведения', validators=[Optional()])
    title = StringField('Название концерта', validators=[DataRequired()])
    teacher_id = SelectField('Ответственный', coerce=int, validators=[DataRequired()])
    has_passed = BooleanField('мероприятие уже прошло', validators=[Optional()])
    submit = SubmitField('Сохранить')


# форма участников концерта, поля из ParticipationMixin + ConcertParticipation
# использовать, только если есть концерты
class ConcertPartForm(FlaskForm):
    concert_id = SelectField('Мероприятие', coerce=int, validators=[DataRequired()])
    student_id = SelectField('Ученик', coerce=int, choices=[(0, '- нет -')], validators=[Optional()])
    ensemble_id = SelectField('Коллектив', coerce=int, choices=[(0, '- нет -')], validators=[Optional()])
    submit = SubmitField('Сохранить')


# форма конкурса (без участников), поля из EventMixin + Contest
# сначала добавить конкурс, потом участников
class ContestForm(FlaskForm):
    # term = SelectField('Четверть', coerce=int, choices=[(1, 'I ч.'), (2, 'II ч.'), (3, 'III ч.'), (4, 'IV ч.')], validators=[DataRequired()])
    date = DateField('Дата', default=date.today, validators=[DataRequired()])
    place = StringField('Место проведения', validators=[Optional()])
    title = StringField('Название конкурса', validators=[DataRequired()])
    teacher_id = SelectField('Ответственный', coerce=int, validators=[DataRequired()])
    submit = SubmitField('Сохранить')


# форма участников конкурса, поля из ParticipationMixin + ContestParticipation
# использовать, только если есть конкурсы
class ContestPartForm(FlaskForm):
    contest_id = SelectField('Конкурс', coerce=int, validators=[DataRequired()])
    student_id = SelectField('Ученик', coerce=int, choices=[(0, '- нет -')], validators=[Optional()])
    ensemble_id = SelectField('Коллектив', coerce=int, choices=[(0, '- нет -')], validators=[Optional()])
    result = StringField('Результат', validators=[DataRequired()])
    submit = SubmitField('Сохранить')


class SchoolForm(FlaskForm):
    full_title = StringField('Полное наименование ОУ', validators=[DataRequired()])
    short_title = StringField('Сокращённое наименование ОУ', validators=[DataRequired()])
    region_id = SelectField('Субъект', coerce=int, validators=[DataRequired()])
    methodist_id = SelectField('Зав. метод. объединения', coerce=int, validators=[Optional()])
    submit = SubmitField('Сохранить изменения')
