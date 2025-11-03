from flask import Blueprint, render_template, redirect, url_for, flash, session, send_file, request
from models import db, Exam, Student, ExamType, Department, ExamItem, Teacher, Subject, ReportItem, ClassReportItem, LectureItem, OpenLessonItem, ConcertParticipation, ContestParticipation, CourseItem
from forms import TeacherForm, ReportForm, LectureForm, OpenLessonForm, ClassReportForm, TeacherCourseForm
from utils import get_academic_year, get_term
from sqlalchemy.exc import IntegrityError

bp = Blueprint('teachers', __name__, url_prefix='/teachers')

@bp.route('/')
def all():
    teachers = Teacher.query.all()
    t_students = {t.id: Student.query.filter_by(lead_teacher_id=t.id, status_id=1).count() for t in teachers}
    t_c_reports = {t.id: [False, False, False, False, False] for t in teachers}
    all_c_reports = ClassReportItem.query.filter_by(academic_year=get_academic_year())
    for t in teachers:
        for term in range(1, 6):
            for report in all_c_reports:
                if report.term == term and report.teacher_id == t.id:
                    t_c_reports[t.id][term-1] = True
    return render_template('teachers/list.html', teachers=teachers, t_reports=t_c_reports, title='Список преподавателей', students=t_students)

@bp.route('/add', methods=['POST', 'GET'])
def add():
    form = TeacherForm()
    form.main_department_id.choices.extend([(d.id, d.title) for d in Department.query.order_by(Department.short_name)])

    if form.validate_on_submit():
        short_name = form.full_name.data.split(' ', maxsplit=2)
        teacher = Teacher(
            full_name=form.full_name.data,
            short_name=f'{short_name[0]} {short_name[1][0]}. {short_name[2][0]}.',
            main_department_id=form.main_department_id.data,
            is_combining=form.is_combining.data
        )
        db.session.add(teacher)
        db.session.commit()

        flash('Преподаватель добавлен', 'success')
        return redirect(url_for('teachers.all'))
    
    return render_template('teachers/add.html', form=form, title='Добавление преподавателя')


@bp.route('/<int:id>')
def view(id):
    teacher = Teacher.query.get_or_404(id)
    reports = db.session.query(ReportItem).filter(ReportItem.teacher_id==teacher.id).order_by(ReportItem.term).all() if teacher.reports else None
    class_reports = db.session.query(ClassReportItem).filter(ClassReportItem.teacher_id==teacher.id).order_by(ClassReportItem.term).all() if teacher.class_reports else None
    lectures = db.session.query(LectureItem).filter(LectureItem.teacher_id==teacher.id).order_by(LectureItem.term, LectureItem.date).all() if teacher.lecture_items else None
    open_lessons = db.session.query(OpenLessonItem).filter(OpenLessonItem.teacher_id==id).order_by(OpenLessonItem.term, OpenLessonItem.date).all() if teacher.open_lesson_items else None
    s_query = Student.query.filter_by(lead_teacher_id=id, status_id=1).order_by(Student.class_level, Student.full_name)
    students = s_query.all()
    students_count = s_query.count()
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
    return render_template('teachers/view.html', teacher=teacher, title=teacher.full_name, reports=reports, class_reports=class_reports, lectures=lectures, open_lessons=open_lessons, students=students, total=students_count)


@bp.route('/<int:id>/edit', methods=['POST', 'GET'])
def edit(id):
    teacher = Teacher.query.get_or_404(id)
    form = TeacherForm(obj=teacher)
    form.main_department_id.choices.extend([(d.id, d.title) for d in Department.query.order_by(Department.short_name)])

    if form.validate_on_submit():
        short_name = form.full_name.data.split(' ', maxsplit=2)
        teacher.full_name = form.full_name.data
        teacher.short_name = f'{short_name[0]} {short_name[1][0]}. {short_name[2][0]}.'
        teacher.main_department_id = form.main_department_id.data
        teacher.is_combining = form.is_combining.data
        db.session.commit()
        flash('Данные обновлены', 'success')
        return redirect(url_for('teachers.all'))

    form.main_department_id.data = teacher.main_department_id
    return render_template('teachers/edit.html', form=form, title='Редактирование преподавателя', teacher=teacher)

@bp.route('/<int:id>/delete')
def delete(id):
    teacher = Teacher.query.get_or_404(id)

    try:
        db.session.delete(teacher)
        db.session.commit()
    except IntegrityError:
        flash('<span uk-icon="warning"></span> <span class="uk-text-bold">Внимание!</span><br>За преподавателем закреплены ученики. Переведите учеников к другому преподавателю, затем повторите операцию', 'warning')
        return redirect(url_for('teachers.all'))

    flash('Преподаватель удалён', 'success')
    return redirect(url_for('teachers.all'))


@bp.route('<int:id>/report', methods=['GET', 'POST'])
def send_report(id):
    teacher = Teacher.query.get_or_404(id)
    form = ReportForm()
    subjects = Subject.query.all()

    form.subject_id.choices = [(s.id, s.title) for s in subjects]
    if form.validate_on_submit():
        for field in [form.got_avg, form.got_bad, form.got_best, form.got_good]:
            field.data = 0 if field.data is None else field.data
        if form.total.data != form.got_best.data + form.got_good.data + form.got_avg.data + form.got_bad.data:
            flash('Общее количество учеников не совпадает с суммой окончивших четверть на оценки', 'danger')
            return redirect(url_for('teachers.send_report', id=teacher.id))
        try:
            subject = Subject.query.get_or_404(form.subject_id.data)
            report = ReportItem(
                subject_id=form.subject_id.data,
                teacher_id=teacher.id,
                term=form.term.data,
                academic_year=get_academic_year(),
                total=form.total.data,
                got_best=form.got_best.data,
                got_good=form.got_good.data,
                got_avg=form.got_avg.data,
                got_bad=form.got_bad.data,
                quality=round((form.got_best.data + form.got_good.data) / form.total.data * 100),
                quantity=round((form.got_best.data + form.got_good.data + form.got_avg.data) / form.total.data * 100)
            )
            db.session.add(report)
            db.session.commit()
            terms = {
                1: 'I четверть',
                2: 'II четверть',
                3: 'III четверть',
                4: 'IV четверть',
                5: 'весь учебный год'
            }
            flash(f'Отчёт об успеваемости по предмету <i>{subject.title}</i> за {terms[form.term.data]} успешно добавлен', 'success')
            return redirect(url_for('teachers.view', id=teacher.id))
        except IntegrityError:
            db.session.rollback()
            flash('Отчёт за данный период по этому предмету уже сдан', 'warning')
            return redirect(url_for('teachers.send_report', id=teacher.id))
    else:
        print(form.errors)
        
    return render_template('teachers/add_report.html', form=form, title='Добавление отчёта по успеваемости', teacher=teacher)


@bp.route('<int:id>/class_report', methods=['GET', 'POST'])
def send_class_report(id):
    teacher = Teacher.query.get_or_404(id)
    form = ClassReportForm()
    total = Student.query.filter_by(lead_teacher_id=teacher.id, status_id=1).count()
    if request.method == 'GET':
        form.term.data = get_term()
    if form.validate_on_submit():
        for field in [form.got_avg, form.got_bad, form.got_best, form.got_good]:
            field.data = 0 if field.data is None else field.data
        if total != form.got_best.data + form.got_good.data + form.got_avg.data + form.got_bad.data:
            flash('Общее количество учеников не совпадает с введёнными данными', 'danger')
            return redirect(url_for('teachers.send_class_report', id=teacher.id))
        try:
            report = ClassReportItem(
                teacher_id=teacher.id,
                department_id=teacher.main_department_id,
                term=form.term.data,
                academic_year=get_academic_year(),
                total=total,
                got_best=form.got_best.data,
                got_good=form.got_good.data,
                got_avg=form.got_avg.data,
                got_bad=form.got_bad.data,
                quality=round((form.got_best.data + form.got_good.data) / total * 100),
                quantity=round((form.got_best.data + form.got_good.data + form.got_avg.data) / total * 100)
            )
            db.session.add(report)
            db.session.commit()
            terms = {
                1: 'I четверть',
                2: 'II четверть',
                3: 'III четверть',
                4: 'IV четверть',
                5: 'учебный год'
            }
            flash(f'Отчёт о классном руководстве за {terms[form.term.data]} успешно добавлен', 'success')
            return redirect(url_for('teachers.view', id=teacher.id))
        except IntegrityError:
            db.session.rollback()
            flash('Отчёт за данный период уже сдан', 'warning')
            return redirect(url_for('teachers.send_class_report', id=teacher.id))
    else:
        print(form.errors)
        
    return render_template('teachers/add_class_report.html', form=form, title='Добавление отчёта по классному руководству', teacher=teacher, total=total)


@bp.route('/<int:id>/lecture', methods=['GET', 'POST'])
def send_lecture(id):
    form = LectureForm()
    teacher = Teacher.query.get_or_404(id)
    teachers = Teacher.query.all()
    form.resp_teacher_id.choices = [(t.id, t.short_name) for t in teachers]

    if form.validate_on_submit():
        try:
            lecture = LectureItem(
                term=get_term(form.date.data),
                academic_year=get_academic_year(),
                date=form.date.data,
                title=form.title.data,
                teacher_id=teacher.id,
                resp_teacher_id=form.resp_teacher_id.data
            )
            db.session.add(lecture)
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash('Такой доклад за этот период уже зафиксирован', 'warning')
            return redirect(url_for('teachers.view', id=teacher.id))
        flash(f'Методический доклад <b>{form.title.data}</b> успешно добавлен', 'success')
        return redirect(url_for('teachers.view', id=teacher.id))
    else:
        print(form.errors)

    return render_template('teachers/add_lecture.html', form=form, title='Добавление методического доклада', teacher=teacher)


@bp.route('/<int:id>/open_lesson', methods=['GET', 'POST'])
def send_open_lesson(id):
    form = OpenLessonForm()
    teacher = Teacher.query.get_or_404(id)
    teachers = Teacher.query.all()
    students = Student.query.filter(Student.lead_teacher_id==id).all()
    form.resp_teacher_id.choices = [(t.id, t.short_name) for t in teachers]
    form.student_id.choices = [(s.id, f'{s.full_name.split(" ")[0]} {s.full_name.split(" ")[1]}') for s in students]
    if form.validate_on_submit():
        try:
            open_lesson = OpenLessonItem(
                term=get_term(form.date.data),
                academic_year=get_academic_year(form.date.data),
                date=form.date.data,
                title=form.title.data,
                teacher_id=teacher.id,
                resp_teacher_id=form.resp_teacher_id.data,
                student_id=form.student_id.data
            )
            db.session.add(open_lesson)
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash('Такой урок за этот период уже проведён', 'warning')
            return redirect(url_for('teachers.view', id=teacher.id))
        flash(f'Открытый урок <b>{form.title.data}</b> успешно добавлен', 'success')
        return redirect(url_for('teachers.view', id=teacher.id))
    else:
        print(form.errors)

    return render_template('teachers/add_open_lesson.html', form=form, title='Добавление открытого урока', teacher=teacher)

@bp.route('/<int:id>/course', methods=['GET', 'POST'])
def add_course(id):
    form = TeacherCourseForm()
    teacher = Teacher.query.get_or_404(id)
    teachers = Teacher.query.all()

    if form.validate_on_submit():
        try:
            course = CourseItem(
                teacher_id=teacher.id,
                course_type=form.course_type.data,
                title=form.title.data,
                hours=form.hours.data,
                start_date=form.start_date.data,
                end_date=form.end_date.data,
                place=form.place.data,
                cert_no=form.cert_no.data,
                academic_year=get_academic_year(form.end_date.data),
                term=get_term(form.end_date.data)
            )
            db.session.add(course)
            db.session.commit()
            course_type = 'повышения квалификации' if form.course_type.data == 1 else 'профессиональной подготовки'
            flash(f'Курс {course_type} <b>{form.title.data}</b> в объёме {form.hours.data}ч успешно добавлен', 'success')
            return redirect(url_for('teachers.view', id=teacher.id))
        except IntegrityError:
            db.rollback()
            flash('Такой курс уже пройден преподавателем', 'warning')
            return redirect(url_for('teachers.view', id=teacher.id))
    
    return render_template('teachers/add_course.html', title='Добавление курса', form=form, teacher=teacher)
