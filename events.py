from flask import Blueprint, render_template, redirect, url_for, flash, session, send_file, request
from models import db, Student, Department, Teacher, Concert, ConcertParticipation, Contest, ContestParticipation, Ensemble
from forms import ConcertForm, ConcertPartForm, ContestForm, ContestPartForm
from utils import get_academic_year, events_plan, get_term
from sqlalchemy.exc import IntegrityError
from sqlalchemy import desc, select, func

bp = Blueprint('events', __name__, url_prefix='/events')

@bp.route('/')
def all_events():
    concerts = Concert.query.order_by(Concert.date).all()
    contests = Contest.query.order_by(Contest.date).all()
    return render_template('events/list.html', concerts=concerts, contests=contests, title='Концерты и конкурсы')

@bp.route('/concert/<int:id>')
def concert_view(id):
    concert = Concert.query.get_or_404(id)
    return render_template('events/view_concert.html', concert=concert)

@bp.route('/concert/add', methods=['GET', 'POST'])
def concert_add():
    form = ConcertForm()
    form.teacher_id.choices = [(t.id, t.short_name) for t in Teacher.query.all()]

    if request.args.get('teacher_id', type=int):
        form.teacher_id.data = request.args.get('teacher_id', type=int)

    if form.validate_on_submit():
        concert = Concert(
            term=get_term(form.date.data),
            academic_year=get_academic_year(form.date.data),
            date=form.date.data,
            title=form.title.data,
            teacher_id=form.teacher_id.data,
            place=form.place.data if form.place.data else 'ДМШ',
            has_passed=form.has_passed.data
        )
        db.session.add(concert)
        db.session.commit()
        last_concert = Concert.query.order_by(desc(Concert.id)).limit(1).scalar()
        flash('Концерт успешно добавлен. Теперь добавьте первого участника', 'success')
        return redirect(url_for('events.concert_part_add', c_id=last_concert.id))
        # return redirect(url_for('events.all_events'))
    
    return render_template('events/add_concert.html', form=form, title='Добавление концерта', academic_year=get_academic_year())

@bp.route('/concert/add_participant', methods=['GET', 'POST'])
def concert_part_add():
    form = ConcertPartForm()
    ensembles = Ensemble.query.all()
    students = Student.query.filter(Student.status_id==1).order_by(Student.full_name).all()
    concerts = Concert.query.order_by(desc(Concert.date)).all()

    if db.session.execute(select(func.count(Concert.id))).scalar_one():
        form.concert_id.choices = [(c.id, f'{c.title} ({c.date.strftime("%d.%m.%Y")})') for c in concerts]
    else:
        flash('Сначала необходимо добавить мероприятие', 'warning')
        return redirect(url_for('events.all_events'))

    if ensembles is not None:
        for e in ensembles:
            form.ensemble_id.choices.append((e.id, e.name))

    if students is not None:
        for s in students:
            form.student_id.choices.append((s.id, s.short_name))

    if request.args.get('c_id', type=int):
        form.concert_id.data = request.args.get('c_id', type=int)

    if form.validate_on_submit():
        if not form.student_id.data and not form.ensemble_id.data:
            flash('Нужно добавить или ученика, или коллектив', 'warning')
            return redirect(url_for('events.concert_part_add'))
        elif form.student_id.data and form.ensemble_id.data:
            flash('Нельзя добавить одновременно и ученика, и коллектив', 'warning')
            return redirect(url_for('events.concert_part_add'))
        
        try:
            concert_part = ConcertParticipation(
                concert_id=int(form.concert_id.data),
                student_id=int(form.student_id.data) if form.student_id.data else None,
                ensemble_id=int(form.ensemble_id.data) if form.ensemble_id.data else None
            )
            db.session.add(concert_part)
            db.session.commit()
            flash('Участник концерта успешно добавлен', 'success')
            return redirect(url_for('events.concert_view', id=form.concert_id.data))
        except IntegrityError as ie:
            flash('Этот участник уже принимает участие в концерте', 'warning')
            db.session.rollback()
            return redirect(url_for('events.concert_part_add'))
    return render_template('events/add_concert_participant.html', form=form, title='Добавление участника концерта')

@bp.route('/concert/<int:id>/edit', methods=['GET', 'POST'])
def concert_edit(id):
    concert = Concert.query.get_or_404(id)
    form = ConcertForm(obj=concert)
    form.teacher_id.choices = [(t.id, t.short_name) for t in Teacher.query.all()]
    
    if form.validate_on_submit():
        form.populate_obj(concert)
        db.session.commit()
        flash(f'Концерт <b>{concert.title}</b> успешно обновлен', 'success')
        return redirect(url_for('events.all_events'))
    
    return render_template('events/edit_concert.html', form=form, title='Изменение концерта', academic_year=get_academic_year(), concert=concert)
    

@bp.route('/concert/<int:id>/complete')
def concert_complete(id):
    concert = Concert.query.get_or_404(id)
    concert.has_passed = True
    db.session.commit()
    flash('Статус концерта обновлён', 'success')
    return redirect(url_for('events.all_events'))

@bp.route('/concert/<int:c_id>/delete')
def concert_delete(c_id):
    concert = Concert.query.get_or_404(c_id)
    c_parts = ConcertParticipation.query.filter(ConcertParticipation.concert_id==c_id).all()

    if c_parts:
        for part in c_parts:
            db.session.delete(part)
    else:
        db.session.delete(concert)
    db.session.commit()
    flash('Концерт удалён', 'success')
    return redirect(url_for('events.all_events'))
    
@bp.route('/concert/<int:c_id>/delete_participant/<int:p_id>')
def concert_delete_participant(c_id, p_id):
    concert = Concert.query.get_or_404(c_id)
    c_part = ConcertParticipation.query.get_or_404(p_id)
    db.session.delete(c_part)
    db.session.commit()
    flash(f'Участник удалён из концерта {concert.title}', 'success')
    return redirect(url_for('events.concert_view', id=c_id))


@bp.route('/contest/view/<int:id>')
def contest_view(id):
    contest = Contest.query.get_or_404(id)
    return render_template('events/view_contest.html', contest=contest)

@bp.route('/contest/add', methods=['GET', 'POST'])
def contest_add():
    form = ContestForm()
    form.teacher_id.choices = [(t.id, t.short_name) for t in Teacher.query.filter(Teacher.main_department_id!=0).all()]

    if request.args.get('teacher_id', type=int):
        form.teacher_id.data = request.args.get('teacher_id', type=int)

    if form.validate_on_submit():
        contest = Contest(
            term=get_term(form.date.data),
            academic_year=get_academic_year(form.date.data),
            date=form.date.data,
            title=form.title.data,
            teacher_id=form.teacher_id.data,
            place=form.place.data if form.place.data else 'ДМШ'
        )
        db.session.add(contest)
        db.session.commit()
        flash('Конкурс успешно добавлен. Теперь можно перейти к добавлению участников', 'primary')
        return redirect(url_for('events.all_events'))
    
    return render_template('events/add_contest.html', form=form, title='Добавление конкурса', academic_year=get_academic_year())


@bp.route('/contest/add_participant', methods=['GET', 'POST'])
def contest_part_add():
    form = ContestPartForm()
    ensembles = Ensemble.query.all()
    students = Student.query.filter(Student.status_id==1).order_by(Student.full_name).all()
    contests = Contest.query.order_by(desc(Contest.date)).limit(10).all()
    contest_selected = None

    if db.session.execute(select(func.count(Contest.id))).scalar_one():
        form.contest_id.choices = [(c.id, f'{c.title} ({c.date.strftime("%d.%m.%Y")})') for c in contests]
    else:
        flash('Сначала необходимо добавить мероприятие', 'warning')
        return redirect(url_for('events.all_events'))

    if ensembles is not None:
        form.ensemble_id.choices.extend([(e.id, e.name) for e in ensembles])

    if students is not None:
        form.student_id.choices.extend([(s.id, f'{s.full_name.split(" ")[0]} {s.full_name.split(" ")[1]}') for s in students])

    if request.args.get('contest_id', type=int):
        form.contest_id.data = request.args.get('contest_id', type=int)
        contest_selected = True

    if form.validate_on_submit():
        try:    
            contest_part = ContestParticipation(
                contest_id=form.contest_id.data,
                student_id=form.student_id.data if form.student_id.data else None,
                ensemble_id=form.ensemble_id.data if form.ensemble_id.data else None,
                result=form.result.data if form.result.data != '' else 'участник'
            )
            db.session.add(contest_part)
            db.session.commit()
            flash('Участник конкурса успешно добавлен', 'success')
            return redirect(url_for('events.contest_view', id=form.contest_id.data))
        except IntegrityError as ie:
            if not form.student_id.data and not form.ensemble_id.data:
                flash('Нужно добавить или ученика, или коллектив', 'warning')
            elif form.student_id.data and form.ensemble_id.data:
                flash('Нельзя добавить одновременно и ученика, и коллектив', 'warning')
            else:
                flash('Этот участник уже принимает участие в концерте', 'warning')
            db.session.rollback()
            return redirect(url_for('events.contest_part_add'))
    else:
        print(form.contest_id.data, form.errors)
    
    return render_template('events/add_contest_participant.html', form=form, contest_selected=contest_selected)

@bp.route('/download')
def get_events():
    file_stream = events_plan()
    filename = f'План_мероприятий_{get_academic_year()}.docx'
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
