# This file contains an example Flask-User application.
# To keep the example simple, we are applying some unusual techniques:
# - Placing everything in one file
# - Using class-based configuration (instead of file-based configuration)
# - Using string-based templates (instead of file-based templates)

import datetime
import glob
import os
import shutil

import jsons
import keyring
import requests
import win32com
from bs4 import BeautifulSoup
from flask import Flask, request, render_template_string, jsonify
from flask_babelex import Babel
from flask_mail import Mail
from flask_migrate import Migrate, MigrateCommand, Manager
from flask_sqlalchemy import SQLAlchemy
from flask_user import current_user, login_required, roles_required, UserManager, UserMixin

# Class-based application configuration
from livereload import Server
from sqlalchemy.orm import Session

from app import flask_user, utils, EMAIL
from app.MSProject import MSProject
from app.domain_db import PTS
import json

HOME_DIR = r'c:\tjv3'


class ConfigClass(object):
    """ Flask application config """

    # Flask settings
    SECRET_KEY = 'Secret Key mkmkmkmkmkmkmkmkmkmkmkmkmkmkmkmkmkmkmkmk'

    # Flask-SQLAlchemy settings
    SQLALCHEMY_DATABASE_URI = 'sqlite:///timejet.sqlite'  # File-based SQL database
    SQLALCHEMY_TRACK_MODIFICATIONS = False  # Avoids SQLAlchemy warning

    # Flask-Mail SMTP server settings
    MAIL_SERVER = 'smtp.gmail.com'
    MAIL_PORT = 587
    MAIL_USE_SSL = False
    MAIL_USE_TLS = True
    MAIL_USERNAME = 'timejet2020@gmail.com'
    MAIL_PASSWORD = keyring.get_password('system', 'password')
    MAIL_DEFAULT_SENDER = '"TimeJet" <timejet2020@gmail.com>'

    # Flask-User settings
    USER_APP_NAME = "TimeJet"  # Shown in and email templates and page footers
    USER_ENABLE_EMAIL = True  # Enable email authentication
    USER_ENABLE_USERNAME = False  # Disable username authentication
    USER_REQUIRE_RETYPE_PASSWORD = False
    USER_EMAIL_SENDER_NAME = USER_APP_NAME
    USER_EMAIL_SENDER_EMAIL = "timejet2020@gmail.com"


def create_app():
    """ Flask application factory """
    flask_user.__file__ = os.path.join(HOME_DIR, r'app\flask_user')

    # Create Flask app load app.config
    app = Flask(__name__)
    app.config.from_object(__name__ + '.ConfigClass')

    # Initialize Flask-BabelEx
    babel = Babel(app)
    mail = Mail()

    # Initialize Flask-SQLAlchemy
    db = SQLAlchemy(app)

    # Define the User data-model.
    # NB: Make sure to add flask_user UserMixin !!!
    class User(db.Model, UserMixin):
        __tablename__ = 'users'
        id = db.Column(db.Integer, primary_key=True)
        active = db.Column('is_active', db.Boolean(), nullable=False, server_default='1')

        # User authentication information. The collation='NOCASE' is required
        # to search case insensitively when USER_IFIND_MODE is 'nocase_collation'.
        email = db.Column(db.String(255, collation='NOCASE'), nullable=False, unique=True)
        email_confirmed_at = db.Column(db.DateTime())
        password = db.Column(db.String(255), nullable=False, server_default='')

        # User information
        first_name = db.Column(db.String(100, collation='NOCASE'), nullable=False, server_default='')
        last_name = db.Column(db.String(100, collation='NOCASE'), nullable=False, server_default='')
        domain = db.Column(db.String(255, collation='NOCASE'), nullable=False, server_default='')

        # Define the relationship to Role via UserRoles
        roles = db.relationship('Role', secondary='user_roles')

    # Define the Role data-model
    class Role(db.Model):
        __tablename__ = 'roles'
        id = db.Column(db.Integer(), primary_key=True)
        name = db.Column(db.String(50), unique=True)

    # Define the UserRoles association table
    class UserRoles(db.Model):
        __tablename__ = 'user_roles'
        id = db.Column(db.Integer(), primary_key=True)
        user_id = db.Column(db.Integer(), db.ForeignKey('users.id', ondelete='CASCADE'))
        role_id = db.Column(db.Integer(), db.ForeignKey('roles.id', ondelete='CASCADE'))

    # Setup Flask-User and specify the User data-model
    user_manager = UserManager(app, db, User)

    migrate = Migrate(app, db)
    migrate.init_app(app, db)
    # Create all database tables
    db.create_all()

    # Create 'member@example.com' user with no roles
    if not User.query.filter(User.email == 'guest@timejet.com').first():
        user = User(
            email='guest@timejet.com',
            email_confirmed_at=datetime.datetime.utcnow(),
            password=user_manager.hash_password('Aa111111'),
        )
        db.session.add(user)
        db.session.commit()

    # Create 'admin@example.com' user with 'Admin' and 'Agent' roles
    if not User.query.filter(User.email == 'admin@timejet.com').first():
        user = User(
            email='admin@timejet.com',
            email_confirmed_at=datetime.datetime.utcnow(),
            password=user_manager.hash_password('Aa111111'),
        )
        user.roles.append(Role(name='Admin'))
        user.roles.append(Role(name='Agent'))
        db.session.add(user)
        db.session.commit()

    @app.context_processor
    def inject_user():
        _dict = {'email': '',
                 'domain': '',
                 'list_projects': ''}

        if current_user.is_authenticated:
            list_mpp_files = glob.glob(f'{HOME_DIR}\\data\\{current_user.domain}\\*.mpp')
            _listn = list()
            # utils.kill_process_by_name('winproj.exe')
            msp = MSProject()
            for f in list_mpp_files:
                try:
                    msp.load(f)
                    tasks = msp.Project.Tasks
                    project_name = tasks[0].Name  # task name in mpp
                    _listn.append((os.path.basename(f), project_name))
                except Exception as err:
                    print(err)
            msp.close()
            _dict = dict(user_email=current_user.email,
                         user_domain=current_user.domain,
                         list_projects=_listn)

        return _dict

    # The Home page is accessible to anyone
    @app.route('/')
    def home_page():
        if current_user.is_authenticated:
            noreg = """"""
        else:
            noreg = "<p><a href={{ url_for('user.register') }}>{%trans%}Register{%endtrans%}</a></p>" \
                    "<p><a href={{ url_for('user.login') }}>{%trans%}Sign in{%endtrans%}</a></p>"

        return render_template_string("""
                {% extends "flask_user_layout.html" %}
                {% block content %}
                    <h2>{%trans%}Home page{%endtrans%}</h2>
                """ + noreg + """
                    <p><a href={{ url_for('projects_page') }}>{%trans%}Projects Page{%endtrans%}</a></p>
                    <p><a href={{ url_for('admin_page') }}>{%trans%}Admin Page{%endtrans%}</a></p>
                    <p><a href={{ url_for('user.logout') }}>{%trans%}Sign out{%endtrans%}</a></p>
                {% endblock %}
                """)

    def get_fname(project_name, templates_dir):
        files = glob.glob(os.path.join(templates_dir, '*.mpp'))
        # utils.kill_process_by_name('winproj.exe')
        msp = MSProject()
        for file in files:
            msp.load(file)
            tasks = msp.Project.Tasks
            name = tasks[0].Name  # task name in mpp
            if name.lower() == project_name.lower():
                return os.path.basename(file)
        return None

    @app.route('/report/<domain>/<project_name>')
    def get_report(domain, project_name):
        print(domain, project_name)
        # utils.kill_process_by_name('winproj.exe')
        data = f'{domain}<br>{project_name}<br>'
        pts = PTS(domain)
        db_engine = pts.engine
        s = Session(db_engine)
        templates_dir = os.path.join(HOME_DIR, 'data', domain)
        # Ищу имя файла шаблона
        fname = get_fname(project_name, templates_dir)
        if fname is None:
            data += '<br>No such project'
            return data

        project_file = os.path.join(templates_dir, fname)
        from datetime import datetime
        today = datetime.today().strftime('%d_%m_%Y')
        project_reports_basedir = os.path.join(os.path.dirname(project_file), 'reports')
        os.makedirs(project_reports_basedir, exist_ok=True)
        new_fname_project_file = os.path.join(project_reports_basedir, project_name + '_' + today + '.mpp')

        shutil.copy(project_file, new_fname_project_file)
        project_file = new_fname_project_file
        data += new_fname_project_file
        # Заполняю шайлон проект файла из внутренней бд
        # if db_project_all is None: return data

        try:
            # utils.kill_process_by_name('winproj.exe')
            ok = False
            while not ok:
                try:
                    msp = MSProject()
                    ok = msp.load(project_file)
                    print(ok)
                except:
                    print('ALARM ALARM ALARM')

            tasks = msp.Project.Tasks
            for task in tasks:
                task_name = task.Name  # task name in mpp
                task_uniqueid = task.UniqueID
                task_duration = task.Duration  # значение в минутах, в прожекте в часах
                task_resourcenames = task.ResourceNames
                task_type = task.Type  # fixed duration=1, fixed work = 2
                task_summary = task.Summary
                task_actualstart = task.ActualStart
                task_actualfinish = task.ActualFinish
                task_actualwork = task.ActualWork
                task_remainingwork = task.RemainingWork
                task_milestone = task.Milestone
                task_percentComplete = task.PercentComplete
                task_notes = task.Notes

                pts_db_uid = s.query(PTS.PTS_DB) \
                    .filter_by(project_name=project_name,
                               uid=task_uniqueid)\
                    .first()

                # если ничего не нашло во внутренней БД, что делаем?
                if pts_db_uid is None:
                    print('Обработчик pts_db_uid == None')
                    continue

                # начинаю повторять первичный сценарий на vba
                if not pts_db_uid.task_note is None:
                    task.Notes = pts_db_uid.task_note

                if not task_summary:
                    task.Type = 2  # FixedWork
                    try:
                        task.ActualStart = pts_db_uid.actual_start
                    except:
                        if not pts_db_uid.actual_start is None:
                            print(f'Ошибка записи Даты начала задания, {str(pts_db_uid.actual_start)}')

                    try:
                        task.ActualFinish = pts_db_uid.actual_finish
                    except:
                        if not pts_db_uid.actual_finish is None:
                            print(f'Ошибка записи Даты окончания задания, {str(pts_db_uid.actual_finish)}')

                    try:
                        f_act_w = float(pts_db_uid.actual_work) * 60  # в минутах
                        task.ActualWork = f_act_w
                    except:
                        pass

                    if task.Milestone and pts_db_uid.is_milestone:
                        task.PercentComplete = 100

                # if 'Test' in task.Name:
                #     task.Name = '111222'
                #     msp.save()
            msp.saveAndClose()
            s.close()
        except Exception as err:
            # import win32api
            # error_code = err.args[0]
            # e_msg = win32api.FormatMessage(error_code)
            print(err)

        dirname = os.path.dirname(new_fname_project_file)
        new_fname_ext = os.path.basename(new_fname_project_file)
        EMAIL.send_email(fromaddr='timejet2020@gmail.com',
                         toaddr=current_user.email,
                         subject=f'Report for {project_name}',
                         body='Report file attached',
                         file_attach=os.path.join(dirname, new_fname_ext))
        return data

    @app.route('/get/<domain>', methods=['GET', 'POST'])
    def get_data(domain):
        pts = PTS(domain)
        # print(domain)
        pts_db = pts.get_pts_db(domain)
        users = pts.get_users_db(domain)
        list_obj = [pts_db, users]
        # data = 'Users:<br>'
        # for u in users:
        #     data += u.email + ', '
        #
        # data += '<br><br>'
        data = ''

        try:
            import jsonpickle
            _json = jsonpickle.encode(pts_db)
        except Exception as err:
            print(err)
            _json = ''

        # for _pts in pts_db:
        #     data += str(_pts.uid) \
        #             + f'\t| {_pts.users_assigned}' \
        #             + '<br>'
        # data += '<br><br>'
        data += str(_json)
        return data

    @app.route('/set/<domain>', methods=['POST'])
    def set_data(domain):
        content = request.get_json(silent=True)
        return_text = ''
        # Проверяю обязательные поля, ошибка если не всё передано
        required_fields = ["uid", "project_name", "task_name", "step_name",
                  "user_working", "is_complete", "actual_work", "additional_time",
                  "task_note", "actual_start", "actual_finish", "security_token"]
        for field in required_fields:
            try:
                element = content[field]
            except Exception as err:
                    if return_text == '':
                        return_text = 'Errors in json:\n'
                    # return_text += f'no {field}\n'  # перечисление списка обязательных полей
                    # return_text += f'no required field\n'  # перечисление списка обязательных полей

        if return_text == '':  # есть все поля
            return_text = 'ok\n'
            # save data to db

            uid = content["uid"]
            project_name = content["project_name"]
            actual_work = content["actual_work"]
            user_working = content["user_working"]
            is_complete = content["is_complete"]
            additional_time = content["additional_time"]
            task_note = content["task_note"]
            actual_start = content["actual_start"]
            actual_finish = content["actual_finish"]
            security_token = content["security_token"]

            pts = PTS(domain)
            s = Session(pts.engine)
            pts_db = s.query(PTS.PTS_DB).filter(
                PTS.PTS_DB.uid == uid,
                PTS.PTS_DB.project_name == project_name) \
                .first()

            if pts_db is None:
                return_text += f'Project: {project_name} with uid: {uid} not found'
            else:
                try:  # записываем принятые данные
                    pts_db.actual_work = actual_work
                    pts_db.user_working = user_working
                    pts_db.is_complete = bool(is_complete)
                    pts_db.additional_time = additional_time
                    pts_db.task_note = task_note
                    pts_db.actual_start = actual_start
                    pts_db.actual_finish = actual_finish
                    s.add(pts_db)
                    s.commit()
                except Exception as err:
                    return_text += str(err)
        return return_text  # HTTP_OK, 200

    # The Members page is only accessible to authenticated users
    @app.route('/projects', methods=['GET', 'POST'])
    @login_required
    def projects_page():
        if not (current_user.domain is None):
            if not (current_user.domain in ''):
                path = os.path.join(HOME_DIR, 'data', current_user.domain)
                os.makedirs(path, exist_ok=True)
                user_domain_ok = True

        if current_user.domain is None \
                or current_user.domain in '':
            user_domain_ok = False
            import flask
            flask.flash('Error, no domain in profile !')

        # for pts_1 in pts_db:
        #     print(pts_1.uid)

        # upload file save to path
        from flask import request
        try:
            uploaded_file = request.files['file']
            print(str(uploaded_file))
            if uploaded_file.filename != '':
                uploaded_file.save(os.path.join(path, uploaded_file.filename))

                from app.domain_db import PTS
                pts = PTS(current_user.domain)
                # Parsing upload files
                pts.parse_mpp_to_pts_db(pts)
                # get pts_db
                pts_db = pts.get_pts_db(current_user.domain)
            import flask
            flask.flash('Upload success')
        except:
            pass

        # region show upload file form or not, if user_domain_ok
        uploaded_file_template = """
                    <h4>File Upload</h4>
                    <form method="POST" action="" enctype="multipart/form-data">
                    <p><input type="file" name="file"></p>
                    <p><input type="submit" value="Upload file"></p>
                    </form>"""

        if not user_domain_ok:
            uploaded_file_template = """"""
        # endregion
        # s = Session(pts.engine)
        # result = s.query()
        return render_template_string("""
                {% extends "flask_user_layout.html" %}
                {% block content %}
                    <h3>Projects page</h3>
                    """ + uploaded_file_template + """
                    <br>
                    <h4>List of project files:</h4>
                    {% for pfile in list_projects %}
                        {{pfile[0]}}, <a target="_blank" rel="noopener noreferrer" 
                        href="report/{{user_domain}}/{{pfile[1]}}">Report for: {{pfile[1]}}</a><br> 
                    {% endfor %}
                    <br>
                {% endblock %}
                """)
                #  {{user_email}}, {{user_domain}}

    # The Admin page requires an 'Admin' role.
    @app.route('/admin')
    @roles_required('Admin')  # Use of @roles_required decorator
    def admin_page():
        return render_template_string("""
                {% extends "flask_user_layout.html" %}
                {% block content %}
                    <h2>{%trans%}Admin Page{%endtrans%}</h2>
                    <p><a href={{ url_for('user.register') }}>{%trans%}Register{%endtrans%}</a></p>
                    <p><a href={{ url_for('user.login') }}>{%trans%}Sign in{%endtrans%}</a></p>
                    <p><a href={{ url_for('home_page') }}>{%trans%}Home Page{%endtrans%}</a></p>
                    <p><a href={{ url_for('projects_page') }}>{%trans%}Member Page{%endtrans%}</a></p>
                    <p><a href={{ url_for('admin_page') }}>{%trans%}Admin Page{%endtrans%}</a></p>
                    <p><a href={{ url_for('user.logout') }}>{%trans%}Sign out{%endtrans%}</a></p>
                {% endblock %}
                """)

    return app


# Start development web server
if __name__ == '__main__':
    app = create_app()
    # server = Server(app.wsgi_app)
    # server.serve()
    app.run(host='0.0.0.0', port=5000, debug=True)
