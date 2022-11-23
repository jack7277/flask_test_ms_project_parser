import glob
import os

import jpype
import mpxj

from flask_migrate import Config, Migrate
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String, ForeignKey, Boolean
from sqlalchemy.orm import declarative_base, relationship, backref, relation, clear_mappers, mapper, declared_attr, \
    registry
from sqlalchemy.orm import Session

class PTS:
    base = declarative_base()

    def __init__(self, domain):
        # self.base = PTS.base
        # создаю базу по имени домена
        self.domain = domain.lower()
        path = os.path.join(r'..\data', self.domain, "sqlite3_" + self.domain + ".db")
        engine = create_engine(f'sqlite:///{path}')
        engine.connect()
        self.engine = engine
        # Создание таблицы
        self.base.metadata.create_all(engine)

    def get_pts_db(self, domain):
        s = Session(self.engine)
        result = s.query(PTS.PTS_DB).all()
        return result

    def get_users_db(self, domain):
        s = Session(self.engine)
        result = s.query(PTS.Users).all()
        return result

    def parse_mpp_to_pts_db(self, pts):
        from flask_login import current_user
        print('current domain: ', current_user.domain)
        list_files = glob.glob(rf'..\data\{current_user.domain}\*.mpp')
        # parse mpp files
        # set path JAVA_HOME = C:\jre\bin\server
        try:
            jpype.startJVM()
        except Exception as err:
            print(err)

        s = Session(pts.engine)
        from net.sf.mpxj.mpp import MPPReader
        _project_name = ''
        for file in list_files:
            print('File to parse: ', file)
            project = MPPReader().read(file)
            try:
                project_res = project.getResources()
                _project_name = str(project.getTasks()[1].getName())

                for res in project_res:  # ищу емейлы в ресурсах
                    if '@' in str(res.getName()):
                        # сперва ищу емейл в базе по имени проекта и домену
                        # res.getName и др возвращает java:string, поэтому обернул в str
                        result = s.query(PTS.Users).filter(
                            PTS.Users.email == str(res.getName()),
                            PTS.Users.project_name == str(_project_name),
                            PTS.Users.domain == current_user.domain)\
                            .first()

                        # если не нахожу юзера, то создаю пустого
                        if result is None:
                            u2 = PTS.Users(email=str(res.getName()),
                                           project_name=str(_project_name),
                                           domain=str(current_user.domain),
                                           coin_balance=0,
                                           phone_time=0,
                                           email_time=0,
                                           meeting_time=0,
                                           travel_time=0)
                            s.add(u2)
                            s.commit()
            except Exception as e:
                print(e)

            # готовим карточки задания
            nextlevel = 0
            taskname = ''
            namestack = []

            for task in project.getTasks():
                # id = task.getID()
                uid = task.getUniqueID()
                name = task.getName()
                is_milestone = task.getMilestone()
                is_active = task.getActive()
                rollup = task.getRollup()
                predecessors = task.getPredecessors()
                ls = []
                if len(predecessors) != 0:
                    for rel in predecessors:
                        ls.append(rel.getTargetTask().getUniqueID())
                outlinelevel = task.getOutlineLevel()
                outlinenumber = task.getOutlineNumber()
                deadline = task.getDeadline()
                baseline1duration = task.getBaselineDuration(1)  # 15.0h
                resource = task.getResourceAssignments()
                users_assigned = ''
                for i in range(0, len(resource)):
                    try:
                        resname = resource[i].getResource().getName()
                    except:
                        continue
                    if resname is None: continue
                    users_assigned += str(resname)
                    if i < len(resource) - 1:
                        users_assigned += '; '

                if not(outlinelevel is None):
                    if outlinelevel > 1 and outlinelevel != nextlevel:
                        try:
                            taskname = namestack.pop()
                        except:
                            taskname = ''
                # rollup == каталог
                if rollup:
                    namestack.append(taskname)
                    if not(outlinelevel is None):
                        nextlevel = outlinelevel + 1
                    taskname = name
                stepname = f'{str(outlinenumber)}, {name}'
                zeroprogress = 0
                if rollup: zeroprogress = None

                if baseline1duration is None:
                    remainingtime = ''
                else:
                    remainingtime = str(baseline1duration).replace('h', '')

                result = s.query(PTS.PTS_DB).filter(
                    PTS.PTS_DB.uid == uid,
                    PTS.PTS_DB.project_name == _project_name)\
                    .first()

                if result is None:
                    pts_db = pts.PTS_DB(uid=uid,
                                        project_name=_project_name,
                                        task_name=str(taskname),
                                        step_name=str(stepname),
                                        users_assigned=users_assigned,
                                        is_milestone=is_milestone,
                                        remaining_time=str(remainingtime))
                    s.add(pts_db)
                    s.commit()
        try:
            jpype.shutdownJVM()
        except Exception as err:
            print(err)

    class PTS_DB(base):
        __tablename__ = 'pts_db'
        # __table_args__ = {'extend_existing': True}
        id = Column(Integer, primary_key=True, autoincrement=True)
        uid = Column(Integer)
        project_name = Column(String)
        task_name = Column(String)
        step_name = Column(String)
        users_assigned = Column(String)
        user_working = Column(String)
        is_complete = Column(Boolean)
        is_milestone = Column(Boolean)
        actual_work = Column(String)
        remaining_time = Column(String)
        additional_time = Column(String)
        task_note = Column(String)
        actual_start = Column(String)
        actual_finish = Column(String)
        project_coin_budget = Column(Integer)
        # users = relationship('Users', uselist=False)
        # return PTS_DB(self.base)

    class Users(base):
        __tablename__ = 'users'
        # __table_args__ = {'extend_existing': True}
        user_id = Column(Integer, primary_key=True, autoincrement=True)
        email = Column(String)
        fullname = Column(String)
        domain = Column(String)
        coin_balance = Column(Integer)
        project_name = Column(String)  # , ForeignKey('pts_db.project_name'))
        phone_time = Column(String)
        email_time = Column(String)
        meeting_time = Column(String)
        travel_time = Column(String)

    class UserActivity(base):
        __tablename__ = 'user_activity'
        id = Column(Integer, primary_key=True, autoincrement=True)
        user_id = Column(Integer)
        project_uid = Column(Integer)
        project_name = Column(String)
        task_name = Column(String)
        step_name = Column(String)
        # EVENT_NAME: start, pause, end, phone, email, meeting, travel
        event_name = Column(String)
        event_start = Column(String)
        event_finish = Column(String)
