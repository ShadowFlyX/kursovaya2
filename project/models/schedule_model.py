from sqlalchemy import Column, Integer, String, Date, Time
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import distinct
from sqlalchemy.orm import session
from collections import namedtuple

Base = declarative_base()
Schedule = namedtuple('Schedule', [
                'dayofweek', 'timebeg', 'classroom', 'discipline', 'teacher', 'kurs', 'overunderline', 'subgroup', 'weekperiod', 'sgroup', 'semestr'
                ])


class ScheduleModel(Base):
    __tablename__ = 'TimeTable'

    id = Column(Integer, primary_key=True)
    AYear = Column(Integer)
    Semestr = Column(Integer)
    DateBeg = Column(Date)
    DateEnd = Column(Date)
    WeekPeriod = Column(String(25))
    TimeBeg = Column(Time)
    TimeEnd = Column(Time)
    Teacher = Column(String(50))
    ClassRoom = Column(String(50))
    SGroup = Column(String(50))
    SubGroup = Column(String(50))
    Discipline = Column(String(150))
    TypeJob = Column(String(50))
    SubSpecial = Column(String(150))
    Kurs = Column(Integer)
    DayOfWeek = Column(Integer)
    Department = Column(String(150))
    DisciplineShort = Column(String(30))
    TeacherShort = Column(String(50))
    OverUnderLine = Column(String(10))
    Faculty = Column(String(150))


    def get_faculty_groups(self, session: session, faculty: str, semester: int) -> dict[int, str]:
        """Получить список всех групп из базы данных."""
        try:
            query = session.query(distinct(ScheduleModel.SGroup), ScheduleModel.Kurs)\
                .filter(ScheduleModel.Faculty == faculty, ScheduleModel.Semestr == semester)\
                .order_by(ScheduleModel.Kurs, ScheduleModel.SGroup)     
            groups = query.all()
            data = {}
            for group, kurs in groups:
                data.setdefault(kurs, []).append(group)
            return data
        except Exception as e:
            print("Ошибка при получении списка групп:", e)              
            return []


    def get_faculties(self, session: session) -> list[tuple[str]]:
        """Получить список всех факультетов из базы данных."""
        query = session.query(distinct(ScheduleModel.Faculty))\
        .order_by(ScheduleModel.Faculty)                           
        faculties = query.all()
        return [faculty[0] for faculty in faculties]              
        

    def get_schedule_for_week(self, session: session, group: str, semester: int) -> list:
        """Получить список расписаний группы на неделю из базы данных."""
        try:
            query = session.query(ScheduleModel.DayOfWeek, ScheduleModel.TimeBeg, ScheduleModel.ClassRoom, ScheduleModel.DisciplineShort, ScheduleModel.TeacherShort, ScheduleModel.Kurs, ScheduleModel.OverUnderLine, ScheduleModel.SubGroup, ScheduleModel.WeekPeriod, ScheduleModel.SGroup, ScheduleModel.Semestr)\
            .filter(ScheduleModel.SGroup == group, ScheduleModel.Semestr == semester)\
            .order_by(ScheduleModel.DayOfWeek, ScheduleModel.TimeBeg, ScheduleModel.OverUnderLine, ScheduleModel.SubGroup)
            groups = query.all()
            return [Schedule(*group) for group in groups]                                        
        except Exception as e:
            print("Ошибка при получении списка расписания групп:", e)              
            return []


    def get_study_time(self,session: session, faculty: str, semester: int) -> list:
        '''Получить все часы начала занятий'''
        try:
            query = session.query(ScheduleModel.TimeBeg.distinct(), ScheduleModel.Semestr, ScheduleModel.Faculty)\
            .filter(ScheduleModel.Faculty == faculty, ScheduleModel.Semestr == semester)\
            .order_by(ScheduleModel.TimeBeg)
            data = query.all()
            return [d[0] for d in data]                                       
        except Exception as e:
            print("Ошибка при получении списка времени занятий:", e)              
            return []
        

    def get_all_groups_schedule_by_course(self, session: session, faculty: str, semester: int, course: int) -> list:
        try:
            query = session.query(ScheduleModel.DayOfWeek, ScheduleModel.TimeBeg, ScheduleModel.ClassRoom, ScheduleModel.DisciplineShort, ScheduleModel.TeacherShort, ScheduleModel.Kurs, ScheduleModel.OverUnderLine, ScheduleModel.SubGroup, ScheduleModel.WeekPeriod, ScheduleModel.SGroup, ScheduleModel.Semestr)\
            .filter(ScheduleModel.Faculty == faculty, ScheduleModel.Semestr == semester, ScheduleModel.Kurs == course)\
            .order_by(ScheduleModel.SGroup, ScheduleModel.DayOfWeek, ScheduleModel.TimeBeg)
            groups = query.all()
            return [Schedule(*group) for group in groups]                                        
        except Exception as e:
            print("Ошибка при получении списка расписания групп:", e)              
            return []