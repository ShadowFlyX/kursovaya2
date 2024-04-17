from sqlalchemy import Column, Integer, String, Date, Time
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import distinct
from sqlalchemy.orm import session

Base = declarative_base()

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

    def __init__(self, session: session):
        self.session = session

    def get_groups(self, faculty: str) -> list[str | None]:
        """Получить список всех групп из базы данных."""
        try:
            # Выбираем уникальные значения из поля 'SGroup'
            querry = self.session.query(distinct(ScheduleModel.SGroup), ScheduleModel.Faculty, ScheduleModel.Kurs)\
                .filter(ScheduleModel.Faculty == faculty)\
                .order_by(ScheduleModel.Kurs, ScheduleModel.SGroup)
            groups = querry.all()
            return [group[0] for group in groups]  # Преобразуем результат в список строк
        except Exception as e:
            # Обработка ошибок, если что-то пошло не так
            print("Ошибка при получении списка групп:", e)
            return []

    def get_faculties(self) -> list[str | None]:
        """Получить список всех факультетов из базы данных."""
        try:
            # Выбираем уникальные значения из поля 'Faculty'
            querry = self.session.query(distinct(ScheduleModel.Faculty))\
            .order_by(ScheduleModel.Faculty)
            groups = querry.all()
            return [group[0] for group in groups]  # Преобразуем результат в список строк
        except Exception as e:
            # Обработка ошибок, если что-то пошло не так
            print("Ошибка при получении списка факультетов:", e)
            return []