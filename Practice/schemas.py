from pydantic import BaseModel
from typing import Optional, List


# Student schemas
class StudentBase(BaseModel):
    name: str
    surname: str
    group_id: Optional[int] = None
    email: Optional[str] = None
    phone: Optional[str] = None


class StudentCreate(StudentBase):
    pass


class Student(StudentBase):
    id: int

    class Config:
        from_attributes = True


# Group schemas
class GroupBase(BaseModel):
    name: str
    description: Optional[str] = None


class GroupCreate(GroupBase):
    pass


class Group(GroupBase):
    id: int

    class Config:
        from_attributes = True


# Subject schemas
class SubjectBase(BaseModel):
    name: str
    description: Optional[str] = None


class SubjectCreate(SubjectBase):
    pass


class Subject(SubjectBase):
    id: int

    class Config:
        from_attributes = True


# Schedule schemas
class ScheduleBase(BaseModel):
    group_id: int
    subject_id: int
    day_of_week: str
    lesson_number: int
    room: Optional[str] = None


class ScheduleCreate(ScheduleBase):
    pass


class Schedule(ScheduleBase):
    id: int

    class Config:
        from_attributes = True