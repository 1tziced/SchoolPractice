from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse
from sqlalchemy.orm import Session
from typing import List
import models
import schemas
from database import engine, get_db
from export_utils import create_student_certificate, create_schedule_excel, create_student_certificate_pdf
import urllib.parse

models.Base.metadata.create_all(bind=engine)

app = FastAPI(title="Учебный учет")

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# API endpoints для студентов
@app.get("/api/students", response_model=List[schemas.Student])
def get_students(db: Session = Depends(get_db)):
    return db.query(models.Student).all()


@app.post("/api/students", response_model=schemas.Student)
def create_student(student: schemas.StudentCreate, db: Session = Depends(get_db)):
    db_student = models.Student(**student.dict())
    db.add(db_student)
    db.commit()
    db.refresh(db_student)
    return db_student


@app.put("/api/students/{student_id}", response_model=schemas.Student)
def update_student(student_id: int, student: schemas.StudentCreate, db: Session = Depends(get_db)):
    db_student = db.query(models.Student).filter(models.Student.id == student_id).first()
    if not db_student:
        raise HTTPException(status_code=404, detail="Студент не найден")
    for key, value in student.dict().items():
        setattr(db_student, key, value)
    db.commit()
    db.refresh(db_student)
    return db_student


@app.delete("/api/students/{student_id}")
def delete_student(student_id: int, db: Session = Depends(get_db)):
    db_student = db.query(models.Student).filter(models.Student.id == student_id).first()
    if not db_student:
        raise HTTPException(status_code=404, detail="Студент не найден")
    db.delete(db_student)
    db.commit()
    return {"message": "Студент удален"}


# API endpoints для групп
@app.get("/api/groups", response_model=List[schemas.Group])
def get_groups(db: Session = Depends(get_db)):
    return db.query(models.Group).all()


@app.post("/api/groups", response_model=schemas.Group)
def create_group(group: schemas.GroupCreate, db: Session = Depends(get_db)):
    db_group = models.Group(**group.dict())
    db.add(db_group)
    db.commit()
    db.refresh(db_group)
    return db_group


@app.put("/api/groups/{group_id}", response_model=schemas.Group)
def update_group(group_id: int, group: schemas.GroupCreate, db: Session = Depends(get_db)):
    db_group = db.query(models.Group).filter(models.Group.id == group_id).first()
    if not db_group:
        raise HTTPException(status_code=404, detail="Группа не найдена")
    for key, value in group.dict().items():
        setattr(db_group, key, value)
    db.commit()
    db.refresh(db_group)
    return db_group


@app.delete("/api/groups/{group_id}")
def delete_group(group_id: int, db: Session = Depends(get_db)):
    db_group = db.query(models.Group).filter(models.Group.id == group_id).first()
    if not db_group:
        raise HTTPException(status_code=404, detail="Группа не найдена")
    db.delete(db_group)
    db.commit()
    return {"message": "Группа удалена"}


# API endpoints для предметов
@app.get("/api/subjects", response_model=List[schemas.Subject])
def get_subjects(db: Session = Depends(get_db)):
    return db.query(models.Subject).all()


@app.post("/api/subjects", response_model=schemas.Subject)
def create_subject(subject: schemas.SubjectCreate, db: Session = Depends(get_db)):
    db_subject = models.Subject(**subject.dict())
    db.add(db_subject)
    db.commit()
    db.refresh(db_subject)
    return db_subject


@app.put("/api/subjects/{subject_id}", response_model=schemas.Subject)
def update_subject(subject_id: int, subject: schemas.SubjectCreate, db: Session = Depends(get_db)):
    db_subject = db.query(models.Subject).filter(models.Subject.id == subject_id).first()
    if not db_subject:
        raise HTTPException(status_code=404, detail="Предмет не найден")
    for key, value in subject.dict().items():
        setattr(db_subject, key, value)
    db.commit()
    db.refresh(db_subject)
    return db_subject


@app.delete("/api/subjects/{subject_id}")
def delete_subject(subject_id: int, db: Session = Depends(get_db)):
    db_subject = db.query(models.Subject).filter(models.Subject.id == subject_id).first()
    if not db_subject:
        raise HTTPException(status_code=404, detail="Предмет не найден")
    db.delete(db_subject)
    db.commit()
    return {"message": "Предмет удален"}


# API endpoints для расписания
@app.get("/api/schedule", response_model=List[schemas.Schedule])
def get_schedule(group_id: int = None, db: Session = Depends(get_db)):
    query = db.query(models.Schedule)
    if group_id:
        query = query.filter(models.Schedule.group_id == group_id)
    return query.all()


@app.post("/api/schedule", response_model=schemas.Schedule)
def create_schedule(schedule: schemas.ScheduleCreate, db: Session = Depends(get_db)):
    db_schedule = models.Schedule(**schedule.dict())
    db.add(db_schedule)
    db.commit()
    db.refresh(db_schedule)
    return db_schedule


@app.put("/api/schedule/{schedule_id}", response_model=schemas.Schedule)
def update_schedule(schedule_id: int, schedule: schemas.ScheduleCreate, db: Session = Depends(get_db)):
    db_schedule = db.query(models.Schedule).filter(models.Schedule.id == schedule_id).first()
    if not db_schedule:
        raise HTTPException(status_code=404, detail="Расписание не найдено")
    for key, value in schedule.dict().items():
        setattr(db_schedule, key, value)
    db.commit()
    db.refresh(db_schedule)
    return db_schedule


@app.delete("/api/schedule/{schedule_id}")
def delete_schedule(schedule_id: int, db: Session = Depends(get_db)):
    db_schedule = db.query(models.Schedule).filter(models.Schedule.id == schedule_id).first()
    if not db_schedule:
        raise HTTPException(status_code=404, detail="Расписание не найдено")
    db.delete(db_schedule)
    db.commit()
    return {"message": "Расписание удалено"}


# HTML страница
@app.get("/")
async def root():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


# === ЭКСПОРТ ДОКУМЕНТОВ ===

@app.get("/api/export/student/{student_id}/certificate-word")
def export_student_certificate_word(student_id: int, db: Session = Depends(get_db)):
    """Экспорт справки студента в Word"""
    student = db.query(models.Student).filter(models.Student.id == student_id).first()
    if not student:
        raise HTTPException(status_code=404, detail="Студент не найден")

    group_name = None
    if student.group_id:
        group = db.query(models.Group).filter(models.Group.id == student.group_id).first()
        group_name = group.name if group else None

    buffer = create_student_certificate(student, group_name)

    # Безопасное имя файла без русских символов
    filename = f"certificate_{student.surname}_{student.name}.docx"
    safe_filename = urllib.parse.quote(filename)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename={safe_filename}; filename*=UTF-8''{safe_filename}"
        }
    )


@app.get("/api/export/student/{student_id}/certificate-pdf")
def export_student_certificate_pdf(student_id: int, db: Session = Depends(get_db)):
    """Экспорт справки студента в PDF"""
    student = db.query(models.Student).filter(models.Student.id == student_id).first()
    if not student:
        raise HTTPException(status_code=404, detail="Студент не найден")

    group_name = None
    if student.group_id:
        group = db.query(models.Group).filter(models.Group.id == student.group_id).first()
        group_name = group.name if group else None

    buffer = create_student_certificate_pdf(student, group_name)

    # Безопасное имя файла без русских символов
    filename = f"certificate_{student.surname}_{student.name}.pdf"
    safe_filename = urllib.parse.quote(filename)

    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f"attachment; filename={safe_filename}; filename*=UTF-8''{safe_filename}"
        }
    )


@app.get("/api/export/schedule/{group_id}/excel")
def export_schedule_excel(group_id: int, db: Session = Depends(get_db)):
    """Экспорт расписания группы в Excel"""
    group = db.query(models.Group).filter(models.Group.id == group_id).first()
    if not group:
        raise HTTPException(status_code=404, detail="Группа не найдена")

    schedules = db.query(models.Schedule).filter(models.Schedule.group_id == group_id).all()
    subjects = db.query(models.Subject).all()
    subjects_dict = {s.id: s.name for s in subjects}

    buffer = create_schedule_excel(group, schedules, subjects_dict)

    # Безопасное имя файла без русских символов
    filename = f"schedule_{group.name}.xlsx"
    safe_filename = urllib.parse.quote(filename)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename={safe_filename}; filename*=UTF-8''{safe_filename}"
        }
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)