
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field
from database.db import get_connection, row_to_dict

router = APIRouter(prefix="/api/students", tags=["Students"])


class Student(BaseModel):
    name: str
    email: str = ""
    course: str = "BCA AI & DS"
    attendance: float = Field(ge=0, le=100)
    study_hours: float = Field(ge=0, le=24)
    previous_marks: float = Field(ge=0, le=100)
    assignment_score: float = Field(ge=0, le=100)
    sleep_hours: float = Field(ge=0, le=24)


@router.get("/")
def get_students():
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM students ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return [row_to_dict(row) for row in rows]


@router.post("/")
def add_student(student: Student):
    conn = get_connection()
    cursor = conn.execute("""
        INSERT INTO students
        (name, email, course, attendance, study_hours,
         previous_marks, assignment_score, sleep_hours)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, tuple(student.model_dump().values()))

    conn.commit()
    student_id = cursor.lastrowid
    row = conn.execute(
        "SELECT * FROM students WHERE id = ?", (student_id,)
    ).fetchone()
    conn.close()

    return {"success": True, "student": row_to_dict(row)}


@router.put("/{student_id}")
def update_student(student_id: int, student: Student):
    conn = get_connection()

    cursor = conn.execute("""
        UPDATE students SET
        name=?, email=?, course=?, attendance=?,
        study_hours=?, previous_marks=?,
        assignment_score=?, sleep_hours=?
        WHERE id=?
    """, (*tuple(student.model_dump().values()), student_id))

    conn.commit()
    conn.close()

    if cursor.rowcount == 0:
        raise HTTPException(status_code=404, detail="Student not found")

    return {"success": True, "message": "Student updated"}


@router.delete("/{student_id}")
def delete_student(student_id: int):
    conn = get_connection()
    cursor = conn.execute(
        "DELETE FROM students WHERE id = ?", (student_id,)
    )
    conn.commit()
    conn.close()

    if cursor.rowcount == 0:
        raise HTTPException(status_code=404, detail="Student not found")

    return {"success": True, "message": "Student deleted"}
