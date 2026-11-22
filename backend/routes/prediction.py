
from fastapi import APIRouter
from pydantic import BaseModel, Field
from services.predictor import predict_performance
from database.db import get_connection

router = APIRouter(prefix="/api", tags=["Prediction"])


class StudentData(BaseModel):
    name: str = "Student"
    attendance: float = Field(ge=0, le=100)
    study_hours: float = Field(ge=0, le=24)
    previous_marks: float = Field(ge=0, le=100)
    assignment_score: float = Field(ge=0, le=100)
    sleep_hours: float = Field(ge=0, le=24)


@router.post("/predict")
def predict(data: StudentData):

    features = data.model_dump(exclude={"name"})
    result = predict_performance(features)

    conn = get_connection()

    conn.execute("""
        INSERT INTO predictions
        (student_name, attendance, study_hours, previous_marks,
         assignment_score, sleep_hours, performance, confidence)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data.name,
        data.attendance,
        data.study_hours,
        data.previous_marks,
        data.assignment_score,
        data.sleep_hours,
        str(result["prediction"]),
        result["confidence"]
    ))

    conn.commit()
    conn.close()

    return {
        "success": True,
        "student_performance": result["prediction"],
        "confidence": result["confidence"]
    }
