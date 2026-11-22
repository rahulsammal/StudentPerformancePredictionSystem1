from fastapi import FastAPI, Depends
from fastapi.middleware.cors import CORSMiddleware

from database.db import init_db
from routes import history, prediction, students
from routes.auth import get_current_admin, init_auth_db, router as auth_router

app = FastAPI(
    title="EduPredict - Student Performance Prediction System",
    description="Admin workspace for student records and AI performance predictions.",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Keep the existing student/prediction tables and add the admin accounts table.
init_db()
init_auth_db()

# Public routes: first-time admin registration, login and current session lookup.
app.include_router(auth_router)

# Protected routes: all existing application data requires a valid admin token.
app.include_router(students.router, dependencies=[Depends(get_current_admin)])
app.include_router(prediction.router, dependencies=[Depends(get_current_admin)])
app.include_router(history.router, dependencies=[Depends(get_current_admin)])


@app.get("/")
def root():
    return {"message": "EduPredict API is running", "docs": "/docs"}
