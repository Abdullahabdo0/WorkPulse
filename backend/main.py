"""
main.py
=======
FastAPI backend for the AI Workplace Burnout Risk Prediction system.
Features:
    - User Authentication (Login, Register with unique username)
    - Role-based Access (Admin vs Employee)
    - Dynamic Employee Re-assessment & History Tracking
    - Live Dashboard Statistics
    - Scikit-learn + DistilBERT + SHAP Integration
    - PostgreSQL & SQLite Dual Support via SQLAlchemy
"""

import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import json
import joblib
import numpy as np
import pandas as pd
from pathlib import Path
from typing import List, Optional
from datetime import datetime
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from sqlalchemy.orm import Session

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from database import init_db, SessionLocal, User, Assessment, hash_password, verify_password
from sentiment import load_sentiment_model, sentiment_risk as _sentiment_risk
from risk import compute_risk
from explain import get_top_reasons

# ── Paths ─────────────────────────────────────────────────────────────────────
DATA_PATH  = ROOT.parent / "data" / "employees.csv"
MODEL_PATH = ROOT.parent / "models" / "workload_model.pkl"

# ── Global state ──────────────────────────────────────────────────────────────
workload_model = None
if MODEL_PATH.exists():
    try:
        workload_model = joblib.load(MODEL_PATH)
    except Exception as e:
        print(f"[Init] Model load note: {e}")

# ── Dependency: DB Session ───────────────────────────────────────────────────
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# ── Startup Lifecycle ─────────────────────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    global workload_model

    # 1. Initialize Database & Tables
    init_db()

    # 2. Load ML Workload Model
    if MODEL_PATH.exists():
        workload_model = joblib.load(MODEL_PATH)
        print(f"[Startup] Workload model loaded from {MODEL_PATH} ✅")
    else:
        print(f"[Startup] ⚠️ Model not found at {MODEL_PATH}.")

    # 3. Load NLP Sentiment Model
    load_sentiment_model()

    # 4. Sync Initial Employees CSV to DB if DB is empty
    db = SessionLocal()
    try:
        employee_count = db.query(User).filter(User.role == "employee").count()
        if employee_count == 0 and DATA_PATH.exists():
            print("[Startup] Seeding initial 50 employees into database...")
            _seed_employees_from_csv(db)
            print("[Startup] Seeding complete ✅")
    finally:
        db.close()

    yield
    print("[Shutdown] Cleaning up resources.")


# ── FastAPI App ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="Tawazon — Enterprise Burnout Risk System",
    description="Full-stack AI Employee Burnout Prediction with Authentication & Database.",
    version="2.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Exception Handler for 422 Validation Errors ──────────────────────────────
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request, exc: RequestValidationError):
    errors = exc.errors()
    error_msgs = []
    for e in errors:
        loc = [str(part) for part in e.get("loc", []) if part != "body"]
        field = ".".join(loc)
        msg = e.get("msg", "Invalid value")
        error_msgs.append(f"{field}: {msg}" if field else msg)
    detail_str = " | ".join(error_msgs)
    print(f"[Validation 422 Error] {request.method} {request.url.path}: {detail_str}")
    return JSONResponse(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        content={"detail": detail_str, "errors": errors}
    )


# ── Pydantic Schemas ──────────────────────────────────────────────────────────
class RegisterRequest(BaseModel):
    username: str = Field(..., min_length=3, max_length=50, example="ahmed123")
    full_name: str = Field(..., min_length=2, max_length=100, example="Ahmed Mohamed")
    password: str = Field(..., min_length=4, max_length=50, example="pass123")


class CreateEmployeeRequest(BaseModel):
    full_name: str = Field(..., min_length=2, max_length=100)
    username: str = Field(..., min_length=3, max_length=50)
    password: str = Field(..., min_length=4, max_length=50)
    gender: Optional[str] = Field(default=None, max_length=20)
    age: Optional[int] = Field(default=None, ge=15, le=100)


class UpdateEmployeeRequest(BaseModel):
    full_name: Optional[str] = Field(None, min_length=2, max_length=100)
    username: Optional[str] = Field(None, min_length=3, max_length=50)
    password: Optional[str] = Field(None, min_length=4, max_length=50)
    gender: Optional[str] = Field(None, max_length=20)
    age: Optional[int] = Field(None, ge=15, le=100)
    avg_daily_hours: Optional[float] = Field(None, ge=0.0, le=24.0)
    overtime_hours_week: Optional[float] = Field(None, ge=0.0, le=80.0)
    absence_days_month: Optional[int] = Field(None, ge=0, le=31)
    late_arrivals_month: Optional[int] = Field(None, ge=0, le=31)
    tasks_assigned: Optional[int] = Field(None, ge=0, le=100)
    vacation_days_last_year: Optional[int] = Field(None, ge=0, le=365)
    feedback: Optional[str] = Field(None, max_length=2000)


class LoginRequest(BaseModel):
    username: str = Field(..., example="admin")
    password: str = Field(..., example="********")


class UserResponse(BaseModel):
    id: int
    username: str
    full_name: str
    gender: Optional[str] = None
    age: Optional[int] = None
    role: str


class AssessmentInput(BaseModel):
    avg_daily_hours: float = Field(..., ge=0.0, le=24.0, example=9.5)
    overtime_hours_week: float = Field(..., ge=0.0, le=80.0, example=12.0)
    absence_days_month: int = Field(..., ge=0, le=31, example=2)
    late_arrivals_month: int = Field(..., ge=0, le=31, example=3)
    tasks_assigned: int = Field(..., ge=0, le=100, example=28)
    vacation_days_last_year: int = Field(..., ge=0, le=365, example=7)
    feedback: Optional[str] = Field(default="", max_length=2000, example="Work is exhausting recently.")


class PredictResponse(BaseModel):
    risk_score: int
    level: str
    workload_probability: float
    sentiment_risk: float
    top_reasons: List[str]


class AssessmentResponse(BaseModel):
    id: int
    avg_daily_hours: float
    overtime_hours_week: float
    absence_days_month: int
    late_arrivals_month: int
    tasks_assigned: int
    vacation_days_last_year: int
    feedback: str
    risk_score: int
    level: str
    workload_probability: float
    sentiment_risk: float
    top_reasons: List[str]
    created_at: str


class EmployeeSummary(BaseModel):
    user_id: int
    employee_id: str
    name: str
    username: str
    gender: Optional[str] = None
    age: Optional[int] = None
    avg_daily_hours: Optional[float] = None
    overtime_hours_week: Optional[float] = None
    absence_days_month: Optional[int] = None
    late_arrivals_month: Optional[int] = None
    tasks_assigned: Optional[int] = None
    vacation_days_last_year: Optional[int] = None
    feedback: Optional[str] = ""
    risk_score: Optional[int] = None
    level: str = "Unassessed"
    last_assessed: str


class HighRiskItem(BaseModel):
    user_id: int
    employee_id: str
    name: str
    risk_score: int
    level: str
    overtime_hours_week: float
    avg_daily_hours: float
    tasks_assigned: int


class StatsResponse(BaseModel):
    total: int
    low: int
    medium: int
    high: int
    low_pct: float
    medium_pct: float
    high_pct: float
    avg_daily_hours: float = 0.0
    avg_overtime_hours: float = 0.0
    avg_tasks: float = 0.0
    avg_risk_score: float = 0.0
    high_risk_list: List[HighRiskItem] = []
    low_metrics: dict = {}
    med_metrics: dict = {}
    high_metrics: dict = {}


NUMERIC_FEATURES = [
    "avg_daily_hours",
    "overtime_hours_week",
    "absence_days_month",
    "late_arrivals_month",
    "tasks_assigned",
    "vacation_days_last_year",
]


# ── Helper Functions ──────────────────────────────────────────────────────────
def _get_workload_prob(data: dict) -> float:
    if workload_model is None:
        h_norm  = (data["avg_daily_hours"] - 6) / 7
        ot_norm = data["overtime_hours_week"] / 25
        ta_norm = data["tasks_assigned"] / 40
        ab_norm = data["absence_days_month"] / 8
        la_norm = data["late_arrivals_month"] / 15
        vac_norm= data["vacation_days_last_year"] / 25
        score = 0.30*h_norm + 0.25*ot_norm + 0.15*ab_norm + 0.10*la_norm + 0.15*ta_norm - 0.05*vac_norm
        return float(np.clip(score, 0, 1))

    df_x = pd.DataFrame([{f: data[f] for f in NUMERIC_FEATURES}])
    proba = workload_model.predict_proba(df_x)[0]
    return float(proba[1])


def _seed_employees_from_csv(db: Session):
    df = pd.read_csv(DATA_PATH)
    if df.empty:
        return

    probs = workload_model.predict_proba(df[NUMERIC_FEATURES])[:, 1] if workload_model else [0.5]*len(df)
    unique_feedbacks = df["feedback"].dropna().unique()
    sentiment_map = {fb: _sentiment_risk(str(fb)) for fb in unique_feedbacks}

    for idx, row in df.iterrows():
        emp_id = str(row["employee_id"])
        full_name = str(row["name"])
        username = emp_id.lower()
        gender_val = str(row["gender"]) if "gender" in row and pd.notna(row["gender"]) else None
        age_val = int(row["age"]) if "age" in row and pd.notna(row["age"]) else None

        user = User(
            username=username,
            full_name=full_name,
            gender=gender_val,
            age=age_val,
            role="employee",
            hashed_password=hash_password("123456")  # default password for seeded accounts
        )
        db.add(user)
        db.flush()

        w_prob = float(probs[idx])
        fb_text = str(row["feedback"])
        s_risk = sentiment_map.get(fb_text, 0.5)
        score, level = compute_risk(w_prob, s_risk)
        reasons = get_top_reasons(workload_model, row.to_dict(), s_risk, top_n=3)

        assessment = Assessment(
            user_id=user.id,
            avg_daily_hours=float(row["avg_daily_hours"]),
            overtime_hours_week=float(row["overtime_hours_week"]),
            absence_days_month=int(row["absence_days_month"]),
            late_arrivals_month=int(row["late_arrivals_month"]),
            tasks_assigned=int(row["tasks_assigned"]),
            vacation_days_last_year=int(row["vacation_days_last_year"]),
            feedback=fb_text,
            risk_score=score,
            level=level,
            workload_probability=round(w_prob, 4),
            sentiment_risk=round(s_risk, 4),
            top_reasons=json.dumps(reasons),
        )
        db.add(assessment)
    db.commit()


# ── Auth Endpoints ────────────────────────────────────────────────────────────

@app.post("/auth/register", response_model=UserResponse, tags=["Auth"])
def register():
    raise HTTPException(
        status_code=status.HTTP_403_FORBIDDEN,
        detail="التسجيل المباشر مغلق. يتم إنشاء حسابات الموظفين بواسطة مسؤول النظام (Admin) فقط."
    )


@app.post("/auth/login", response_model=UserResponse, tags=["Auth"])
def login(req: LoginRequest, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == req.username.lower()).first()
    if not user or not verify_password(req.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="اسم المستخدم أو كلمة المرور غير صحيحة (Invalid username or password)"
        )

    return UserResponse(
        id=user.id,
        username=user.username,
        full_name=user.full_name,
        gender=user.gender,
        age=user.age,
        role=user.role
    )


# ── Direct Prediction Endpoint ───────────────────────────────────────────────

@app.post("/predict", response_model=PredictResponse, tags=["ML Engine"])
def predict_assessment(input_data: AssessmentInput):
    data = input_data.model_dump()
    w_prob = _get_workload_prob(data)
    s_risk = _sentiment_risk(data.get("feedback", ""))
    score, level = compute_risk(w_prob, s_risk)
    reasons = get_top_reasons(workload_model, data, s_risk, top_n=3)
    return PredictResponse(
        risk_score=score,
        level=level,
        workload_probability=round(w_prob, 4),
        sentiment_risk=round(s_risk, 4),
        top_reasons=reasons,
    )


# ── Employee Assessment & Update Endpoints ───────────────────────────────────

@app.post("/assessments/{user_id}", response_model=AssessmentResponse, tags=["Assessments"])
def submit_or_update_assessment(user_id: int, input_data: AssessmentInput, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="المستخدم غير موجود (User not found)")

    data = input_data.model_dump()
    w_prob = _get_workload_prob(data)
    s_risk = _sentiment_risk(data["feedback"])
    score, level = compute_risk(w_prob, s_risk)
    reasons = get_top_reasons(workload_model, data, s_risk, top_n=3)

    new_assessment = Assessment(
        user_id=user.id,
        avg_daily_hours=data["avg_daily_hours"],
        overtime_hours_week=data["overtime_hours_week"],
        absence_days_month=data["absence_days_month"],
        late_arrivals_month=data["late_arrivals_month"],
        tasks_assigned=data["tasks_assigned"],
        vacation_days_last_year=data["vacation_days_last_year"],
        feedback=data["feedback"],
        risk_score=score,
        level=level,
        workload_probability=round(w_prob, 4),
        sentiment_risk=round(s_risk, 4),
        top_reasons=json.dumps(reasons),
    )
    db.add(new_assessment)
    db.commit()
    db.refresh(new_assessment)

    return AssessmentResponse(
        id=new_assessment.id,
        avg_daily_hours=new_assessment.avg_daily_hours,
        overtime_hours_week=new_assessment.overtime_hours_week,
        absence_days_month=new_assessment.absence_days_month,
        late_arrivals_month=new_assessment.late_arrivals_month,
        tasks_assigned=new_assessment.tasks_assigned,
        vacation_days_last_year=new_assessment.vacation_days_last_year,
        feedback=new_assessment.feedback,
        risk_score=new_assessment.risk_score,
        level=new_assessment.level,
        workload_probability=new_assessment.workload_probability,
        sentiment_risk=new_assessment.sentiment_risk,
        top_reasons=reasons,
        created_at=new_assessment.created_at.strftime("%Y-%m-%d %H:%M"),
    )


@app.get("/assessments/{user_id}/history", response_model=List[AssessmentResponse], tags=["Assessments"])
def get_user_assessment_history(user_id: int, db: Session = Depends(get_db)):
    assessments = db.query(Assessment).filter(Assessment.user_id == user_id).order_by(Assessment.created_at.desc()).all()
    results = []
    for a in assessments:
        try:
            reasons = json.loads(a.top_reasons)
        except Exception:
            reasons = ["Workload factors"]
        results.append(AssessmentResponse(
            id=a.id,
            avg_daily_hours=a.avg_daily_hours,
            overtime_hours_week=a.overtime_hours_week,
            absence_days_month=a.absence_days_month,
            late_arrivals_month=a.late_arrivals_month,
            tasks_assigned=a.tasks_assigned,
            vacation_days_last_year=a.vacation_days_last_year,
            feedback=a.feedback,
            risk_score=a.risk_score,
            level=a.level,
            workload_probability=a.workload_probability,
            sentiment_risk=a.sentiment_risk,
            top_reasons=reasons,
            created_at=a.created_at.strftime("%Y-%m-%d %H:%M"),
        ))
    return results


# ── Admin Dashboard & Employees Endpoints ─────────────────────────────────────

@app.get("/employees", response_model=List[EmployeeSummary], tags=["Admin"])
def get_employees(
    level: Optional[str] = None,
    search: Optional[str] = None,
    limit: Optional[int] = None,
    db: Session = Depends(get_db)
):
    users = db.query(User).filter(User.role == "employee").order_by(User.id.desc()).all()
    results = []

    for u in users:
        latest = u.assessments[0] if u.assessments else None
        emp_id = f"EMP{str(u.id).zfill(4)}"

        if latest:
            lvl = latest.level
            hours = latest.avg_daily_hours
            ot = latest.overtime_hours_week
            abs_days = latest.absence_days_month
            late = latest.late_arrivals_month
            tasks = latest.tasks_assigned
            vac = latest.vacation_days_last_year
            fb = latest.feedback
            score = latest.risk_score
            assessed_at = latest.created_at.strftime("%Y-%m-%d %H:%M")
        else:
            lvl = "Unassessed"
            hours = None
            ot = None
            abs_days = None
            late = None
            tasks = None
            vac = None
            fb = ""
            score = None
            assessed_at = "لم يقيّم بعد / Unassessed"

        if level and lvl.lower() != level.lower():
            continue
        if search:
            q = search.lower()
            if q not in u.full_name.lower() and q not in emp_id.lower() and q not in u.username.lower():
                continue

        results.append(EmployeeSummary(
            user_id=u.id,
            employee_id=emp_id,
            name=u.full_name,
            username=u.username,
            gender=u.gender,
            age=u.age,
            avg_daily_hours=hours,
            overtime_hours_week=ot,
            absence_days_month=abs_days,
            late_arrivals_month=late,
            tasks_assigned=tasks,
            vacation_days_last_year=vac,
            feedback=fb,
            risk_score=score,
            level=lvl,
            last_assessed=assessed_at,
        ))

    if limit is not None:
        results = results[:limit]
    return results


@app.post("/employees", response_model=EmployeeSummary, tags=["Admin"])
def create_employee(req: CreateEmployeeRequest, db: Session = Depends(get_db)):
    # Check if username is taken
    existing = db.query(User).filter(User.username == req.username.lower()).first()
    if existing:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="اسم المستخدم مسجل مسبقاً، يرجى اختيار اسم آخر (Username already exists)"
        )

    user = User(
        username=req.username.lower(),
        full_name=req.full_name.strip(),
        gender=req.gender.strip() if req.gender else None,
        age=req.age,
        role="employee",
        hashed_password=hash_password(req.password)
    )
    db.add(user)
    db.commit()
    db.refresh(user)

    emp_id = f"EMP{str(user.id).zfill(4)}"
    return EmployeeSummary(
        user_id=user.id,
        employee_id=emp_id,
        name=user.full_name,
        username=user.username,
        gender=user.gender,
        age=user.age,
        avg_daily_hours=None,
        overtime_hours_week=None,
        absence_days_month=None,
        late_arrivals_month=None,
        tasks_assigned=None,
        vacation_days_last_year=None,
        feedback="",
        risk_score=None,
        level="Unassessed",
        last_assessed="لم يتم التقييم بعد (Pending)"
    )


@app.put("/employees/{user_id}", response_model=EmployeeSummary, tags=["Admin"])
def update_employee(user_id: int, req: UpdateEmployeeRequest, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="الموظف غير موجود (Employee not found)")

    if req.username and req.username.lower() != user.username:
        existing = db.query(User).filter(User.username == req.username.lower(), User.id != user_id).first()
        if existing:
            raise HTTPException(status_code=400, detail="اسم المستخدم مسجل لموظف آخر بالفعل")
        user.username = req.username.lower()

    if req.full_name is not None and req.full_name.strip():
        user.full_name = req.full_name.strip()

    if req.gender is not None:
        user.gender = req.gender.strip() or None

    if req.age is not None:
        user.age = req.age

    if req.password is not None and req.password.strip():
        user.hashed_password = hash_password(req.password.strip())

    latest = user.assessments[0] if user.assessments else None

    # Check if workload metrics were submitted to update or create an assessment
    has_workload_update = any(x is not None for x in [
        req.avg_daily_hours, req.overtime_hours_week, req.absence_days_month,
        req.late_arrivals_month, req.tasks_assigned, req.vacation_days_last_year, req.feedback
    ])

    if has_workload_update or latest:
        avg_daily_hours = req.avg_daily_hours if req.avg_daily_hours is not None else (latest.avg_daily_hours if latest else 8.0)
        overtime_hours_week = req.overtime_hours_week if req.overtime_hours_week is not None else (latest.overtime_hours_week if latest else 0.0)
        absence_days_month = req.absence_days_month if req.absence_days_month is not None else (latest.absence_days_month if latest else 0)
        late_arrivals_month = req.late_arrivals_month if req.late_arrivals_month is not None else (latest.late_arrivals_month if latest else 0)
        tasks_assigned = req.tasks_assigned if req.tasks_assigned is not None else (latest.tasks_assigned if latest else 10)
        vacation_days_last_year = req.vacation_days_last_year if req.vacation_days_last_year is not None else (latest.vacation_days_last_year if latest else 15)
        feedback = req.feedback if req.feedback is not None else (latest.feedback if latest else "")

        data = {
            "avg_daily_hours": avg_daily_hours,
            "overtime_hours_week": overtime_hours_week,
            "absence_days_month": absence_days_month,
            "late_arrivals_month": late_arrivals_month,
            "tasks_assigned": tasks_assigned,
            "vacation_days_last_year": vacation_days_last_year,
            "feedback": feedback
        }
        w_prob = _get_workload_prob(data)
        s_risk = _sentiment_risk(feedback)
        score, level = compute_risk(w_prob, s_risk)
        reasons = get_top_reasons(workload_model, data, s_risk, top_n=3)

        if latest:
            latest.avg_daily_hours = avg_daily_hours
            latest.overtime_hours_week = overtime_hours_week
            latest.absence_days_month = absence_days_month
            latest.late_arrivals_month = late_arrivals_month
            latest.tasks_assigned = tasks_assigned
            latest.vacation_days_last_year = vacation_days_last_year
            latest.feedback = feedback
            latest.risk_score = score
            latest.level = level
            latest.workload_probability = round(w_prob, 4)
            latest.sentiment_risk = round(s_risk, 4)
            latest.top_reasons = json.dumps(reasons)
            latest.created_at = datetime.utcnow()
        else:
            new_assessment = Assessment(
                user_id=user.id,
                avg_daily_hours=avg_daily_hours,
                overtime_hours_week=overtime_hours_week,
                absence_days_month=absence_days_month,
                late_arrivals_month=late_arrivals_month,
                tasks_assigned=tasks_assigned,
                vacation_days_last_year=vacation_days_last_year,
                feedback=feedback,
                risk_score=score,
                level=level,
                workload_probability=round(w_prob, 4),
                sentiment_risk=round(s_risk, 4),
                top_reasons=json.dumps(reasons),
            )
            db.add(new_assessment)

    db.commit()
    db.refresh(user)

    latest = user.assessments[0] if user.assessments else None
    emp_id = f"EMP{str(user.id).zfill(4)}"
    return EmployeeSummary(
        user_id=user.id,
        employee_id=emp_id,
        name=user.full_name,
        username=user.username,
        gender=user.gender,
        age=user.age,
        avg_daily_hours=latest.avg_daily_hours if latest else None,
        overtime_hours_week=latest.overtime_hours_week if latest else None,
        absence_days_month=latest.absence_days_month if latest else None,
        late_arrivals_month=latest.late_arrivals_month if latest else None,
        tasks_assigned=latest.tasks_assigned if latest else None,
        vacation_days_last_year=latest.vacation_days_last_year if latest else None,
        feedback=latest.feedback if latest else "",
        risk_score=latest.risk_score if latest else None,
        level=latest.level if latest else "Unassessed",
        last_assessed=latest.created_at.strftime("%Y-%m-%d %H:%M") if latest else "لم يقيّم بعد / Unassessed",
    )


@app.delete("/employees/{user_id}", tags=["Admin"])
def delete_employee(user_id: int, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="الموظف غير موجود (Employee not found)")
    if user.role == "admin" or user.username == "admin":
        raise HTTPException(status_code=400, detail="لا يمكن حذف حساب الأدمن الأساسي للنظام")

    name = user.full_name
    db.delete(user)
    db.commit()
    return {"status": "success", "detail": f"تم حذف الموظف '{name}' وجميع سجلاته بنجاح"}


@app.get("/stats", response_model=StatsResponse, tags=["Admin"])
def get_stats(db: Session = Depends(get_db)):
    users = db.query(User).filter(User.role == "employee").all()
    total = len(users)
    if total == 0:
        return StatsResponse(
            total=0, low=0, medium=0, high=0,
            low_pct=0, medium_pct=0, high_pct=0,
            avg_daily_hours=0, avg_overtime_hours=0, avg_tasks=0, avg_risk_score=0,
            high_risk_list=[], low_metrics={}, med_metrics={}, high_metrics={}
        )

    low = medium = high = 0
    total_hours = 0.0
    total_ot = 0.0
    total_tasks = 0
    total_score = 0
    count_assessed = 0

    high_risk_list = []
    level_buckets = {
        "Low": {"hours": [], "ot": [], "tasks": []},
        "Medium": {"hours": [], "ot": [], "tasks": []},
        "High": {"hours": [], "ot": [], "tasks": []},
    }

    for u in users:
        if u.assessments:
            latest = u.assessments[0]
            lvl = latest.level
            if lvl == "Low":
                low += 1
            elif lvl == "Medium":
                medium += 1
            elif lvl == "High":
                high += 1

            total_hours += latest.avg_daily_hours
            total_ot += latest.overtime_hours_week
            total_tasks += latest.tasks_assigned
            total_score += latest.risk_score
            count_assessed += 1

            if lvl in level_buckets:
                level_buckets[lvl]["hours"].append(latest.avg_daily_hours)
                level_buckets[lvl]["ot"].append(latest.overtime_hours_week)
                level_buckets[lvl]["tasks"].append(latest.tasks_assigned)

            if lvl == "High":
                high_risk_list.append(HighRiskItem(
                    user_id=u.id,
                    employee_id=f"EMP{str(u.id).zfill(4)}",
                    name=u.full_name,
                    risk_score=latest.risk_score,
                    level=latest.level,
                    overtime_hours_week=latest.overtime_hours_week,
                    avg_daily_hours=latest.avg_daily_hours,
                    tasks_assigned=latest.tasks_assigned
                ))

    # sort high risk by risk_score desc
    high_risk_list.sort(key=lambda x: x.risk_score, reverse=True)

    denom = max(1, count_assessed)
    def avg_val(lst): return round(sum(lst)/len(lst), 1) if lst else 0.0

    return StatsResponse(
        total=total,
        low=low,
        medium=medium,
        high=high,
        low_pct=round(low / total * 100, 1) if total else 0,
        medium_pct=round(medium / total * 100, 1) if total else 0,
        high_pct=round(high / total * 100, 1) if total else 0,
        avg_daily_hours=round(total_hours / denom, 1),
        avg_overtime_hours=round(total_ot / denom, 1),
        avg_tasks=round(total_tasks / denom, 1),
        avg_risk_score=round(total_score / denom, 1),
        high_risk_list=high_risk_list[:5],
        low_metrics={"avg_hours": avg_val(level_buckets["Low"]["hours"]), "avg_ot": avg_val(level_buckets["Low"]["ot"]), "avg_tasks": avg_val(level_buckets["Low"]["tasks"])},
        med_metrics={"avg_hours": avg_val(level_buckets["Medium"]["hours"]), "avg_ot": avg_val(level_buckets["Medium"]["ot"]), "avg_tasks": avg_val(level_buckets["Medium"]["tasks"])},
        high_metrics={"avg_hours": avg_val(level_buckets["High"]["hours"]), "avg_ot": avg_val(level_buckets["High"]["ot"]), "avg_tasks": avg_val(level_buckets["High"]["tasks"])},
    )


@app.get("/health", tags=["System"])
def health_check():
    return {"status": "ok", "system": "Tawazon Enterprise", "db": "Connected"}
