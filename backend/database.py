"""
database.py
===========
SQLAlchemy Database setup supporting PostgreSQL (and fallback SQLite for local demo).
Models:
    - User (admin or employee, with unique username and hashed password)
    - Assessment (history of assessments linked to an employee)
"""

import os
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, ForeignKey, Text
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from passlib.context import CryptContext

# ── Database URL Configuration ────────────────────────────────────────────────
# Supports PostgreSQL via environment variable (e.g. postgresql://user:password@localhost:5432/workpulse)
# Falls back to local sqlite file if PostgreSQL URL is not set.
DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    # Local fallback so the application runs immediately without requiring local Postgres daemon
    DATABASE_URL = "sqlite:///./workpulse.db"

# Handle heroku/render postgres prefix compatibility
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

connect_args = {"check_same_thread": False} if "sqlite" in DATABASE_URL else {}
engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

import bcrypt

# ── Password Utilities (Native bcrypt) ───────────────────────────────────────
def hash_password(password: str) -> str:
    pwd_bytes = password.encode("utf-8")[:72]
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(pwd_bytes, salt).decode("utf-8")

def verify_password(plain_password: str, hashed_password: str) -> bool:
    try:
        pwd_bytes = plain_password.encode("utf-8")[:72]
        hash_bytes = hashed_password.encode("utf-8")
        return bcrypt.checkpw(pwd_bytes, hash_bytes)
    except Exception:
        return False


# ── Database Models ───────────────────────────────────────────────────────────
class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, index=True, nullable=False)
    full_name = Column(String(100), nullable=False)
    gender = Column(String(20), nullable=True)  # "ذكر" / "أنثى" or "Male" / "Female"
    age = Column(Integer, nullable=True)
    role = Column(String(20), default="employee")  # "admin" or "employee"
    hashed_password = Column(String(255), nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    assessments = relationship("Assessment", back_populates="user", cascade="all, delete-orphan", order_by="desc(Assessment.created_at)")


class Assessment(Base):
    __tablename__ = "assessments"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    
    avg_daily_hours = Column(Float, nullable=False)
    overtime_hours_week = Column(Float, nullable=False)
    absence_days_month = Column(Integer, nullable=False)
    late_arrivals_month = Column(Integer, nullable=False)
    tasks_assigned = Column(Integer, nullable=False)
    vacation_days_last_year = Column(Integer, nullable=False)
    feedback = Column(Text, nullable=False)

    risk_score = Column(Integer, nullable=False)
    level = Column(String(20), nullable=False)  # "Low", "Medium", "High"
    workload_probability = Column(Float, nullable=False)
    sentiment_risk = Column(Float, nullable=False)
    top_reasons = Column(Text, nullable=False)  # JSON or comma-separated string
    created_at = Column(DateTime, default=datetime.utcnow)

    user = relationship("User", back_populates="assessments")


# ── DB Initialization & Seed ──────────────────────────────────────────────────
def init_db():
    Base.metadata.create_all(bind=engine)

    # Lightweight migration for new columns: gender & age
    from sqlalchemy import inspect, text
    with engine.connect() as conn:
        inspector = inspect(engine)
        columns = [c["name"] for c in inspector.get_columns("users")]
        if "gender" not in columns:
            try:
                conn.execute(text("ALTER TABLE users ADD COLUMN gender VARCHAR(20)"))
                conn.commit()
                print("[DB] Added 'gender' column to users table.")
            except Exception as e:
                print(f"[DB] Migration note (gender): {e}")
        if "age" not in columns:
            try:
                conn.execute(text("ALTER TABLE users ADD COLUMN age INTEGER"))
                conn.commit()
                print("[DB] Added 'age' column to users table.")
            except Exception as e:
                print(f"[DB] Migration note (age): {e}")

    db = SessionLocal()
    try:
        # Create Default Admin if not exists
        admin = db.query(User).filter(User.username == "admin").first()
        if not admin:
            admin_pwd = os.getenv("ADMIN_PASSWORD", "04072005")
            admin_user = User(
                username="admin",
                full_name="HR Administrator",
                role="admin",
                gender="Male",
                age=32,
                hashed_password=hash_password(admin_pwd)
            )
            db.add(admin_user)
            db.commit()
            print("[DB] Default Admin account created: admin")
    finally:
        db.close()
