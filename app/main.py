from datetime import datetime, timedelta
from typing import Optional, List
import os
from fastapi import FastAPI, Depends, HTTPException
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from jose import jwt, JWTError
from passlib.context import CryptContext
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime
from sqlalchemy.orm import sessionmaker, declarative_base, Session

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./talho.db")

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

pwd_context = CryptContext(schemes=["pbkdf2_sha256"])
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/login")

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True)
    username = Column(String, unique=True)
    password_hash = Column(String)
    role = Column(String)

class CashClose(Base):
    __tablename__ = "cash_closes"
    id = Column(Integer, primary_key=True)
    section = Column(String)
    real_total = Column(Float)
    diff_total = Column(Float)
    created_at = Column(DateTime, default=datetime.utcnow)

Base.metadata.create_all(bind=engine)

def seed():
    db = SessionLocal()
    if not db.query(User).first():
        db.add(User(username="admin", password_hash=pwd_context.hash("1234"), role="admin"))
        db.commit()
    db.close()

seed()

class Token(BaseModel):
    access_token: str
    token_type: str
    username: str

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def create_token(data):
    data["exp"] = datetime.utcnow() + timedelta(hours=12)
    return jwt.encode(data, "secret", algorithm="HS256")

def get_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    try:
        payload = jwt.decode(token, "secret", algorithms=["HS256"])
        username = payload.get("sub")
    except:
        raise HTTPException(401)
    user = db.query(User).filter(User.username == username).first()
    return user

app = FastAPI()

@app.post("/api/login")
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == form_data.username).first()
    if not user or not pwd_context.verify(form_data.password, user.password_hash):
        raise HTTPException(400, "Login inválido")
    return {
        "access_token": create_token({"sub": user.username}),
        "token_type": "bearer",
        "username": user.username
    }

@app.get("/api/dashboard")
def dashboard(user=Depends(get_user)):
    return {"ok": True}

@app.post("/api/cash-closes")
def save_close(section: str, real: float, diff: float, db: Session = Depends(get_db)):
    row = CashClose(section=section, real_total=real, diff_total=diff)
    db.add(row)
    db.commit()
    return {"ok": True}

app.mount("/", StaticFiles(directory="app/static", html=True), name="static")
