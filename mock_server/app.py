"""In-memory mock subscriber and usage API, so the runner works anywhere.

All data is fictional and lives in memory, so restarting the server resets
it. Subscribers are keyed by phone number to keep spreadsheet rows simple.

Run it with::

    python -m uvicorn mock_server.app:app --port 8000

Business rules (so tests have something meaningful to check):
- plans have data/voice/SMS allowances; "unlimited" has none
- usage sends a notification when it first crosses 80% and 100%
- suspended or closed subscribers can't record usage
- a closed subscriber can't be changed
"""

from __future__ import annotations

import os
import secrets
from typing import Literal

from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel, Field

app = FastAPI(title="Mock Subscriber Usage API", version="1.0.0")
security = HTTPBasic()

API_USER = os.getenv("MOCK_API_USER", "demo")
API_PASSWORD = os.getenv("MOCK_API_PASSWORD", "demo-password")

UsageType = Literal["data_mb", "voice_min", "sms"]

# Allowance per billing cycle. None = unlimited.
PLANS: dict[str, dict[str, int | None]] = {
    "basic": {"data_mb": 1000, "voice_min": 300, "sms": 300},
    "plus": {"data_mb": 5000, "voice_min": 1000, "sms": 1000},
    "unlimited": {"data_mb": None, "voice_min": None, "sms": None},
}
THRESHOLDS = (80, 100)  # percent of allowance that triggers a notification

LABELS = {"data_mb": "Data", "voice_min": "Voice", "sms": "SMS"}


class CreateBody(BaseModel):
    phone_number: str = Field(pattern=r"^\d{10}$")
    plan: Literal["basic", "plus", "unlimited"]
    features: list[str] = []


class FeatureBody(BaseModel):
    add: list[str] = []
    remove: list[str] = []


class UsageBody(BaseModel):
    type: UsageType
    amount: int = Field(gt=0)


class Subscriber(BaseModel):
    phone_number: str
    plan: str
    status: Literal["active", "suspended", "closed"] = "active"
    features: list[str] = []
    used: dict[str, int] = Field(default_factory=lambda: {"data_mb": 0, "voice_min": 0, "sms": 0})
    # Thresholds already notified per usage type, so each fires only once.
    notified: dict[str, list[int]] = Field(default_factory=lambda: {"data_mb": [], "voice_min": [], "sms": []})


_db: dict[str, Subscriber] = {}


# ---- Helpers -------------------------------------------------------------

def require_auth(credentials: HTTPBasicCredentials = Depends(security)) -> None:
    """Reject requests without valid basic auth (constant-time comparison)."""
    ok = secrets.compare_digest(credentials.username, API_USER) and \
        secrets.compare_digest(credentials.password, API_PASSWORD)
    if not ok:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid credentials")


def _get(phone: str) -> Subscriber:
    sub = _db.get(phone)
    if sub is None:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Subscriber not found")
    return sub


def _require_active(sub: Subscriber) -> None:
    if sub.status != "active":
        raise HTTPException(status.HTTP_409_CONFLICT, f"Subscriber is {sub.status}")


def _view(sub: Subscriber, notifications: list[str] | None = None) -> dict:
    """Public representation: remaining allowance instead of internal counters."""
    remaining: dict[str, int | str] = {}
    for kind, allowance in PLANS[sub.plan].items():
        remaining[kind] = "unlimited" if allowance is None else max(allowance - sub.used[kind], 0)
    return {
        "phone_number": sub.phone_number,
        "plan": sub.plan,
        "status": sub.status,
        "features": sub.features,
        "remaining": remaining,
        "notifications": notifications or [],
    }


# ---- Endpoints -----------------------------------------------------------

@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/api/v1/subscribers", status_code=201, dependencies=[Depends(require_auth)])
def create(body: CreateBody) -> dict:
    existing = _db.get(body.phone_number)
    if existing and existing.status != "closed":
        raise HTTPException(status.HTTP_409_CONFLICT, "Phone number already in use")
    sub = Subscriber(phone_number=body.phone_number, plan=body.plan, features=body.features)
    _db[sub.phone_number] = sub
    return _view(sub)


@app.get("/api/v1/subscribers/{phone}", dependencies=[Depends(require_auth)])
def get(phone: str) -> dict:
    return _view(_get(phone))


@app.post("/api/v1/subscribers/{phone}/suspend", dependencies=[Depends(require_auth)])
def suspend(phone: str) -> dict:
    sub = _get(phone)
    _require_active(sub)
    sub.status = "suspended"
    return _view(sub)


@app.post("/api/v1/subscribers/{phone}/restore", dependencies=[Depends(require_auth)])
def restore(phone: str) -> dict:
    sub = _get(phone)
    if sub.status != "suspended":
        raise HTTPException(status.HTTP_409_CONFLICT, f"Subscriber is {sub.status}")
    sub.status = "active"
    return _view(sub)


@app.post("/api/v1/subscribers/{phone}/features", dependencies=[Depends(require_auth)])
def features(phone: str, body: FeatureBody) -> dict:
    sub = _get(phone)
    if sub.status == "closed":
        raise HTTPException(status.HTTP_409_CONFLICT, "Subscriber is closed")
    current = [f for f in sub.features if f not in body.remove]
    current += [f for f in body.add if f not in current]
    sub.features = current
    return _view(sub)


@app.post("/api/v1/subscribers/{phone}/usage", dependencies=[Depends(require_auth)])
def record_usage(phone: str, body: UsageBody) -> dict:
    """Record usage and return any notifications triggered by this event."""
    sub = _get(phone)
    _require_active(sub)
    sub.used[body.type] += body.amount

    notifications: list[str] = []
    allowance = PLANS[sub.plan][body.type]
    if allowance:
        percent = sub.used[body.type] * 100 // allowance
        for threshold in THRESHOLDS:
            if percent >= threshold and threshold not in sub.notified[body.type]:
                sub.notified[body.type].append(threshold)
                label = LABELS[body.type]
                notifications.append(
                    f"{label} allowance exhausted" if threshold == 100
                    else f"{label} usage at {threshold}% of allowance"
                )
    return _view(sub, notifications)


@app.post("/api/v1/subscribers/{phone}/close", dependencies=[Depends(require_auth)])
def close(phone: str) -> dict:
    sub = _get(phone)
    if sub.status == "closed":
        raise HTTPException(status.HTTP_409_CONFLICT, "Subscriber is closed")
    sub.status = "closed"
    return _view(sub)


@app.post("/api/v1/_reset", include_in_schema=False, dependencies=[Depends(require_auth)])
def reset() -> dict:
    """Test helper: clear all data."""
    _db.clear()
    return {"reset": True}
