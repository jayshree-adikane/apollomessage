import os
from dotenv import load_dotenv

# Secrets and credentials come from the local .env file (see .env.example).
load_dotenv()

# 
from fastapi import FastAPI, Form, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from pymongo import MongoClient, ReturnDocument
import requests
import time

from datetime import datetime, timedelta, timezone
from datetime import time as dt_time
from apscheduler.schedulers.background import BackgroundScheduler
from msal import PublicClientApplication, SerializableTokenCache
import os
import smtplib
import ssl
import csv
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import make_msgid
from email.utils import formatdate
import http.client
import json
import asyncio
import threading
import uuid
import re
from urllib.parse import urlparse
from urllib.parse import urlparse, urlunparse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from email.mime.image import MIMEImage
from google import genai
import traceback
from html import escape
from zoneinfo import ZoneInfo
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import logging
from logging.handlers import RotatingFileHandler
from starlette.middleware.sessions import SessionMiddleware

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

logger = logging.getLogger("email_campaign")
logger.setLevel(logging.DEBUG)

file_handler = RotatingFileHandler(
    LOG_DIR / "email_campaign.log",
    maxBytes=10 * 1024 * 1024,
    backupCount=5,
    encoding="utf-8"
)

formatter = logging.Formatter(
    "%(asctime)s | %(levelname)s | %(message)s"
)

file_handler.setFormatter(formatter)

if not logger.handlers:
    logger.addHandler(file_handler)
    logger.addHandler(logging.StreamHandler())

openapi_tags = [
    {
        "name": "Email Campaign",
        "description": "Send or queue initial campaign emails."
    },
    {
        "name": "Queue",
        "description": "Inspect queued initial emails and reply-triggered follow-ups."
    },
    {
        "name": "Email Thread",
        "description": "Read the stored conversation thread for a prospect."
    },
    {
        "name": "Prospects",
        "description": "Search and retrieve prospect records through the application."
    }
]

app = FastAPI(
    title="Apollo Prospect Finder API",
    description=(
        "Swagger/OpenAPI documentation for the Apollo Prospect Finder email campaign "
        "application. Initial emails are restricted to the configured weekday sending "
        "window. Follow-up #1-#5 are scheduled 5 minutes after the previous "
        "outbound email when there is no customer reply."
        "follow-up, subject to the valid sending window."
    ),
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
    openapi_tags=openapi_tags,
    swagger_ui_parameters={
        "docExpansion": "none",
        "displayRequestDuration": True,
        "persistAuthorization": True
    }
)

APP_USERNAME = os.getenv("APP_USERNAME", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
SESSION_SECRET = os.getenv("SESSION_SECRET", "")




app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")


@app.middleware("http")
async def require_login(request: Request, call_next):
    """Require login for the application while leaving login/static public."""
    public_paths = {
        "/login",
        "/logout",
        "/favicon.ico",
        "/docs",
        "/redoc",
        "/openapi.json"
    }

    if request.url.path in public_paths or request.url.path.startswith("/static/"):
        return await call_next(request)

    if request.session.get("authenticated") is True:
        return await call_next(request)

    return RedirectResponse(url="/login", status_code=303)

app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    session_cookie="apollo_login_session",
    max_age=60 * 60 * 8,
    same_site="lax",
    https_only=False,
)


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    if request.session.get("authenticated") is True:
        return RedirectResponse(url="/", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"error": ""}
    )


@app.post("/login", response_class=HTMLResponse)
async def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    if username.strip() == APP_USERNAME and password == APP_PASSWORD:
        request.session.clear()
        request.session["authenticated"] = True
        request.session["username"] = APP_USERNAME
        return RedirectResponse(url="/", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"error": "Invalid username or password."},
        status_code=401
    )


@app.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse(url="/login", status_code=303)



#Apollo ID
API_KEY = os.getenv("APOLLO_API_KEY", "")

headers = {
    "X-Api-Key": API_KEY,
    "Content-Type": "application/json"
}


# ============================================================
# APOLLO CREDIT USAGE
# ============================================================



APOLLO_CREDIT_USAGE_URL = (
    "https://api.apollo.io/api/v1/usage_stats/credit_usage_stats"
)



def get_apollo_credit_usage():
    try:
        response = requests.post(
            "https://api.apollo.io/api/v1/usage_stats/credit_usage_stats",
            headers={
                "x-api-key": API_KEY,
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            json={},
            timeout=20,
        )

        print("STATUS:", response.status_code)
        print("RESPONSE:fkjkjkgjfdkgkfd", response.text)

        if response.status_code != 200:
            return {
                "success": False,
                "error": f"Apollo returned HTTP {response.status_code}",
                "credit_usage_stats": {},
                "current_credit_cycle": {},
            }

        data = response.json()

        return {
            "success": True,
            "credit_usage_stats": data.get(
                "credit_usage_stats",
                {}
            ),
            "current_credit_cycle": data.get(
                "current_credit_cycle",
                {}
            ),
        }

    except requests.RequestException as exc:
        return {
            "success": False,
            "error": str(exc),
            "credit_usage_stats": {},
            "current_credit_cycle": {},
        }


# ============================================================
# APOLLO CREDITS API
# ============================================================

@app.get("/apollo-credits")
def apollo_credits():

    credits = get_apollo_credit_usage()

    if not credits["success"]:
        return credits

    return credits

# # MongoDB local
# client = MongoClient("mongodb://localhost:27017")
# db = client["Apollo_Email"]
# collection = db["prospects"]
# blacklist_collection = db["Blacklist"]

# MongoDB online
client = MongoClient(os.getenv("MONGODB_URI", ""))
db = client["Apollo_Email"]
collection = db["prospects"]
blacklist_collection = db["Blacklist"]

# Prevent duplicate prospect records with the same email
try:
    collection.create_index(
        [("email", 1)],
        unique=True,
        sparse=True,
        name="unique_prospect_email"
    )
except Exception as e:
    print("Unique email index warning:", e)


# SMTP Configuration
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SENDER_EMAIL = os.getenv("SMTP_SENDER_EMAIL", "")
PASSWORD = os.getenv("SMTP_PASSWORD", "")

# Microsoft Graph

CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
TENANT_ID = os.getenv("MS_TENANT_ID", "")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "User.Read",
    "Mail.Read",
    "Chat.Read",
    "Chat.ReadWrite",
    "ChatMessage.Send"
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, "token_cache.bin")

cache = SerializableTokenCache()

if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        cache.deserialize(f.read())

msal_app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    token_cache=cache
)

# Microsoft Teams notification configuration.
# The notification is posted to the signed-in user's own Teams self-chat
# (the chat displayed by Teams as e.g. "Rameshwar Thakur (You)").
TEAMS_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
TEAMS_NOTIFICATION_EMAIL = SENDER_EMAIL

SERPER_API_KEY = os.getenv("SERPER_API_KEY", "")

gemini_client = genai.Client(
    api_key=os.getenv("GEMINI_API_KEY", "")
)

def get_signature():

    with open(
        "templates/signature.html",
        "r",
        encoding="utf-8"
    ) as f:

        return f.read()

DOMAIN_REPLACEMENTS = {
    "news.microsoft.com": "microsoft.com",
    "goo.gle": "google.com",
    "aboutamazon.com": "amazon.com"
}


IST = ZoneInfo("Asia/Kolkata")
EMAIL_START_TIME = dt_time(9, 0)    # 9:00 AM
EMAIL_END_TIME = dt_time(17, 0)      # 5:00 PM
EMAIL_INTERVAL_SECONDS = 30

# ============================================================
# FOLLOW-UP CADENCE
# ============================================================
# Follow-up #1: 48 hours after the initial email
# Follow-up #2: 48 hours after Follow-up #1
# Follow-up #3: 48 hours after Follow-up #2
# Follow-up #4: 120 hours after Follow-up #3
# Follow-up #5: 120 hours after Follow-up #4
FOLLOWUP_INTERVALS = {
    1: timedelta(hours=48),
    2: timedelta(hours=48),
    3: timedelta(hours=48),
    4: timedelta(hours=120),
    5: timedelta(hours=120),
}

def _email_window_start_datetime(date_value):
    """Return the configured weekday email-window start for a calendar date."""
    return datetime.combine(
        date_value,
        EMAIL_START_TIME,
        tzinfo=IST
    )


def is_email_sending_window_at(value):
    """Return True only during Monday-Friday 09:00-17:00 IST."""
    candidate = _coerce_ist_datetime(value)
    if candidate is None:
        return False

    return (
        candidate.weekday() < 5
        and EMAIL_START_TIME <= candidate.time() < EMAIL_END_TIME
    )


def is_email_sending_window():
    """Return True only during the configured Monday-Friday email window."""
    return is_email_sending_window_at(datetime.now(IST))


def get_next_email_window_start():
    """Return the next valid Monday-Friday 09:00 IST email-window start."""
    now = datetime.now(IST)

    # Already inside the sending window.
    if is_email_sending_window_at(now):
        return now.replace(microsecond=0)

    # Before 09:00 on a weekday -> start today at 09:00.
    if now.weekday() < 5 and now.time() < EMAIL_START_TIME:
        return _email_window_start_datetime(now.date())

    # At/after 17:00 on a weekday, or any weekend -> next weekday at 09:00.
    candidate_date = now.date() + timedelta(days=1)
    while candidate_date.weekday() >= 5:
        candidate_date += timedelta(days=1)

    return _email_window_start_datetime(candidate_date)

def _get_followup_interval(followup_number):
    """Return the configured interval for Follow-up #1-#5."""
    try:
        number = int(followup_number)
    except (TypeError, ValueError):
        raise ValueError("followup_number must be an integer from 1 to 5")

    if number not in FOLLOWUP_INTERVALS:
        raise ValueError(f"Unsupported followup number: {number}")

    return FOLLOWUP_INTERVALS[number]


def get_first_followup_send_time(base_dt):
    """Return Follow-up #1 48 hours after the initial email."""
    return get_next_followup_send_time(base_dt, 1)


def repair_existing_first_followup_slots():
    """Recalculate pending Follow-up #1 schedules from the initial email date.

    A customer reply cancels the follow-up sequence, so only no-reply records
    are eligible for this repair.
    """
    repaired = 0

    records = collection.find({
        "has_reply": False,
        "email_sent": True,
        "followup_count": 0,
        "email_sent_date": {"$exists": True, "$nin": ["", None]}
    }, {
        "_id": 1,
        "email_sent_date": 1,
        "followup_queue_status": 1,
        "followup_queued_for": 1
    })

    for record in records:
        try:
            new_slot = get_first_followup_send_time(record["email_sent_date"])
            new_text = new_slot.strftime("%Y-%m-%d %H:%M:%S")
            old_text = str(record.get("followup_queued_for") or "")

            if (
                old_text != new_text
                or record.get("followup_queue_status") != "scheduled"
            ):
                collection.update_one(
                    {"_id": record["_id"]},
                    {
                        "$set": {
                            "followup_due_at": new_text,
                            "followup_queued_for": new_text,
                            "followup_queue_status": "scheduled",
                            "followup_queue_number": 1,
                            "sending_followup": False
                        }
                    }
                )
                repaired += 1
        except Exception as exc:
            logger.warning("Unable to repair Follow-up #1 slot for %s: %s", record.get("_id"), exc)

    if repaired:
        logger.info("Repaired %s pending Follow-up #1 schedule(s) from the initial-email date.", repaired)

    return repaired


def repair_followups_after_replies():
    """Cancel any pending follow-up sequence for prospects who have replied."""
    result = collection.update_many(
        {
            "has_reply": True,
            "followup_queue_status": "scheduled",
            "followup_count": {"$lt": 5}
        },
        {
            "$set": {
                "campaign_status": "Replied",
                "followup_queue_status": "cancelled",
                "followup_queue_number": 0,
                "followup_queued_for": "",
                "followup_due_at": "",
                "sending_followup": False,
                "active": False
            }
        }
    )

    if result.modified_count:
        logger.info(
            "Cancelled %s pending follow-up sequence(s) because the prospect already replied.",
            result.modified_count
        )

    return result.modified_count


def get_next_followup_send_time(base_dt, followup_number=2):
    """Return the next follow-up time using the configured #1-#5 cadence, normalized into the email window."""
    if isinstance(base_dt, str):
        text = base_dt.strip()
        try:
            base_dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            base_dt = datetime.strptime(text, "%Y-%m-%d %H:%M:%S")

    if not isinstance(base_dt, datetime):
        raise TypeError("base_dt must be a datetime or date string")

    if base_dt.tzinfo is None:
        base_dt = base_dt.replace(tzinfo=IST)
    else:
        base_dt = base_dt.astimezone(IST)

    interval = _get_followup_interval(followup_number)
    candidate = base_dt + interval

    # Follow-ups may be due after hours, before 09:00, or on weekends.
    # Never send outside the configured Monday-Friday 09:00-17:00 IST window.
    return normalize_to_email_window(candidate)


def repair_completed_followup_statuses():
    """Keep legacy records with 5 completed follow-ups marked as completed.

    This fixes older records whose database status was not updated after
    the fifth follow-up was successfully sent.
    """
    result = collection.update_many(
        {
            "followup_count": {"$gte": 5},
            "campaign_status": {"$ne": "Completed - Follow-ups Sent"}
        },
        {
            "$set": {
                "campaign_status": "Completed - Follow-ups Sent",
                "followup_queue_status": "completed",
                "followup_queue_number": 5,
                "followup_queued_for": "",
                "followup_due_at": "",
                "followup_sent": True,
                "active": False,
                "sending_followup": False
            }
        }
    )

    if result.modified_count:
        logger.info(
            "Marked %s record(s) with 5+ follow-ups as Completed - Follow-ups Sent.",
            result.modified_count
        )

    return result.modified_count


def repair_existing_followup_slots():
    """Repair pending Follow-up #2-#5 schedules using the configured cadence.
    Follow-up #2-#3 are 48 hours after the previous follow-up; Follow-up #4-#5 are 120 hours.
    Follow-up #1 is handled separately.
    """
    repaired = 0

    records = collection.find(
        {
            "has_reply": False,
            "email_sent": True,
            "followup_count": {"$gte": 1, "$lt": 5},
            "followup_queue_status": "scheduled",
            "last_followup_date": {"$exists": True, "$nin": ["", None]},
            "followup_queued_for": {"$exists": True, "$nin": ["", None]}
        },
        {
            "_id": 1,
            "followup_count": 1,
            "last_followup_date": 1,
            "followup_queued_for": 1
        }
    )

    for record in records:
        try:
            next_followup_number = record.get("followup_count", 0) + 1
            new_slot = get_next_followup_send_time(
                record["last_followup_date"],
                next_followup_number
            )
            new_text = new_slot.strftime("%Y-%m-%d %H:%M:%S")
            old_text = str(record.get("followup_queued_for") or "")

            if old_text != new_text:
                collection.update_one(
                    {"_id": record["_id"]},
                    {
                        "$set": {
                            "followup_due_at": new_text,
                            "followup_queued_for": new_text,
                            "followup_queue_status": "scheduled",
                            "followup_queue_number": record.get("followup_count", 0) + 1,
                            "sending_followup": False
                        }
                    }
                )
                repaired += 1
        except Exception as exc:
            logger.warning(
                "Unable to repair Follow-up #%s slot for %s: %s",
                (record.get("followup_count", 0) + 1),
                record.get("_id"),
                exc
            )

    if repaired:
        logger.info(
            "Repaired %s pending Follow-up #2-#5 schedule(s) to the 48/48/120/120-hour cadence.",
            repaired
        )

    return repaired


def get_next_followup_queue_slot(base_dt, followup_number=2):
    """Return the next valid follow-up queue slot with a 30-second minimum gap.

    If the current day's 09:00-17:00 window is full, the slot moves to the
    next Monday-Friday 09:00 window, skipping Saturday and Sunday.
    """
    candidate = get_next_followup_send_time(base_dt, followup_number)

    while True:
        date_prefix = candidate.strftime("%Y-%m-%d")

        latest = collection.find_one(
            {
                "followup_queue_status": "scheduled",
                "followup_queued_for": {
                    "$regex": f"^{date_prefix} "
                }
            },
            {"followup_queued_for": 1},
            sort=[("followup_queued_for", -1)]
        )

        if not latest or not latest.get("followup_queued_for"):
            return candidate

        latest_slot = _coerce_ist_datetime(latest.get("followup_queued_for"))
        if not latest_slot or latest_slot < candidate:
            return candidate

        candidate = normalize_to_email_window(
            latest_slot + timedelta(seconds=EMAIL_INTERVAL_SECONDS)
        )


def get_next_retry_time():
    """Return the next Monday-Friday 09:00 IST time for a failed initial email."""
    return get_next_email_window_start()


def retryable_initial_email_filter():
    """Mongo condition: new prospects or retryable delivery failures whose retry time has arrived."""
    now_string = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    return {
        "$and": [
            {
                "$or": [
                    {"email_sent": {"$exists": False}},
                    {"email_sent": False},
                    {"email_sent": None}
                ]
            },
            {
                "$or": [
                    {"email_retry_after": {"$exists": False}},
                    {"email_retry_after": None},
                    {"email_retry_after": {"$lte": now_string}}
                ]
            },
            # Permanent undeliverable addresses must stay visible for review
            # but must not be automatically retried.
            {
                "email_failure_type": {
                    "$ne": "undeliverable"
                }
            }
        ]
    }


def prospect_result_filter():
    """Mongo filter for the main Prospect Results table.

    Failed/undeliverable initial emails stay visible immediately so the user
    can see the record again and manually retry it from the main page.
    """
    return {
        "$and": [
            {
                "$or": [
                    {"email_sent": {"$exists": False}},
                    {"email_sent": False},
                    {"email_sent": None}
                ]
            },
            {
                "email_queue_status": {"$ne": "queued"}
            },
            {
                "email": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            },
            {
                "domain": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            }
        ]
    }


def _coerce_ist_datetime(value):
    """Convert a datetime/string value into a timezone-aware IST datetime."""
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=IST)
        return value.astimezone(IST)

    text = str(value or "").strip()
    if not text:
        return None

    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        parsed = datetime.strptime(text, "%Y-%m-%d %H:%M:%S")

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=IST)
    else:
        parsed = parsed.astimezone(IST)

    return parsed


def normalize_to_email_window(candidate):
    """Move a datetime into the next valid Monday-Friday 09:00-17:00 IST window."""
    candidate = _coerce_ist_datetime(candidate)
    if candidate is None:
        candidate = get_next_email_window_start()

    # Already inside the configured window.
    if is_email_sending_window_at(candidate):
        return candidate

    # Before 09:00 on a weekday -> today at 09:00.
    if candidate.weekday() < 5 and candidate.time() < EMAIL_START_TIME:
        return _email_window_start_datetime(candidate.date())

    # At/after 17:00 on a weekday, or any weekend -> next weekday at 09:00.
    candidate_date = candidate.date() + timedelta(days=1)
    while candidate_date.weekday() >= 5:
        candidate_date += timedelta(days=1)

    return _email_window_start_datetime(candidate_date)


def get_next_initial_queue_slot(base_window):
    """Return the next free initial-email queue slot with a 30-second gap.

    Queue slots stay inside Monday-Friday 09:00-17:00 IST. Once the current
    day's 17:00 limit is reached, the next slot moves to the next weekday at
    09:00, skipping Saturday and Sunday.
    """
    candidate = normalize_to_email_window(base_window)

    while True:
        date_prefix = candidate.strftime("%Y-%m-%d")

        latest = collection.find_one(
            {
                "email_queue_status": "queued",
                "email_sent": {"$ne": True},
                "email_queued_for": {
                    "$regex": f"^{date_prefix} "
                }
            },
            {"email_queued_for": 1},
            sort=[("email_queued_for", -1)]
        )

        if not latest or not latest.get("email_queued_for"):
            return candidate

        latest_slot = _coerce_ist_datetime(latest.get("email_queued_for"))
        if not latest_slot or latest_slot < candidate:
            return candidate

        candidate = normalize_to_email_window(
            latest_slot + timedelta(seconds=EMAIL_INTERVAL_SECONDS)
        )


def _queue_slot_is_valid(value):
    """Return True when a stored initial queue time is in the active window rules."""
    return is_email_sending_window_at(value)


def repair_existing_initial_queue_slots():
    """Repair pending initial queue records into the Monday-Friday 09:00-17:00 window."""
    queued_records = list(
        collection.find(
            {
                "email_queue_status": "queued",
                "email_sent": {"$ne": True},
                "email_queued_for": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            },
            {
                "_id": 1,
                "email_queued_for": 1,
                "email_queued_at": 1
            }
        ).sort("email_queued_at", 1)
    )

    if not queued_records:
        return 0

    # Only rebuild when at least one pending record still has a time outside
    # the configured Monday-Friday 09:00-17:00 window.
    if all(
        _queue_slot_is_valid(_coerce_ist_datetime(record.get("email_queued_for")))
        for record in queued_records
    ):
        return 0

    now = datetime.now(IST)
    if is_email_sending_window_at(now):
        slot = now.replace(microsecond=0)
    else:
        slot = get_next_email_window_start()

    repaired = 0

    for record in queued_records:
        slot = normalize_to_email_window(slot)
        queued_for = slot.strftime("%Y-%m-%d %H:%M:%S")

        collection.update_one(
            {"_id": record["_id"]},
            {
                "$set": {
                    "email_queued_for": queued_for,
                    "email_queue_reason": "weekday_daytime_window_repaired"
                }
            }
        )

        repaired += 1
        slot += timedelta(seconds=EMAIL_INTERVAL_SECONDS)

    logger.info(
        "Repaired %s pending initial queue record(s) into the %s-%s IST weekday window with %ss spacing.",
        repaired,
        EMAIL_START_TIME.strftime("%H:%M"),
        EMAIL_END_TIME.strftime("%H:%M"),
        EMAIL_INTERVAL_SECONDS
    )
    return repaired


def queue_initial_email(
    person_id,
    receiver,
    subject,
    body,
    signature_html,
    next_window,
    reason="outside_email_window"
):
    """
    Persist an initial email using the next available weekday daytime slot.

    Queue slots are always Monday-Friday 09:00-17:00 IST with the configured
    30-second spacing. If the current day's window is exhausted, the slot
    automatically rolls to the next weekday at 09:00.
    """
    queued_at = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    queued_datetime = get_next_initial_queue_slot(next_window)
    queued_for = queued_datetime.strftime("%Y-%m-%d %H:%M:%S")

    collection.update_one(
        {"_id": person_id},
        {
            "$set": {
                "sending_initial": False,
                "email_sending": False,
                "email_queue_status": "queued",
                "email_queued_for": queued_for,
                "email_queued_at": queued_at,
                "email_queue_reason": reason,
                "email_queue_subject": subject,
                "email_queue_body": body,
                "email_queue_signature": signature_html,
                "email_delivery_status": "queued"
            }
        }
    )

    return queued_for




def wait_for_initial_email_gap():
    """Ensure at least EMAIL_INTERVAL_SECONDS since the previous initial email send."""
    latest = collection.find_one(
        {
            "email_sent": True,
            "email_sent_date": {
                "$exists": True,
                "$nin": ["", None]
            }
        },
        {"email_sent_date": 1},
        sort=[("email_sent_date", -1)]
    )

    if not latest or not latest.get("email_sent_date"):
        return

    last_sent = _coerce_ist_datetime(latest.get("email_sent_date"))
    if not last_sent:
        return

    elapsed = (datetime.now(IST) - last_sent).total_seconds()
    remaining = EMAIL_INTERVAL_SECONDS - elapsed

    if remaining > 0:
        logger.info(
            "Waiting %.1f seconds to preserve the %s-second initial-email gap.",
            remaining,
            EMAIL_INTERVAL_SECONDS
        )
        time.sleep(remaining)


def wait_for_outbound_email_gap():
    """Ensure at least EMAIL_INTERVAL_SECONDS between any two outbound emails."""
    latest_initial = collection.find_one(
        {
            "email_sent": True,
            "email_sent_date": {"$exists": True, "$nin": ["", None]}
        },
        {"email_sent_date": 1},
        sort=[("email_sent_date", -1)]
    )

    latest_followup = collection.find_one(
        {
            "followup_sent": True,
            "followup_sent_date": {"$exists": True, "$nin": ["", None]}
        },
        {"followup_sent_date": 1},
        sort=[("followup_sent_date", -1)]
    )

    timestamps = []

    if latest_initial and latest_initial.get("email_sent_date"):
        parsed = _coerce_ist_datetime(latest_initial.get("email_sent_date"))
        if parsed:
            timestamps.append(parsed)

    if latest_followup and latest_followup.get("followup_sent_date"):
        parsed = _coerce_ist_datetime(latest_followup.get("followup_sent_date"))
        if parsed:
            timestamps.append(parsed)

    if not timestamps:
        return

    last_sent = max(timestamps)
    elapsed = (datetime.now(IST) - last_sent).total_seconds()
    remaining = EMAIL_INTERVAL_SECONDS - elapsed

    if remaining > 0:
        logger.info(
            "Waiting %.1f seconds to preserve the %s-second outbound-email gap.",
            remaining,
            EMAIL_INTERVAL_SECONDS
        )
        time.sleep(remaining)


def clear_initial_email_queue(person_id):
    collection.update_one(
        {"_id": person_id},
        {
            "$unset": {
                "email_queue_status": "",
                "email_queued_for": "",
                "email_queued_at": "",
                "email_queue_reason": "",
                "email_queue_subject": "",
                "email_queue_body": "",
                "email_queue_signature": ""
            }
        }
    )


def build_initial_email_message(receiver, subject, body):
    """Build the same HTML email message used by the bulk sender."""
    message = MIMEMultipart("related")
    alternative = MIMEMultipart("alternative")

    alternative.attach(
        MIMEText(body, "html", "utf-8")
    )
    message.attach(alternative)

    message["From"] = SENDER_EMAIL
    message["To"] = receiver
    message["Subject"] = subject

    message_id = make_msgid()
    message["Message-ID"] = message_id
    message["Date"] = formatdate(localtime=True)

    images = {
        "logo": "static/logo.png",
        "regards": "static/regards.png",
        "awards": "static/awards.png",
        "facebook": "static/facebook.png",
        "instagram": "static/instagram.png",
        "linkedin": "static/linkedin.png",
        "youtube": "static/youtube.png",
        "pinterest": "static/pinterest.png",
        "x": "static/x.png"
    }

    for cid, image_path in images.items():
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Image not found: {image_path}")

        with open(image_path, "rb") as f:
            img = MIMEImage(f.read())

        img.add_header("Content-ID", f"<{cid}>")
        img.add_header(
            "Content-Disposition",
            "inline",
            filename=os.path.basename(image_path)
        )
        message.attach(img)

    return message, message_id.strip("<>")


def send_queued_initial_emails():
    """
    Send queued initial emails only Monday-Friday, 09:00-17:00 IST.

    Any queued email that cannot be sent before the current window closes
    remains queued. Its queue slot is normalized to the next valid weekday
    at 09:00, so Saturday/Sunday are always skipped.
    """
    if not is_email_sending_window():
        return

    queued = list(
        collection.find(
            {
                "email_queue_status": "queued",
                "email_sent": {"$ne": True},
                "$or": [
                    {
                        "email_queued_for": {
                            "$lte": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                        }
                    },
                    {"email_queued_for": {"$exists": False}},
                    {"email_queued_for": None},
                    {"email_queued_for": ""}
                ]
            }
        ).sort("email_queued_for", 1)
    )

    if not queued:
        return

    print(
        f"📬 QUEUED EMAIL WORKER: {len(queued)} email(s) ready "
        f"at {datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S')}"
    )

    server = None

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60)
        server.starttls(context=ssl.create_default_context())
        server.login(SENDER_EMAIL, PASSWORD)

        for index, person in enumerate(queued):
            if not is_email_sending_window():
                print("⏸ QUEUED EMAIL WORKER: window closed; remaining emails stay queued.")
                break

            receiver = (person.get("email") or "").strip().lower()
            subject = person.get("email_queue_subject", "")
            body = person.get("email_queue_body", "")
            signature_html = person.get("email_queue_signature", "")

            if not receiver or not subject or not body:
                clear_initial_email_queue(person["_id"])
                continue

            # Atomic claim so two workers/processes do not send the same queued email.
            claim = collection.update_one(
                {
                    "_id": person["_id"],
                    "email_sent": {"$ne": True},
                    "email_queue_status": "queued",
                    "email_sending": {"$ne": True}
                },
                {
                    "$set": {
                        "email_sending": True,
                        "sending_initial": True
                    }
                }
            )

            if claim.modified_count == 0:
                continue

            try:
                message, clean_message_id = build_initial_email_message(
                    receiver,
                    subject,
                    body
                )

                # Preserve a real 30-second gap even when the 1-minute scheduler
                # releases multiple queued records at the same time.
                wait_for_outbound_email_gap()

                if not is_email_sending_window():
                    collection.update_one(
                        {"_id": person["_id"]},
                        {"$set": {"email_sending": False, "sending_initial": False}}
                    )
                    break

                print(f"📤 Sending queued email to {receiver}")
                queued_send_result = server.sendmail(
                    SENDER_EMAIL,
                    [receiver],
                    message.as_string()
                )

                # A non-empty return value means one or more recipients were
                # refused by SMTP. Do not mark the queued email as sent.
                if queued_send_result:
                    raise smtplib.SMTPRecipientsRefused(queued_send_result)

                sent_time = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                first_followup_due_at = (
                    get_first_followup_send_time(sent_time)
                    .strftime("%Y-%m-%d %H:%M:%S")
                )

                collection.update_one(
                    {"_id": person["_id"]},
                    {
                        "$set": {
                            "email_sent": True,
                            "email_sending": False,
                            "sending_initial": False,
                            "email_sent_date": sent_time,
                            "last_outbound_sent_date": sent_time,
                            "email_delivery_status": "sent",
                            "last_email_subject": subject,
                            "last_email_body": body,
                            "email_signature": signature_html,
                            "message_id": clean_message_id,
                            "thread_message_id": clean_message_id,
                            "last_message_id": clean_message_id,
                            "thread_type": "Same Thread",
                            "followup_count": 0,
                            "followup_sent": False,
                            "followup_sent_date": "",
                            "last_followup_date": "",
                            "followup_due_at": first_followup_due_at,
                            "followup_queue_status": "scheduled",
                            "followup_queue_number": 1,
                            "followup_queued_for": first_followup_due_at,
                            "campaign_status": "Running",
                            "has_reply": False,
                            "active": True,
                            "followup_cancel_reason": "",
                            "reply_subject": "",
                            "reply_received_date": "",
                            "email_error": "",
                            "email_thread": [{
                                "type": "initial",
                                "number": 0,
                                "subject": subject,
                                "body": body,
                                "signature": signature_html,
                                "message_id": clean_message_id,
                                "date": sent_time
                            }]
                        },
                        "$unset": {
                            "email_queue_status": "",
                            "email_queued_for": "",
                            "email_queued_at": "",
                            "email_queue_reason": "",
                            "email_queue_subject": "",
                            "email_queue_body": "",
                            "email_queue_signature": ""
                        }
                    }
                )

                print(f"✅ Queued email sent: {receiver}")

            except Exception as e:
                print(f"❌ Queued email failed for {receiver}: {e}")
                is_undeliverable = isinstance(
                    e,
                    smtplib.SMTPRecipientsRefused
                )

                collection.update_one(
                    {"_id": person["_id"]},
                    {
                        "$set": {
                            "email_sent": False,
                            "email_sending": False,
                            "sending_initial": False,
                            "email_delivery_status": (
                                "undeliverable"
                                if is_undeliverable
                                else "failed"
                            ),
                            "email_failure_type": (
                                "undeliverable"
                                if is_undeliverable
                                else "delivery_error"
                            ),
                            "undeliverable_address": (
                                receiver if is_undeliverable else ""
                            ),
                            "email_error": (
                                "Recipient refused by SMTP: " +
                                "; ".join(
                                    f"{addr}: {err!r}"
                                    for addr, err in getattr(
                                        e,
                                        "recipients",
                                        {}
                                    ).items()
                                )
                            )
                            if is_undeliverable and getattr(e, "recipients", None)
                            else str(e),
                            "email_failed_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "email_retry_after": get_next_retry_time().strftime("%Y-%m-%d %H:%M:%S"),
                            "campaign_status": (
                                "Undeliverable"
                                if is_undeliverable
                                else "Delivery Failed"
                            )
                        },
                        "$unset": {
                            "email_queue_status": "",
                            "email_queued_for": "",
                            "email_queued_at": "",
                            "email_queue_reason": "",
                            "email_queue_subject": "",
                            "email_queue_body": "",
                            "email_queue_signature": ""
                        }
                    }
                )

            # Keep the same configured gap between bulk/queued emails.
            if index < len(queued) - 1 and is_email_sending_window():
                time.sleep(EMAIL_INTERVAL_SECONDS)

    except Exception as e:
        print(f"❌ Queued email worker error: {e}")
        logger.exception("Queued email worker error")
    finally:
        if server:
            try:
                server.quit()
            except Exception:
                pass


def clean_website(url):
    try:
        parsed = urlparse(url)
        return urlunparse((parsed.scheme, parsed.netloc, "", "", "", ""))
    except Exception:
        return url

def extract_domain(url):
    try:
        domain = urlparse(url).netloc.lower()
        domain = domain.replace("www.", "")
        return DOMAIN_REPLACEMENTS.get(domain, domain)
    except Exception:
        return None

def get_domain(url):
    try:
        domain = urlparse(str(url)).netloc
        return domain.replace("www.", "")
    except:
        return ""

def find_pattern_and_sample(domain):

    conn = http.client.HTTPSConnection(
        "google.serper.dev",
        timeout=10
    )

    payload = json.dumps({
        "q": f"email@{domain} + RocketReach + Universal Pattern"
    })

    headers = {
        "X-API-KEY": SERPER_API_KEY,
        "Content-Type": "application/json"
    }

    try:
        conn.request("POST", "/search", payload, headers)

        res = conn.getresponse()

        if res.status != 200:
            print(f"Serper error: HTTP {res.status}")
            return None, None

        data = json.loads(
            res.read().decode("utf-8")
        )

    except Exception as e:
        print(f"Serper request failed for {domain}: {e}")
        return None, None

    finally:
        conn.close()

    example_email = None
    pattern = None

    position_one = next(
        (
            item for item in data.get("organic", [])
            if item.get("position") == 1
        ),
        None
    )

    if position_one:

        snippet = position_one.get("snippet", "")

        email_match = re.search(
            r'\b[A-Za-z0-9._%+-]+@' + re.escape(domain),
            snippet
        )

        if email_match:
            example_email = email_match.group(0)

        pattern_match = re.search(
            r'(\[[^\]]+\][\.\-_]?)+',
            snippet
        )

        if pattern_match:
            pattern = pattern_match.group(0)

    return example_email, pattern


def generate_company_about(company, website, title):

    prompt = f"""
You are writing the FIRST cold email for a B2B prospect.

Prospect Company:
Company Name: {company}
Website: {website}

Prospect Job Title:
{title}

Your task:

Write ONLY the opening paragraph for the first email.

Goals:

- Introduce what the company appears to do.
- Mention one likely business challenge related to their business.
- Mention ONLY ONE relevant VAIS capability.
- Do NOT explain all VAIS services.
- Do NOT sound like a sales pitch.
- Create curiosity rather than trying to sell everything.
- Write naturally as if written by a human.

You may choose ONLY ONE of these capabilities:

• Real-time Intent Data

OR

• AI Lead Scoring

OR

• MQL to BANT Campaign Fulfilment

Never mention more than one capability.

Personalize the paragraph for someone working as {title}.

Requirements:

- Maximum 20 words.
- No greeting.
- No sign-off.
- HTML ready.
- No Markdown.
- No bullet points.
- Return ONLY the paragraph.
- Do NOT use *, **, _, #, or backticks.
"""
    
    for attempt in range(5):

        try:

            response = gemini_client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt
            )

            text = response.text.strip()

            text = text.replace("**", "")
            text = text.replace("*", "")
            text = text.replace("_", "")
            text = text.replace("`", "")

            return text

        except Exception as e:

            print(f"Gemini Error (Attempt {attempt+1}): {e}")

            if attempt < 4:
                time.sleep(5)
            else:
                return None


FOLLOWUP_TEMPLATES = {
    1: """
    <p>Hi {first_name},</p>

    <p>Following up on my note about how your team spots in-market accounts.</p>

    <p>There's a live list right now of companies actively researching solutions like {company}'s — accounts you may not have reached yet. We turn that intent into HQL and BANT-qualified leads, ready for your team to work.</p>

    <p>I'd like to walk you through that list — do you have 15-20 minutes this week?</p>
    """,

    2: """
    <p>Hi {first_name},</p>

    <p>Most teams don't find out an account was in-market until after a competitor's already in the conversation.</p>

    <p>We flag that activity early and convert it into HQL and BANT-qualified leads, so your team can reach out while the account's still deciding.</p>

    <p>Worth a quick call this week to see which accounts are active right now for {company}?</p>
    """,

    3: """
    <p>Hi {first_name},</p>

    <p>Not sure if my last couple of notes landed on a busy stretch, so trying a different angle.</p>

    <p>Several APAC teams use our HQL and BANT-qualified leads — built from real-time intent data — to prioritize outreach and cut time-to-first-meeting, instead of working broad, unqualified lists.</p>

    <p>If that would help {company}'s pipeline, happy to show you what it looks like — 15 minutes is all I need.</p>
    """,

    4: """
    <p>Hi {first_name},</p>

    <p>Is capacity or signal the bigger bottleneck for your team right now?</p>

    <p>Most teams already have some way to spot in-market accounts — what's harder is acting on all of them without adding headcount.</p>

    <p>We act as that extra layer: turning in-market accounts into HQL and BANT-ready leads on demand, so your team can scale output without a hiring cycle.</p>

    <p>Worth a quick conversation to see where that could help {company}?</p>
    """,

    5: """
    <p>Hi {first_name},</p>

    <p>I don't want to keep filling your inbox, so this will be my last note for now.</p>

    <p>If qualified, ready-to-engage leads aren't a priority right now, no worries at all — I understand. But if it's something you'd explore down the line, happy to reconnect whenever timing's better.</p>

    <p>Worth a quick reply either way so I know where things stand?</p>
    """
}


def get_followup_template(person, followup_number):
    """Return the prospect's saved follow-up template, or the default one."""
    custom_templates = (person or {}).get("followup_templates") or {}
    custom = custom_templates.get(str(followup_number))
    if isinstance(custom, str) and custom.strip():
        return custom
    return FOLLOWUP_TEMPLATES.get(followup_number)


def render_followup_template(template, first_name="", company="", title="", website=""):
    """Fill follow-up placeholders.

    Uses plain replacement instead of str.format so user-edited HTML that
    contains other braces (inline CSS etc.) cannot break rendering.
    """
    return (
        (template or "")
        .replace("{first_name}", first_name or "")
        .replace("{company}", company or "")
        .replace("{title}", title or "")
        .replace("{website}", website or "")
    )

def generate_email(first, last, domain, pattern):

    first = (first or "").lower()
    last = (last or "").lower()

    replacements = {
        "[first]": first,
        "[last]": last,
        "[first_initial]": first[:1],
        "[last_initial]": last[:1],
        "[f]": first[:1],
        "[l]": last[:1]
    }

    email = pattern.lower()

    for k, v in replacements.items():
        email = email.replace(k, v)

    return email + "@" + domain


def get_access_token():

    accounts = msal_app.get_accounts()

    print("Cached accounts:", len(accounts))

    result = None
    consent_required = False

    if accounts:
        result = msal_app.acquire_token_silent_with_error(
            SCOPES,
            account=accounts[0]
        )

        print(
            "Silent token result:",
            {
                "has_access_token": bool(result.get("access_token")),
                "error": result.get("error"),
                "suberror": result.get("suberror"),
                "error_description": result.get("error_description", "")[:300]
            }
            if isinstance(result, dict)
            else result
        )

        consent_required = (
            isinstance(result, dict)
            and result.get("error") in {"invalid_grant", "interaction_required"}
            and (
                result.get("suberror") == "consent_required"
                or "consent" in str(result.get("error_description", "")).lower()
            )
        )

    if not result or consent_required or "access_token" not in result:
        print("Microsoft authorization is required for Outlook + Teams.")
        print("Starting Microsoft device-code login...")

        flow = msal_app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            raise Exception(
                "Unable to start Microsoft device login: "
                f"{flow}"
            )

        print("=" * 70)
        print(flow.get("message", ""))
        print("=" * 70)
        print("Complete the Microsoft sign-in in your browser.")
        print("The code above is generated by Microsoft; do not invent one.")

        result = msal_app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception(
            f"Authentication failed: {result}"
        )

    if cache.has_state_changed:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())

        print("Token cache updated:", CACHE_FILE)

    print("Microsoft authentication successful.")
    return result["access_token"]


@app.get("/", response_class=HTMLResponse)
def home(request: Request):

    query = {
        "$and": [
            prospect_result_filter(),
            {
                "email_queue_status": {"$ne": "queued"}
            },
            {
                "email": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            },
            {
                "domain": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            }
        ]
    }

    data = list(
        collection.find(
            query,
            {"_id": 0}
        ).sort("extracted_date", -1)
    )

    dates = collection.distinct("extracted_date")
    min_date = min(dates) if dates else None
    max_date = max(dates) if dates else None

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "data": data,
            "min_date": min_date,
            "max_date": max_date
        }
    )


@app.post(
    "/",
    response_class=HTMLResponse,
    tags=["Prospects"],
    summary="Search prospects and enrich records",
    description="Search Apollo, enrich returned prospects, apply blacklist/duplicate checks, and store new prospects."
)
async def search(
    request: Request,
    job_titles: str = Form(None),
    country: str = Form(None),
    company_name: str = Form(None),
    company_csv: UploadFile = File(None),
    limit: int = Form(10),
    from_date: str = Form(None),
    to_date: str = Form(None)
):
    
    companies = []
    titles = []
    countries = []

    search_url = "https://api.apollo.io/api/v1/mixed_people/api_search"

    
    # If date filter is applied
    if from_date or to_date:
        query = {
            "$and": [
                prospect_result_filter(),
                {
                    "email_queue_status": {"$ne": "queued"}
                },
                {
                    "email": {
                        "$exists": True,
                        "$nin": ["", None]
                    }
                },
                {
                    "domain": {
                        "$exists": True,
                        "$nin": ["", None]
                    }
                }
            ]
        }
        if from_date and to_date:
            query["extracted_date"] = {
                "$gte": from_date,
                "$lte": to_date
            }
        elif from_date:
            query["extracted_date"] = {
                "$gte": from_date
            }
        elif to_date:
            query["extracted_date"] = {
                "$lte": to_date
            }
        
        data = list(collection.find(query, {"_id": 0}).sort("extracted_date", -1))
        
        dates = collection.distinct("extracted_date")
        min_date = min(dates) if dates else None
        max_date = max(dates) if dates else None
        
        return templates.TemplateResponse(
            request=request,
            name="index.html",
            context={
                "data": data,
                "min_date": min_date,
                "max_date": max_date,
                "from_date": from_date,
                "to_date": to_date
            }
        )
    
    # ============================================================
    # JOB TITLES
    # ============================================================

    titles = []

    if job_titles:
        titles = [
            x.strip()
            for x in job_titles.split(",")
            if x.strip()
        ]


    # ============================================================
    # COMPANY INPUT
    # Manual Company Name + CSV Upload
    # ============================================================

    companies = []


    # ------------------------------------------------------------
    # 1. Manual Company Name
    # ------------------------------------------------------------

    if company_name:

        manual_companies = [
            x.strip()
            for x in company_name.split(",")
            if x.strip()
        ]

        companies.extend(manual_companies)


    # ------------------------------------------------------------
    # 2. CSV Company Names
    # ------------------------------------------------------------

    if company_csv and company_csv.filename:

        print("=" * 60)
        print("CSV FILE:", company_csv.filename)

        try:

           # Read uploaded CSV
            contents = await company_csv.read()

            # Try UTF-8 first, then Windows-1252
            try:
                decoded = contents.decode("utf-8-sig")
            except UnicodeDecodeError:
                decoded = contents.decode("cp1252")

            # Read CSV
            csv_reader = csv.DictReader(
                io.StringIO(decoded)
            )
            print("CSV Columns:", csv_reader.fieldnames)

            # Find Company Name column
            company_column = None

            for column in csv_reader.fieldnames or []:

                normalized = column.strip().lower()

                if normalized in [
                    "company",
                    "company name",
                    "company_name",
                    "organization",
                    "organization name"
                ]:
                    company_column = column
                    break

            if not company_column:

                return templates.TemplateResponse(
                    request=request,
                    name="index.html",
                    context={
                        "data": [],
                        "error": (
                            "CSV must contain a Company Name column. "
                            "Accepted names: Company Name, company, company_name"
                        )
                    }
                )


            # Extract companies
            csv_companies = []

            for row in csv_reader:

                company_value = (
                    row.get(company_column) or ""
                ).strip()

                if company_value:
                    csv_companies.append(
                        company_value
                    )


            print(
                "Companies from CSV:",
                csv_companies
            )

            companies.extend(
                csv_companies
            )

        except Exception as e:

            print(
                "CSV Error:",
                str(e)
            )

            return templates.TemplateResponse(
                request=request,
                name="index.html",
                context={
                    "data": [],
                    "error": f"Unable to read CSV: {str(e)}"
                }
            )


    # ------------------------------------------------------------
    # Remove duplicate companies
    # ------------------------------------------------------------

    companies = list(
        dict.fromkeys(
            companies
        )
    )


    print("=" * 60)
    print(
        "Companies to search:",
        companies
    )

    print(
        "Total companies:",
        len(companies)
    )


    search_url = (
        "https://api.apollo.io/api/v1/mixed_people/api_search"
    )

    # ============================================================
    # APOLLO SEARCH
    # Result limit is GLOBAL across all companies
    # ============================================================

    countries = []

    if country:
        countries = [
            x.strip()
            for x in country.split(",")
            if x.strip()
        ]


    all_people = []

    total_collected = 0

    # ============================================================
    # COMPANY SEARCH
    # ============================================================

    if companies:

        for company_name in companies:

            # ----------------------------------------------------
            # STOP when GLOBAL LIMIT is reached
            # ----------------------------------------------------
            if total_collected >= limit:
                break

            remaining_limit = limit - total_collected

            print("=" * 60)
            print(
                f"Searching company: {company_name}"
            )

            print(
                f"Remaining global limit: {remaining_limit}"
            )


            # ====================================================
            # FIND APOLLO ORGANIZATION
            # ====================================================

            org_url = (
                "https://api.apollo.io/api/v1/organizations/search"
            )

            org_payload = {
                "q_organization_name": company_name,
                "page": 1,
                "per_page": 1
            }
            org_response = requests.post(
                org_url,
                json=org_payload,
                headers=headers,
                timeout=60
            )

            org_data = org_response.json()

            organizations = org_data.get(
                "organizations",
                []
            )


            # ----------------------------------------------------
            # COMPANY NOT FOUND
            # ----------------------------------------------------

            if not organizations:

                print(
                    f"Company not found: {company_name}"
                )

                continue


            organization_id = organizations[0].get("id")

            print(
                f"Organization ID: {organization_id}"
            )

            # ============================================================
            # APOLLO PEOPLE SEARCH
            # Keep fetching until we get LIMIT NEW prospects
            # ============================================================

            page = 1

            while total_collected < limit:

                remaining_limit = limit - total_collected

                print("=" * 60)
                print(
                    f"Apollo page {page} | Need {remaining_limit} NEW prospects"
                )

                payload = {
                    "page": page,

                    # Ask Apollo for more than remaining because
                    # some results may already exist in MongoDB
                    "per_page": min(remaining_limit, 10),

                    "organization_ids": [
                        organization_id
                    ]
                }

                # ----------------------------------------------------
                # JOB TITLES
                # ----------------------------------------------------

                if titles:
                    payload["person_titles"] = titles

                # ----------------------------------------------------
                # COUNTRY
                # ----------------------------------------------------

                if countries:
                    payload["person_locations"] = countries

                # ----------------------------------------------------
                # SEND REQUEST TO APOLLO
                # ----------------------------------------------------

                response = requests.post(
                    search_url,
                    json=payload,
                    headers=headers,
                    timeout=60
                )

                response_data = response.json()

                people = response_data.get(
                    "people",
                    []
                )

                print(
                    f"Apollo returned {len(people)} people on page {page}"
                )

                # ----------------------------------------------------
                # No more Apollo results
                # ----------------------------------------------------

                if not people:
                    print(
                        f"No more Apollo people available for {company_name}"
                    )
                    break

                # ----------------------------------------------------
                # CHECK EACH PERSON AGAINST MONGODB
                # ----------------------------------------------------

                for person in people:

                    if total_collected >= limit:
                        break

                    person_email = (
                        person.get("email") or ""
                    ).strip().lower()

                    person_name = person.get("name")

                    person_company = (
                        person.get("organization", {})
                        .get("name")
                    )

                    # ------------------------------------------------
                    # Check existing email
                    # ------------------------------------------------

                    if person_email:

                        existing = collection.find_one({
                            "email": person_email
                        })

                    else:

                        # ------------------------------------------------
                        # If Apollo does not provide email,
                        # check name + company
                        # ------------------------------------------------

                        existing = collection.find_one({
                            "full_name": person_name,
                            "company": person_company
                        })

                    # ------------------------------------------------
                    # SKIP EXISTING
                    # ------------------------------------------------

                    if existing:

                        print(
                            f"⚠ Already exists in DB: {person_email or person_name}"
                        )

                        continue

                    # ------------------------------------------------
                    # NEW PERSON
                    # ------------------------------------------------

                    all_people.append(person)

                    total_collected += 1

                    print(
                        f"✓ NEW prospect: "
                        f"{person_email or person_name} "
                        f"({total_collected}/{limit})"
                    )

                # ----------------------------------------------------
                # Move to next Apollo page
                # ----------------------------------------------------

                page += 1

                # Safety: Apollo returned fewer than requested,
                # so there may be no more useful results.
                if len(people) < payload["per_page"]:
                    print(
                        "Apollo returned fewer results than requested."
                    )

                    # Don't immediately break because the returned
                    # results may have contained existing records.
                    # Continue to next page to find new prospects.
                    continue


    # ============================================================
    # NO COMPANY FILTER
    # EXACT LIMIT OF NEW PROFILES
    # ============================================================

    else:

        page = 1

        while total_collected < limit:

            remaining_limit = limit - total_collected

            payload = {
                "page": page,
                "per_page": remaining_limit
            }

            if titles:
                payload["person_titles"] = titles

            if countries:
                payload["person_locations"] = countries

            response = requests.post(
                search_url,
                json=payload,
                headers=headers,
                timeout=60
            )

            response_data = response.json()

            people = response_data.get(
                "people",
                []
            )

            if not people:
                break

            for person in people:

                if total_collected >= limit:
                    break

                person_email = (
                    person.get("email") or ""
                ).strip().lower()

                person_name = (
                    person.get("name") or ""
                ).strip()

                person_company = (
                    person.get("organization", {})
                    .get("name")
                )

                if person_email:
                    existing = collection.find_one({
                        "email": person_email
                    })
                else:
                    existing = collection.find_one({
                        "full_name": person_name,
                        "company": person_company
                    })

                if existing:
                    continue

                all_people.append(person)

                total_collected += 1

            page += 1

    # ============================================================
    # FINAL PEOPLE LIST
    # ============================================================

    # HARD SAFETY LIMIT
    # Never allow more profiles than requested
    people = all_people[:limit]

    print("=" * 60)
    print(f"REQUESTED LIMIT: {limit}")
    print(f"FINAL PEOPLE COUNT: {len(people)}")
    print("=" * 60)
    
    # Get current date and time
    current_datetime = datetime.now()
    extracted_date = current_datetime.strftime("%Y-%m-%d")

    inserted_count = 0
    skipped_count = 0

    for index, p in enumerate(people):
        if index >= limit:
            print(f"Stopping insertion. Limit {limit} reached.")
            break

        person_id = p.get("id")

        enrich_url = (
            "https://api.apollo.io/api/v1/people/match"
            "?run_waterfall_email=false"
            "&run_waterfall_phone=false"
        )

        enrich_response = requests.post(
            enrich_url,
            json={"id": person_id},
            headers=headers,
            timeout=60
        )

        person = enrich_response.json().get("person", {})

        org = person.get("organization") or {}

        website = org.get("website_url", "")

        website = clean_website(website)

        domain = extract_domain(website)

        email = (person.get("email") or "").strip()

        email_pattern = ""
        example_email = ""

        if not email and domain:

            # ======================================================
            # STEP 1 : Reuse existing pattern from MongoDB
            # ======================================================
            existing_pattern = collection.find_one(
                {
                    "domain": domain,
                    "email_pattern": {
                        "$nin": ["", None, "Not Found"]
                    }
                },
                {
                    "email_pattern": 1,
                    "example_email": 1
                }
            )

            if existing_pattern:

                email_pattern = existing_pattern.get("email_pattern", "")
                example_email = existing_pattern.get("example_email", "")

                print(f"Using existing pattern for {domain}: {email_pattern}")

            else:

                # ======================================================
                # STEP 2 : Search Serper only if pattern not found
                # ======================================================
                print(f"Searching RocketReach pattern for {domain}")

                example_email, email_pattern = find_pattern_and_sample(domain)

            # ======================================================
            # STEP 3 : Generate email
            # ======================================================
            if email_pattern:
                email = generate_email(
                    person.get("first_name", ""),
                    person.get("last_name", ""),
                    domain,
                    email_pattern
                )

                print("Generated Email:", email)

        record = {
            "first_name": person.get("first_name"),
            "last_name": person.get("last_name"),
            "full_name": person.get("name"),
            "title": person.get("title"),
            "email": email,
            "email_pattern": email_pattern,
            "example_email": example_email,
            "email_status": person.get("email_status"),
            "linkedin": person.get("linkedin_url"),
            "state": person.get("state"),
            "country": person.get("country"),
            "company": org.get("name"),
            "website": website,
            "domain": domain,
            "extracted_date": extracted_date,
        }

        # CHECK BLACKLIST FIRST
        # Email is the primary blacklist key so a removed/blacklisted
        # prospect cannot re-appear even if Apollo later returns a
        # changed title, company, or domain.
        blacklist_match = None

        if record.get("email"):
            blacklist_match = blacklist_collection.find_one({
                "email": record.get("email")
            })

        if not blacklist_match:
            blacklist_match = blacklist_collection.find_one({
                "full_name": record.get("full_name"),
                "title": record.get("title"),
                "domain": record.get("domain")
            })

        if blacklist_match:
            print(f"🚫 Blacklist match - skipping: {record.get('email')}")
            continue

        if email:
            existing = collection.find_one({"email": email})
        else:
            existing = collection.find_one({
                "full_name": person.get("name"),
                "company": org.get("name")
            })

        if not existing:
            collection.insert_one(record)
            inserted_count += 1
            print(f"✓ Inserted: {email}")
        else:
            print(f"⚠ Already exists: {email}")


    print(f"\n📊 Summary: {inserted_count} inserted, {skipped_count} skipped (no email)")

    query = {
        "$and": [
            prospect_result_filter(),
            {
                "email_queue_status": {"$ne": "queued"}
            },
            {
                "email": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            },
            {
                "domain": {
                    "$exists": True,
                    "$nin": ["", None]
                }
            }
        ]
    }

    data = list(
        collection.find(
            query,
            {"_id": 0}
        ).sort("extracted_date", -1)
    )
        
    dates = collection.distinct("extracted_date")
    min_date = min(dates) if dates else None
    max_date = max(dates) if dates else None

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "data": data,
            "min_date": min_date,
            "max_date": max_date
        }
    )

class EmailRecipient(BaseModel):
    email: str = Field(
        ...,
        description="Recipient email address",
        examples=["person@example.com"]
    )
    first_name: str = Field(
        "",
        description="Recipient first name",
        examples=["John"]
    )
    last_name: str = Field(
        "",
        description="Recipient last name",
        examples=["Doe"]
    )
    full_name: str = Field(
        "",
        description="Recipient full name",
        examples=["John Doe"]
    )
    title: str = Field(
        "",
        description="Recipient job title",
        examples=["Marketing Manager"]
    )
    company: str = Field(
        "",
        description="Recipient company",
        examples=["Example Corp"]
    )


class EmailCampaignRequest(BaseModel):
    recipients: list[EmailRecipient] = Field(
        ...,
        description="One or more recipients to process.",
        min_length=1
    )
    subject: str = Field(
        ...,
        description="Email subject template. Supported placeholders: {company}, {first_name}, {title}.",
        examples=["How is {company} finding in-market accounts?"]
    )
    body: str = Field(
        ...,
        description="HTML email body template. Supported placeholders: {company}, {first_name}, {title}, {company_about}, {signature}.",
        examples=["<p>Hi {first_name},</p><p>{company_about}</p><p>Regards</p>"]
    )
    signature: str = Field(
        "",
        description="HTML signature appended to the email body.",
        examples=["<p>Regards,<br>Valasys Media</p>"]
    )

    class Config:
        schema_extra = {
            "example": {
                "recipients": [
                    {
                        "email": "person@example.com",
                        "first_name": "John",
                        "last_name": "Doe",
                        "full_name": "John Doe",
                        "title": "Marketing Manager",
                        "company": "Example Corp"
                    }
                ],
                "subject": "How is {company} finding in-market accounts?",
                "body": "<p>Hi {first_name},</p><p>{company_about}</p>",
                "signature": "<p>Regards,<br>Valasys Media</p>"
            }
        }


class EmailCampaignResult(BaseModel):
    email: str = Field(..., description="Recipient email address")
    status: str = Field(..., description="Processing result for this recipient")
    error: str | None = Field(None, description="Error message when processing fails")
    queued_for: str | None = Field(None, description="Scheduled queue time in IST")


class EmailCampaignResponse(BaseModel):
    success: bool = Field(..., description="Whether the API request completed")
    sent: int = Field(0, description="Number of emails sent immediately")
    queued: int = Field(0, description="Number of emails queued for later")
    failed: int = Field(0, description="Number of failed recipients")
    skipped: int = Field(0, description="Number of skipped recipients")
    results: list[EmailCampaignResult] = Field(default_factory=list, description="Per-recipient results")
    message: str = Field("", description="Human-readable result summary")
    error: str | None = Field(None, description="Top-level API error")


class ErrorResponse(BaseModel):
    success: bool = Field(False, description="Always false for an error response")
    error: str = Field(..., description="Error message")


@app.post(
    "/queue-emails",
    tags=["Email Campaign"],
    summary="Queue initial emails",
    description="Queue selected initial campaign emails for the next valid weekday sending window.",
    response_model=EmailCampaignResponse,
    responses={
        400: {"model": ErrorResponse, "description": "Invalid request"},
        500: {"model": ErrorResponse, "description": "Server error"}
    }
)
async def queue_emails(payload: EmailCampaignRequest):
    """Queue selected initial emails for the next Monday-Friday 09:00 IST window.

    Queued records remain in MongoDB with email_queue_status='queued' and are
    deliberately excluded from the main prospect Results table until sent.
    """
    try:
        data = payload.dict()
        recipients = data.get("recipients", [])
        subject_template = data.get("subject", "")
        body_template = data.get("body", "")
        signature_html = data.get("signature", "")

        if not recipients:
            return JSONResponse(
                status_code=400,
                content={"success": False, "error": "No recipients provided"}
            )

        next_window = get_next_email_window_start()
        queued_count = 0
        skipped_count = 0
        results = []
        seen_emails = set()

        for recipient in recipients:
            receiver = (recipient.get("email") or "").strip().lower()
            if not receiver or receiver in seen_emails:
                continue
            seen_emails.add(receiver)

            person = collection.find_one({"email": receiver})
            if not person:
                skipped_count += 1
                results.append({
                    "email": receiver,
                    "status": "failed",
                    "error": "MongoDB record not found"
                })
                continue

            if person.get("email_sent") is True:
                skipped_count += 1
                results.append({
                    "email": receiver,
                    "status": "already_sent"
                })
                continue

            first_name = recipient.get("first_name") or person.get("first_name", "")
            title = recipient.get("title") or person.get("title", "")
            company = recipient.get("company") or person.get("company", "")
            website = person.get("website", "")

            company_about = generate_company_about(
                company,
                website,
                title
            ) or ""

            subject = (
                subject_template
                .replace("{company}", company)
                .replace("{first_name}", first_name)
                .replace("{title}", title)
            )

            body = (
                body_template
                .replace("{company}", company)
                .replace("{first_name}", first_name)
                .replace("{title}", title)
                .replace("{company_about}", company_about)
            )
            body = body + "<br><br>" + signature_html
            body = body.replace("{signature}", get_signature())

            queued_for = queue_initial_email(
                person["_id"],
                receiver,
                subject,
                body,
                signature_html,
                next_window,
                reason="manual_queue_request"
            )

            queued_count += 1
            results.append({
                "email": receiver,
                "status": "queued",
                "queued_for": queued_for
            })

        return JSONResponse(content={
            "success": True,
            "queued": queued_count,
            "skipped": skipped_count,
            "results": results,
            "message": (
                f"{queued_count} email(s) queued for "
                f"{next_window.strftime('%d %b %Y %I:%M %p')} IST."
            )
        })

    except Exception as e:
        logger.exception("QUEUE EMAIL REQUEST FAILED")
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e)}
        )




# ============================================================
# BACKGROUND EMAIL JOBS
# Send/queue runs in a worker thread so the browser gets a job id at once,
# the compose modal can close, and progress can be followed from any page.
# Jobs live in memory only; they are lost if the server restarts.
# ============================================================

EMAIL_JOBS = {}
EMAIL_JOBS_LOCK = threading.Lock()
EMAIL_JOB_TTL_SECONDS = 6 * 60 * 60


class EmailJobRequest(EmailCampaignRequest):
    action: str = Field(
        "send",
        description="'send' to send now (queues automatically outside the window) or 'queue' to queue."
    )


def _prune_email_jobs():
    cutoff = time.time() - EMAIL_JOB_TTL_SECONDS
    with EMAIL_JOBS_LOCK:
        for job_id in [key for key, job in EMAIL_JOBS.items() if job["created"] < cutoff]:
            EMAIL_JOBS.pop(job_id, None)


def _run_email_job(job_id, action, payload):
    try:
        if action == "queue":
            response = asyncio.run(queue_emails(payload))
        else:
            response = send_emails(payload)
        result = json.loads(response.body)
    except Exception as exc:
        logger.exception("EMAIL JOB %s FAILED", job_id)
        result = {"success": False, "error": str(exc), "results": []}

    with EMAIL_JOBS_LOCK:
        job = EMAIL_JOBS.get(job_id)
        if job:
            job["status"] = "done"
            job["finished"] = time.time()
            job["result"] = result


@app.post("/email-jobs", tags=["Email Campaign"], summary="Start a background send/queue job")
def start_email_job(payload: EmailJobRequest):
    _prune_email_jobs()

    action = "queue" if payload.action == "queue" else "send"
    emails = []
    for recipient in payload.recipients:
        email = (recipient.email or "").strip().lower()
        if email and email not in emails:
            emails.append(email)

    if not emails:
        return JSONResponse(status_code=400, content={"success": False, "error": "No recipients provided"})

    campaign = EmailCampaignRequest(
        recipients=payload.recipients,
        subject=payload.subject,
        body=payload.body,
        signature=payload.signature
    )

    job_id = uuid.uuid4().hex
    with EMAIL_JOBS_LOCK:
        EMAIL_JOBS[job_id] = {
            "id": job_id,
            "action": action,
            "emails": emails,
            "status": "running",
            "created": time.time(),
            "finished": None,
            "result": None
        }

    threading.Thread(
        target=_run_email_job,
        args=(job_id, action, campaign),
        name=f"email-job-{job_id[:8]}",
        daemon=True
    ).start()

    logger.info("EMAIL JOB %s started: %s x %s", job_id, action, len(emails))
    return {"success": True, "job_id": job_id, "action": action, "emails": emails}


@app.get("/email-jobs/{job_id}", tags=["Email Campaign"], summary="Per-recipient progress of a background job")
def get_email_job(job_id: str):
    with EMAIL_JOBS_LOCK:
        job = EMAIL_JOBS.get(job_id)
        job = dict(job) if job else None

    if not job:
        return JSONResponse(status_code=404, content={"success": False, "error": "Job not found"})

    emails = job["emails"]
    records = {
        str(row.get("email", "")).lower(): row
        for row in collection.find(
            {"email": {"$in": emails}},
            {
                "_id": 0,
                "email": 1,
                "email_sent": 1,
                "email_delivery_status": 1,
                "email_queue_status": 1,
                "email_queued_for": 1,
                "sending_initial": 1,
                "email_error": 1
            }
        )
    }

    final = {}
    if job["status"] == "done" and job["result"]:
        final = {
            str(item.get("email", "")).lower(): item
            for item in (job["result"].get("results") or [])
        }

    failed_states = {"failed", "undeliverable", "smtp_failed", "bounced", "quarantined"}
    items = []
    for email in emails:
        row = records.get(email, {})
        outcome = final.get(email, {})
        status = "waiting"
        error = ""

        if row.get("email_sent") is True:
            status = "sent"
        elif row.get("email_queue_status") == "queued":
            status = "queued"
        elif row.get("sending_initial") is True:
            status = "sending"
        elif row.get("email_delivery_status") in failed_states:
            status = "failed"
            error = row.get("email_error", "")

        if job["status"] == "done" and status in {"waiting", "sending"}:
            outcome_status = str(outcome.get("status", "")).lower()
            if outcome_status in {"sent", "already_sent", "already_sent_or_in_progress"}:
                status = "sent"
            elif outcome_status == "queued":
                status = "queued"
            else:
                status = "failed"
                error = outcome.get("error") or job["result"].get("error") or "Not sent"

        items.append({
            "email": email,
            "status": status,
            "error": error,
            "queued_for": row.get("email_queued_for", "") or outcome.get("queued_for", "")
        })

    # Recipients are processed in order: the first still-waiting one is being worked on.
    if job["status"] == "running" and not any(item["status"] == "sending" for item in items):
        for item in items:
            if item["status"] == "waiting":
                item["status"] = "sending"
                break

    counts = {key: sum(1 for item in items if item["status"] == key)
              for key in ("sent", "queued", "failed", "sending", "waiting")}

    return {
        "success": True,
        "job_id": job_id,
        "action": job["action"],
        "status": job["status"],
        "total": len(items),
        "done": counts["sent"] + counts["queued"] + counts["failed"],
        "counts": counts,
        "items": items,
        "message": (job["result"] or {}).get("message", "") if job["status"] == "done" else ""
    }


@app.get("/email-progress", tags=["Email Campaign"])
def email_progress(emails: str = ""):
    """
    Return current stored send/queue state for a comma-separated set of emails.
    Used by the UI to give live status while a send/queue request is being processed.
    """
    email_list = [
        item.strip().lower()
        for item in (emails or "").split(",")
        if item.strip()
    ]

    if not email_list:
        return {
            "success": True,
            "total": 0,
            "sent": 0,
            "queued": 0,
            "pending": 0,
            "failed": 0,
            "records": []
        }

    records = list(
        collection.find(
            {"email": {"$in": email_list}},
            {
                "_id": 0,
                "email": 1,
                "email_sent": 1,
                "email_delivery_status": 1,
                "email_queue_status": 1,
                "email_queued_for": 1,
                "campaign_status": 1
            }
        )
    )

    by_email = {str(row.get("email", "")).lower(): row for row in records}

    sent = queued = failed = pending = 0
    result_rows = []

    for email in email_list:
        row = by_email.get(email, {})
        is_sent = row.get("email_sent") is True
        is_queued = row.get("email_queue_status") == "queued"
        delivery = row.get("email_delivery_status", "")

        if is_sent:
            status = "sent"
            sent += 1
        elif is_queued:
            status = "queued"
            queued += 1
        elif delivery in {"failed", "undeliverable", "smtp_failed", "bounced", "quarantined"}:
            status = "failed"
            failed += 1
        else:
            status = "pending"
            pending += 1

        result_rows.append({
            "email": email,
            "status": status,
            "queued_for": row.get("email_queued_for", ""),
            "campaign_status": row.get("campaign_status", "")
        })

    return {
        "success": True,
        "total": len(email_list),
        "sent": sent,
        "queued": queued,
        "pending": pending,
        "failed": failed,
        "records": result_rows
    }

@app.post(
    "/send-emails",
    tags=["Email Campaign"],
    summary="Send or queue initial emails",
    description=(
        "Send initial campaign emails immediately when the Monday-Friday 09:00-17:00 IST "
        "window is open; otherwise queue them for the next valid window."
    ),
    response_model=EmailCampaignResponse,
    responses={
        400: {"model": ErrorResponse, "description": "Invalid request"},
        500: {"model": ErrorResponse, "description": "Server error"}
    }
)
def _open_smtp_connection():
    """Connect, start TLS and log in to the outbound SMTP server."""
    logger.info("Connecting SMTP: %s:%s", SMTP_SERVER, SMTP_PORT)

    smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60)
    logger.info("SMTP connection established")

    smtp.set_debuglevel(1)

    logger.info("Starting TLS")
    smtp.starttls(context=ssl.create_default_context())
    logger.info("TLS started successfully")

    logger.info("Logging into SMTP as: %s", SENDER_EMAIL)
    smtp.login(SENDER_EMAIL, PASSWORD)
    logger.info("SMTP LOGIN SUCCESS")

    return smtp


def send_emails(payload: EmailCampaignRequest):

    server = None

    logger.info("=" * 80)
    logger.info("SEND EMAIL REQUEST STARTED")

    try:

        # ============================================================
        # READ REQUEST
        # ============================================================

        data = payload.dict()

        recipients = data.get("recipients", [])
        subject_template = data.get("subject", "")
        body_template = data.get("body", "")
        signature_html = data.get("signature", "")

        logger.info(
            "Recipients received from frontend: %s",
            len(recipients)
        )

        logger.info(
            "Subject template: %s",
            subject_template
        )

        logger.info(
            "Body length: %s",
            len(body_template)
        )

        # ============================================================
        # LOG ALL RECIPIENTS
        # ============================================================

        for r in recipients:

            logger.info(
                "Frontend recipient: email=%s | name=%s | company=%s",
                r.get("email"),
                r.get("full_name"),
                r.get("company")
            )

        # ============================================================
        # REMOVE DUPLICATES
        # ============================================================

        unique_recipients = []
        seen_emails = set()

        for recipient in recipients:

            email = (
                recipient.get("email") or ""
            ).strip().lower()

            if not email:

                logger.warning(
                    "Skipping recipient because email is empty: %s",
                    recipient
                )

                continue

            if email in seen_emails:

                logger.warning(
                    "Duplicate recipient removed: %s",
                    email
                )

                continue

            seen_emails.add(email)
            unique_recipients.append(recipient)

        recipients = unique_recipients

        logger.info(
            "Unique recipients after duplicate removal: %s",
            len(recipients)
        )

        if not recipients:

            logger.error(
                "NO RECIPIENTS PROVIDED"
            )

            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "No recipients provided"
                }
            )

        # ============================================================
        # QUEUE WHOLE BULK REQUEST IF OUTSIDE IST WINDOW
        # ============================================================
        if not is_email_sending_window():
            next_window = get_next_email_window_start()
            queued_count = 0
            queue_results = []

            for i, recipient in enumerate(recipients):
                receiver = (recipient.get("email") or "").strip().lower()
                if not receiver:
                    continue

                person = collection.find_one({"email": receiver})
                if not person:
                    queue_results.append({
                        "email": receiver,
                        "status": "failed",
                        "error": "MongoDB record not found"
                    })
                    continue

                if person.get("email_sent") is True:
                    queue_results.append({
                        "email": receiver,
                        "status": "already_sent"
                    })
                    continue

                subject_template_selected = subject_patterns[i % len(subject_patterns)] if 'subject_patterns' in locals() else (
                    "How is {company} finding in-market accounts?"
                    if i % 2 == 0
                    else "How does {company} spot ready-to-buy accounts?"
                )

                first_name = recipient.get("first_name", person.get("first_name", ""))
                title = recipient.get("title", person.get("title", ""))
                company = recipient.get("company", person.get("company", ""))
                website = person.get("website", "")

                company_about = generate_company_about(
                    company,
                    website,
                    title
                ) or ""

                subject = (
                    subject_template_selected
                    .replace("{company}", company)
                    .replace("{first_name}", first_name)
                    .replace("{title}", title)
                )

                body = (
                    body_template
                    .replace("{company}", company)
                    .replace("{first_name}", first_name)
                    .replace("{title}", title)
                    .replace("{company_about}", company_about)
                )
                body = body + "<br><br>" + signature_html
                body = body.replace("{signature}", get_signature())

                queued_for = queue_initial_email(
                    person["_id"],
                    receiver,
                    subject,
                    body,
                    signature_html,
                    next_window,
                    reason="bulk_requested_outside_email_window"
                )

                queued_count += 1
                queue_results.append({
                    "email": receiver,
                    "status": "queued",
                    "queued_for": queued_for
                })

            return JSONResponse(
                content={
                    "success": True,
                    "sent": 0,
                    "queued": queued_count,
                    "failed": sum(
                        1 for r in queue_results
                        if r.get("status") == "failed"
                    ),
                    "results": queue_results,
                    "message": (
                        f"{queued_count} email(s) queued for "
                        f"{next_window.strftime('%d %b %Y %I:%M %p')} IST."
                    )
                }
            )

        # ============================================================
        # CONNECT SMTP
        # ============================================================

        server = _open_smtp_connection()

        sent_count = 0
        failed_count = 0
        results = []

        subject_patterns = [
            "How is {company} finding in-market accounts?",
            "How does {company} spot ready-to-buy accounts?"
        ]

        # ============================================================
        # PROCESS EACH RECIPIENT
        # ============================================================

        for i, recipient in enumerate(recipients):

            receiver = (
                recipient.get("email") or ""
            ).strip().lower()

            logger.info("=" * 80)
            logger.info(
                "PROCESSING RECIPIENT %s/%s: %s",
                i + 1,
                len(recipients),
                receiver
            )

            try:

                if not receiver:

                    logger.warning(
                        "EMPTY EMAIL - SKIPPING"
                    )

                    continue

                # ====================================================
                # FIND MONGODB RECORD
                # ====================================================

                person = collection.find_one(
                    {
                        "email": receiver
                    }
                )

                if not person:

                    logger.error(
                        "❌ MONGODB RECORD NOT FOUND FOR: %s",
                        receiver
                    )

                    results.append({
                        "email": receiver,
                        "status": "failed",
                        "error": "MongoDB record not found"
                    })

                    failed_count += 1

                    continue

                logger.info(
                    "MongoDB record FOUND: _id=%s",
                    person.get("_id")
                )

                logger.info(
                    "Mongo record email_sent=%s | email_sending=%s | sending_initial=%s",
                    person.get("email_sent"),
                    person.get("email_sending"),
                    person.get("sending_initial")
                )

                logger.info(
                    "Mongo record company=%s | first_name=%s | title=%s",
                    person.get("company"),
                    person.get("first_name"),
                    person.get("title")
                )

                # ====================================================
                # CHECK ALREADY SENT
                # ====================================================

                if person.get("email_sent") is True:

                    logger.warning(
                        "❌ EMAIL ALREADY SENT: %s",
                        receiver
                    )

                    results.append({
                        "email": receiver,
                        "status": "already_sent"
                    })

                    continue

                # ====================================================
                # CHECK SENDING LOCK
                # ====================================================

                if person.get("sending_initial") is True:

                    logger.warning(
                        "❌ INITIAL EMAIL ALREADY BEING SENT: %s",
                        receiver
                    )

                    results.append({
                        "email": receiver,
                        "status": "already_sending"
                    })

                    continue

                # ====================================================
                # ATOMIC INITIAL EMAIL LOCK
                # ====================================================

                logger.info(
                    "Attempting initial-email lock for %s",
                    receiver
                )

                claim_result = collection.update_one(
                    {
                        "_id": person["_id"],

                        "email_sent": {
                            "$ne": True
                        },

                        "sending_initial": {
                            "$ne": True
                        }
                    },
                    {
                        "$set": {
                            "sending_initial": True,
                            "initial_send_started_at": datetime.now(
                                IST
                            ).strftime(
                                "%Y-%m-%d %H:%M:%S"
                            )
                        }
                    }
                )

                logger.info(
                    "Initial lock result: matched=%s modified=%s",
                    claim_result.matched_count,
                    claim_result.modified_count
                )

                if claim_result.modified_count == 0:

                    logger.warning(
                        "❌ FAILED TO CLAIM INITIAL EMAIL LOCK: %s",
                        receiver
                    )

                    latest = collection.find_one(
                        {"_id": person["_id"]}
                    )

                    logger.warning(
                        "Current Mongo state: email_sent=%s | sending_initial=%s | email_sending=%s",
                        latest.get("email_sent") if latest else None,
                        latest.get("sending_initial") if latest else None,
                        latest.get("email_sending") if latest else None
                    )

                    results.append({
                        "email": receiver,
                        "status": "already_sent_or_in_progress"
                    })

                    continue

                logger.info(
                    "🔒 INITIAL EMAIL LOCK ACQUIRED: %s",
                    receiver
                )

                # ====================================================
                # GET DATA
                # ====================================================

                first_name = recipient.get(
                    "first_name",
                    person.get("first_name", "")
                )

                title = recipient.get(
                    "title",
                    person.get("title", "")
                )

                company = recipient.get(
                    "company",
                    person.get("company", "")
                )

                website = person.get(
                    "website",
                    ""
                )

                logger.info(
                    "Email data: first_name=%s | title=%s | company=%s | website=%s",
                    first_name,
                    title,
                    company,
                    website
                )

                # ====================================================
                # GENERATE COMPANY ABOUT
                # ====================================================

                logger.info(
                    "Generating company_about for %s",
                    receiver
                )

                company_about = generate_company_about(
                    company,
                    website,
                    title
                )

                company_about = company_about or ""

                logger.info(
                    "Generated company_about: %s",
                    company_about
                )

                # ====================================================
                # SUBJECT
                # ====================================================

                subject_template_selected = (
                    subject_patterns[
                        i % len(subject_patterns)
                    ]
                )

                subject = (
                    subject_template_selected
                    .replace("{company}", company)
                    .replace("{first_name}", first_name)
                    .replace("{title}", title)
                )

                logger.info(
                    "Final subject: %s",
                    subject
                )

                # ====================================================
                # BODY
                # ====================================================

                body = (
                    body_template
                    .replace("{company}", company)
                    .replace("{first_name}", first_name)
                    .replace("{title}", title)
                    .replace("{company_about}", company_about)
                )

                body = (
                    body
                    + "<br><br>"
                    + signature_html
                )

                body = body.replace(
                    "{signature}",
                    get_signature()
                )

                logger.info(
                    "Final email body length: %s",
                    len(body)
                )

                # ====================================================
                # CREATE MESSAGE
                # ====================================================

                message = MIMEMultipart(
                    "related"
                )

                alternative = MIMEMultipart(
                    "alternative"
                )

                html_part = MIMEText(
                    body,
                    "html",
                    "utf-8"
                )

                alternative.attach(
                    html_part
                )

                message.attach(
                    alternative
                )

                message["From"] = SENDER_EMAIL
                message["To"] = receiver
                message["Subject"] = subject

                message_id = make_msgid()

                clean_message_id = (
                    message_id.strip("<>")
                )

                message["Message-ID"] = message_id
                message["Date"] = formatdate(
                    localtime=True
                )

                logger.info(
                    "Message-ID generated: %s",
                    clean_message_id
                )

                # ====================================================
                # ATTACH INLINE IMAGES
                # ====================================================

                images = {
                    "logo": "static/logo.png",
                    "regards": "static/regards.png",
                    "awards": "static/awards.png",
                    "facebook": "static/facebook.png",
                    "instagram": "static/instagram.png",
                    "linkedin": "static/linkedin.png",
                    "youtube": "static/youtube.png",
                    "pinterest": "static/pinterest.png",
                    "x": "static/x.png"
                }

                for cid, path in images.items():

                    logger.debug(
                        "Attaching image: %s",
                        path
                    )

                    if not os.path.exists(path):

                        logger.error(
                            "❌ IMAGE FILE DOES NOT EXIST: %s",
                            path
                        )

                        raise FileNotFoundError(
                            f"Image not found: {path}"
                        )

                    with open(
                        path,
                        "rb"
                    ) as f:

                        img = MIMEImage(
                            f.read()
                        )

                        img.add_header(
                            "Content-ID",
                            f"<{cid}>"
                        )

                        img.add_header(
                            "Content-Disposition",
                            "inline",
                            filename=os.path.basename(path)
                        )

                        message.attach(
                            img
                        )

                logger.info(
                    "All inline images attached"
                )

                # ====================================================
                # SECOND LOCK
                #
                # YOUR ORIGINAL CODE HAS THIS SECOND LOCK.
                # LOGGING WILL SHOW IF THIS IS THE PART BLOCKING SEND.
                # ====================================================

                logger.info(
                    "Attempting email_sending lock for %s",
                    receiver
                )

                lock = collection.update_one(
                    {
                        "_id": person["_id"],
                        "email_sent": {
                            "$ne": True
                        },
                        "email_sending": {
                            "$ne": True
                        }
                    },
                    {
                        "$set": {
                            "email_sending": True
                        }
                    }
                )

                logger.info(
                    "email_sending lock result: matched=%s modified=%s",
                    lock.matched_count,
                    lock.modified_count
                )

                if lock.modified_count == 0:

                    logger.error(
                        "❌ SECOND LOCK FAILED - EMAIL WILL NOT BE SENT: %s",
                        receiver
                    )

                    latest = collection.find_one(
                        {"_id": person["_id"]}
                    )

                    logger.error(
                        "Mongo state after failed second lock: %s",
                        {
                            "email_sent": latest.get("email_sent") if latest else None,
                            "email_sending": latest.get("email_sending") if latest else None,
                            "sending_initial": latest.get("sending_initial") if latest else None
                        }
                    )

                    # IMPORTANT:
                    # Release first lock
                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_initial": False
                            }
                        }
                    )

                    continue

                logger.info(
                    "🔒 email_sending LOCK ACQUIRED: %s",
                    receiver
                )

                # ====================================================
                # CHECK IST EMAIL SENDING WINDOW
                # ====================================================

                if not is_email_sending_window():

                    next_window = get_next_email_window_start()

                    queued_for = queue_initial_email(
                        person["_id"],
                        receiver,
                        subject,
                        body,
                        signature_html,
                        next_window,
                        reason="bulk_window_closed_mid_batch_next_weekday_9am"
                    )

                    logger.info(
                        "⏸ INITIAL EMAIL QUEUED: %s -> %s",
                        receiver,
                        queued_for
                    )

                    results.append({
                        "email": receiver,
                        "status": "queued",
                        "queued_for": queued_for
                    })

                    continue

                # ====================================================
                # GLOBAL 30-SECOND OUTBOUND GAP
                # ====================================================

                wait_for_outbound_email_gap()

                if not is_email_sending_window():
                    next_window = get_next_email_window_start()
                    queued_for = queue_initial_email(
                        person["_id"],
                        receiver,
                        subject,
                        body,
                        signature_html,
                        next_window,
                        reason="bulk_window_closed_after_gap_wait_next_weekday_9am"
                    )

                    logger.info(
                        "⏸ INITIAL EMAIL RE-QUEUED AFTER GAP: %s -> %s",
                        receiver,
                        queued_for
                    )

                    results.append({
                        "email": receiver,
                        "status": "queued",
                        "queued_for": queued_for
                    })
                    continue

                # ====================================================
                # SMTP SEND
                # ====================================================

                logger.info(
                    "📤 ABOUT TO SEND EMAIL - EMAIL WINDOW OPEN"
                )

                logger.info(
                    "SMTP FROM: %s",
                    SENDER_EMAIL
                )

                logger.info(
                    "SMTP TO: %s",
                    receiver
                )

                logger.info(
                    "SMTP SUBJECT: %s",
                    subject
                )

                send_result = server.sendmail(
                    SENDER_EMAIL,
                    [receiver],
                    message.as_string()
                )

                logger.info(
                    "SMTP sendmail() returned: %s",
                    send_result
                )

                # IMPORTANT: smtplib may return a dictionary of refused
                # recipients instead of raising an exception. Treat any
                # refusal as an undeliverable/failed email and DO NOT mark
                # the prospect as sent.
                if send_result:
                    refusal_details = "; ".join(
                        f"{addr}: {err!r}"
                        for addr, err in send_result.items()
                    )
                    raise smtplib.SMTPRecipientsRefused(send_result)

                logger.info(
                    "✅ SMTP SEND COMPLETED FOR: %s",
                    receiver
                )

                # ====================================================
                # UPDATE DATABASE
                # ====================================================

                current_time = datetime.now(
                    IST
                ).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                first_followup_due_at = (
                    get_first_followup_send_time(current_time)
                    .strftime("%Y-%m-%d %H:%M:%S")
                )

                update_result = collection.update_one(
                    {
                        "_id": person["_id"]
                    },
                    {
                        "$set": {
                            "email_sent": True,

                            "email_sending": False,

                            "sending_initial": False,

                            "email_sent_date":
                                current_time,

                            "last_outbound_sent_date":
                                current_time,

                            "email_delivery_status":
                                "sent",

                            "last_email_subject":
                                subject,

                            "last_email_body":
                                body,

                            "email_signature":
                                signature_html,

                            "message_id":
                                clean_message_id,

                            "thread_message_id":
                                clean_message_id,

                            "last_message_id":
                                clean_message_id,

                            "thread_type":
                                "Same Thread",

                            "followup_count":
                                0,

                            "followup_sent":
                                False,

                            "followup_sent_date":
                                "",

                            "last_followup_date":
                                "",

                            "followup_due_at":
                                first_followup_due_at,

                            "followup_queue_status":
                                "scheduled",

                            "followup_queue_number":
                                1,

                            "followup_queued_for":
                                first_followup_due_at,

                            "campaign_status":
                                "Running",

                            "has_reply":
                                False,

                            "active":
                                True,

                            "followup_cancel_reason":
                                "",

                            "reply_subject":
                                "",

                            "reply_received_date":
                                "",

                            "email_error":
                                "",

                            "email_retry_after":
                                None,

                            "email_thread": [
                                {
                                    "type": "initial",
                                    "number": 0,
                                    "subject": subject,
                                    "body": body,
                                    "signature": signature_html,
                                    "message_id": clean_message_id,
                                    "date": current_time
                                }
                            ]
                        }
                    }
                )

                # Keep the configured gap between bulk emails.
                # The next email starts only after this 30-second delay.
                if i < len(recipients) - 1 and is_email_sending_window():
                    logger.info(
                        "⏳ Waiting %s seconds before next bulk email",
                        EMAIL_INTERVAL_SECONDS
                    )
                    time.sleep(EMAIL_INTERVAL_SECONDS)

                logger.info(
                    "MongoDB update after send: matched=%s modified=%s",
                    update_result.matched_count,
                    update_result.modified_count
                )

                # ====================================================
                # SUCCESS
                # ====================================================

                sent_count += 1

                results.append({
                    "email": receiver,
                    "status": "sent"
                })

                logger.info(
                    "🎉 EMAIL SUCCESSFULLY SENT: %s",
                    receiver
                )

            except Exception as e:

                failed_count += 1

                logger.exception(
                    "❌ EMAIL FAILED FOR %s",
                    receiver
                )

                current_time = datetime.now(
                    IST
                ).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                try:

                    is_undeliverable = isinstance(
                        e,
                        smtplib.SMTPRecipientsRefused
                    )

                    collection.update_one(
                        {
                            "email": receiver
                        },
                        {
                            "$set": {
                                "email_sent": False,

                                "email_delivery_status":
                                    "undeliverable" if is_undeliverable else "failed",

                                "email_failure_type":
                                    "undeliverable" if is_undeliverable else "delivery_error",

                                "undeliverable_address":
                                    receiver if is_undeliverable else "",

                                "email_error":
                                    (
                                        "Recipient refused by SMTP: " +
                                        "; ".join(
                                            f"{addr}: {err!r}"
                                            for addr, err in getattr(
                                                e,
                                                "recipients",
                                                {}
                                            ).items()
                                        )
                                    )
                                    if is_undeliverable and getattr(e, "recipients", None)
                                    else str(e),

                                "email_failed_date":
                                    current_time,

                                "email_retry_after":
                                    get_next_retry_time().strftime("%Y-%m-%d %H:%M:%S"),

                                "campaign_status":
                                    "Undeliverable" if is_undeliverable else "Delivery Failed",

                                "sending_initial":
                                    False,

                                "email_sending":
                                    False
                            }
                        }
                    )

                    logger.info(
                        "MongoDB failure state saved for %s",
                        receiver
                    )

                except Exception as db_error:

                    logger.exception(
                        "❌ FAILED TO UPDATE ERROR STATE: %s",
                        db_error
                    )

                results.append({
                    "email": receiver,
                    "status": (
                        "undeliverable"
                        if isinstance(e, smtplib.SMTPRecipientsRefused)
                        else "failed"
                    ),
                    "error": str(e)
                })

                # One failure must not stop the batch: if the SMTP session
                # dropped with it, reconnect before the next recipient.
                try:
                    server.noop()
                except Exception:
                    logger.warning("SMTP session lost after failure for %s; reconnecting", receiver)
                    try:
                        server.close()
                    except Exception:
                        pass
                    try:
                        server = _open_smtp_connection()
                    except Exception:
                        logger.exception("SMTP reconnect failed; remaining recipients will be retried individually")

        # ============================================================
        # CLOSE SMTP
        # ============================================================

        if server:

            try:

                server.quit()

                logger.info(
                    "SMTP connection closed"
                )

            except Exception as e:

                logger.warning(
                    "SMTP quit error: %s",
                    e
                )

        logger.info("=" * 80)

        logger.info(
            "SEND EMAIL FINISHED | sent=%s | failed=%s",
            sent_count,
            failed_count
        )

        logger.info("=" * 80)

        return JSONResponse(
            content={
                "success": True,
                "sent": sent_count,
                "failed": failed_count,
                "results": results,
                "message": (
                    f"Successfully sent {sent_count} emails. "
                    f"Failed: {failed_count}"
                )
            }
        )

    except Exception as e:

        logger.exception(
            "❌ GLOBAL SEND EMAIL ERROR"
        )

        if server:

            try:
                server.quit()
            except Exception:
                pass

        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": str(e)
            }
        )


def _clean_graph_message_text(value):
    """Normalize Graph HTML/text content so NDR phrases can be detected reliably."""
    text = str(value or "")
    text = re.sub(r"<[^>]+>", " ", text)
    text = (
        text.replace("&nbsp;", " ")
            .replace("&amp;", "&")
            .replace("&#39;", "'")
            .replace("&quot;", '"')
    )
    text = text.replace("’", "'").replace("`", "'")
    return re.sub(r"\s+", " ", text).strip()


def _is_undeliverable_notification(msg, message_text):
    """Return True for common Outlook/Microsoft 365 non-delivery notifications."""
    subject = _clean_graph_message_text(msg.get("subject", "")).lower()
    text = _clean_graph_message_text(message_text).lower()

    subject_markers = (
        "undeliverable",
        "delivery status notification",
        "delivery has failed",
        "mail delivery failed"
    )

    body_markers = (
        "couldn't be delivered",
        "could not be delivered",
        "wasn't found at",
        "was not found at",
        "recipient not found",
        "user unknown",
        "mailbox unavailable",
        "550 5.1.1",
        "550 5.1.10",
        "delivery has failed"
    )

    if any(marker in subject for marker in subject_markers):
        return True

    return sum(marker in text for marker in body_markers) >= 1


def _extract_undeliverable_recipient(message_text):
    """Extract the recipient address from an Outlook NDR."""
    text = _clean_graph_message_text(message_text)

    patterns = [
        r"Your message to\s*<?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})>?\s*(?:couldn't|could not)\s+be\s+delivered",
        r"(?:Original-Recipient|Final-Recipient):\s*(?:rfc822;)?\s*<?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})>?",
        r"(?:recipient|to)\s*[:\-]?\s*<?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})>?"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip().lower()

    # Fallback: choose an email address that actually exists in our prospects.
    candidates = []
    for candidate in re.findall(
        r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b",
        text
    ):
        candidate = candidate.lower()
        if candidate not in candidates and candidate != SENDER_EMAIL.lower():
            candidates.append(candidate)

    for candidate in candidates:
        if collection.find_one({"email": candidate}, {"_id": 1}):
            return candidate

    return candidates[0] if candidates else ""


def _build_undeliverable_error(recipient, message_text):
    """Create a short dashboard-friendly NDR message."""
    text = _clean_graph_message_text(message_text)

    match = re.search(
        r"(Your message to\s*<?"
        + re.escape(recipient)
        + r">?\s*(?:couldn't|could not)\s+be\s+delivered\."
        r"(?:\s+[^.?!]{1,240}[.?!])?)",
        text,
        re.IGNORECASE
    )

    if match:
        return match.group(1).strip()

    return (
        f"Your message to {recipient} couldn't be delivered. "
        f"{text[:450]}"
    ).strip()


def _process_outlook_undeliverable(msg, message_text):
    """
    Detect an Outlook NDR and mark the corresponding prospect as
    undeliverable. Returns True when the message is an NDR.
    """
    if not _is_undeliverable_notification(msg, message_text):
        return False

    recipient = _extract_undeliverable_recipient(message_text)
    if not recipient:
        logger.warning(
            "Outlook NDR detected but recipient could not be identified. Subject=%s",
            msg.get("subject", "")
        )
        return True

    prospect = collection.find_one(
        {"email": recipient},
        {
            "_id": 1,
            "email": 1,
            "email_sent": 1,
            "last_message_id": 1
        }
    )

    if not prospect:
        logger.warning(
            "Outlook NDR recipient %s was not found in prospects.",
            recipient
        )
        return True

    notification_id = msg.get("id", "")
    existing_notification_id = collection.find_one(
        {"_id": prospect["_id"]},
        {"undeliverable_notification_id": 1}
    )
    if (
        notification_id
        and existing_notification_id
        and existing_notification_id.get("undeliverable_notification_id") == notification_id
    ):
        return True

    received_at = msg.get("receivedDateTime") or ""
    try:
        failed_dt = (
            datetime.fromisoformat(
                received_at.replace("Z", "+00:00")
            )
            .astimezone(IST)
            .strftime("%Y-%m-%d %H:%M:%S")
        )
    except Exception:
        failed_dt = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

    error_message = _build_undeliverable_error(
        recipient,
        message_text
    )

    collection.update_one(
        {"_id": prospect["_id"]},
        {
            "$set": {
                "email_sent": False,
                "email_delivery_status": "undeliverable",
                "email_failure_type": "undeliverable",
                "undeliverable_address": recipient,
                "email_error": error_message,
                "email_failed_date": failed_dt,
                "undeliverable_at": failed_dt,
                "undeliverable_notification_id": notification_id,
                "email_retry_after": None,
                "campaign_status": "Undeliverable",
                "sending_initial": False,
                "email_sending": False
            }
        }
    )

    logger.warning(
        "OUTLOOK NDR: %s marked as undeliverable. %s",
        recipient,
        error_message
    )

    return True



def _graph_get_me(token):
    """Return the signed-in Microsoft Graph user's identity."""
    response = requests.get(
        f"{TEAMS_GRAPH_BASE_URL}/me?$select=id,displayName,userPrincipalName,mail",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        timeout=30,
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"Graph /me failed (HTTP {response.status_code}): {response.text[:1200]}"
        )
    return response.json()


def _find_teams_self_chat(token):
    """
    Find the signed-in user's actual Teams self-chat, the chat shown in Teams
    with '(You)' in its display name. We discover it through Graph /me/chats
    instead of hard-coding a chat id.
    """
    me = _graph_get_me(token)
    my_id = str(me.get("id") or "").strip()
    if not my_id:
        raise RuntimeError("Graph /me did not return the signed-in user's id.")

    url = f"{TEAMS_GRAPH_BASE_URL}/me/chats?$expand=members&$top=50"
    pages_checked = 0
    candidates = []

    while url:
        response = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            timeout=30,
        )
        if response.status_code != 200:
            raise RuntimeError(
                f"Teams chat list failed (HTTP {response.status_code}): "
                f"{response.text[:1500]}"
            )

        payload = response.json()
        chats = payload.get("value", [])
        pages_checked += 1

        for chat in chats:
            if str(chat.get("chatType") or "").lower() != "oneonone":
                continue

            members = chat.get("members") or []
            user_members = []
            for member in members:
                if str(member.get("@odata.type") or "").lower().endswith("aaduserconversationmember"):
                    user = member.get("user") or {}
                    user_members.append({
                        "id": str(user.get("id") or ""),
                        "displayName": str(user.get("displayName") or ""),
                        "email": str(
                            user.get("email")
                            or user.get("userPrincipalName")
                            or member.get("email")
                            or ""
                        ).lower(),
                    })

            # A Teams self-chat is represented as a one-on-one chat involving
            # the signed-in user itself. The exact number of returned member
            # objects can vary, so match by user identity rather than assuming
            # a fixed member count.
            if any(member.get("id") == my_id for member in user_members):
                display_names = " / ".join(
                    item.get("displayName") or item.get("email") or item.get("id")
                    for item in user_members
                )
                candidates.append((chat.get("id"), display_names, len(user_members)))

        url = payload.get("@odata.nextLink")

    # Prefer a candidate whose only/primary user is the signed-in account.
    exact_self = [
        candidate for candidate in candidates
        if candidate[2] <= 1
    ]
    if exact_self:
        chat_id, display_names, _ = exact_self[0]
        logger.info(
            "Teams self-chat discovered: id=%s participants=%s",
            chat_id,
            display_names,
        )
        return chat_id

    if candidates:
        chat_id, display_names, _ = candidates[0]
        logger.info(
            "Teams one-on-one self-chat candidate discovered: id=%s participants=%s",
            chat_id,
            display_names,
        )
        return chat_id

    raise RuntimeError(
        "Teams self-chat for the signed-in user was not found in Graph /me/chats. "
        "Open Teams and make sure the chat is the real self-chat named '(You)'."
    )


def _mark_teams_chat_unread(token, chat_id):
    """Mark the Teams self-chat unread for the signed-in user."""
    me = _graph_get_me(token)
    user_id = str(me.get("id") or "").strip()
    if not user_id:
        raise RuntimeError("Unable to determine signed-in Teams user id.")

    response = requests.post(
        f"{TEAMS_GRAPH_BASE_URL}/chats/{chat_id}/markChatUnreadForUser",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        json={
            "user": {
                "id": user_id,
                "tenantId": TENANT_ID,
            }
        },
        timeout=30,
    )

    if response.status_code != 204:
        raise RuntimeError(
            f"Teams mark-unread failed (HTTP {response.status_code}): "
            f"{response.text[:1200]}"
        )


def send_teams_reply_notification(
    token,
    sender,
    subject,
    reply_received,
    company,
    prospect_name,
    prospect_email,
    followup_due_at,
    reply_preview,
):
    """Post a customer-reply alert into the Rameshwar Teams self-chat."""
    chat_id = _find_teams_self_chat(token)

    safe_sender = escape(sender or "-")
    safe_subject = escape(subject or "-")
    safe_received = escape(reply_received or "-")
    safe_company = escape(company or "-")
    safe_name = escape(prospect_name or "-")
    safe_email = escape(prospect_email or "-")
    safe_due = escape(followup_due_at or "-")
    safe_preview = escape((reply_preview or "-").strip())

    content = (
        "<strong>🔔 Valasys Apollo — Customer Replied</strong>"
        f"<br><br><strong>Prospect:</strong> {safe_name}"
        f"<br><strong>Company:</strong> {safe_company}"
        f"<br><strong>Email:</strong> {safe_email}"
        f"<br><strong>Reply From:</strong> {safe_sender}"
        f"<br><strong>Subject:</strong> {safe_subject}"
        f"<br><strong>Received:</strong> {safe_received}"
        f"<br><strong>Follow-up #1:</strong> {safe_due}"
        f"<br><br><strong>Reply:</strong><br>{safe_preview}"
    )

    response = requests.post(
        f"{TEAMS_GRAPH_BASE_URL}/chats/{chat_id}/messages",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        json={
            "importance": "urgent",
            "body": {
                "contentType": "html",
                "content": content,
            },
        },
        timeout=30,
    )

    if response.status_code not in (200, 201):
        raise RuntimeError(
            f"Teams message send failed (HTTP {response.status_code}): "
            f"{response.text[:1500]}"
        )

    message_data = response.json() if response.content else {}
    teams_message_id = str(message_data.get("id") or "")

    # Mark the self-chat unread so it visibly appears as new/unread in Teams.
    unread_warning = ""
    try:
        _mark_teams_chat_unread(token, chat_id)
    except Exception as unread_exc:
        unread_warning = str(unread_exc)
        logger.warning(
            "Teams message sent but self-chat could not be marked unread: %s",
            unread_exc,
        )

    return {
        "chat_id": chat_id,
        "message_id": teams_message_id,
        "marked_unread": not bool(unread_warning),
        "unread_error": unread_warning,
    }


def _normalize_subject_for_match(value):
    """Normalize an Outlook subject for safe reply-thread matching."""
    text = _clean_graph_message_text(value).lower().strip()
    # Remove repeated reply/forward prefixes.
    while True:
        updated = re.sub(r"^(re|fw|fwd)\s*:\s*", "", text, flags=re.IGNORECASE).strip()
        if updated == text:
            break
        text = updated
    return re.sub(r"\s+", " ", text)


def _message_id_tokens(value):
    """Return normalized Message-ID tokens from a References/In-Reply-To value."""
    text = str(value or "")
    tokens = re.findall(r"<([^>]+)>", text)
    if not tokens:
        tokens = re.split(r"\s+", text.strip()) if text.strip() else []
    return [token.strip().strip("<>") for token in tokens if token.strip().strip("<>")]


def _find_prospect_for_reply(sender, in_reply_to, references, subject):
    """Robustly map an incoming Outlook reply to the stored outbound campaign record."""
    sender = (sender or "").strip().lower()

    # 1) Exact In-Reply-To Message-ID match across ALL sent prospect records.
    in_reply_to_clean = str(in_reply_to or "").strip().strip("<>")
    if in_reply_to_clean:
        existing = collection.find_one({
            "email_sent": True,
            "$or": [
                {"thread_message_id": in_reply_to_clean},
                {"last_message_id": in_reply_to_clean},
                {"message_id": in_reply_to_clean},
                {"email_thread.message_id": in_reply_to_clean},
            ]
        })
        if existing:
            return existing, "in-reply-to"

    # 2) References header can contain multiple outbound Message-IDs.
    reference_ids = _message_id_tokens(references)
    if reference_ids:
        existing = collection.find_one({
            "email_sent": True,
            "$or": [
                {"thread_message_id": {"$in": reference_ids}},
                {"last_message_id": {"$in": reference_ids}},
                {"message_id": {"$in": reference_ids}},
                {"email_thread.message_id": {"$in": reference_ids}},
            ]
        })
        if existing:
            return existing, "references"

    # 3) Exact sender email match, regardless of current campaign status.
    # This is important for replies received after a follow-up or after a legacy
    # record was left with an outdated status.
    if sender:
        existing = collection.find_one({
            "email": sender,
            "email_sent": True,
        })
        if existing:
            return existing, "sender-email"

    # 4) Subject fallback for aliases/forwarded replies where the sender address
    # differs from the stored prospect address. Prefer the newest sent record.
    normalized_subject = _normalize_subject_for_match(subject)
    if normalized_subject:
        candidates = collection.find(
            {
                "email_sent": True,
                "last_email_subject": {"$exists": True, "$nin": ["", None]},
            },
            {
                "_id": 1,
                "email": 1,
                "last_email_subject": 1,
                "email_sent_date": 1,
                "first_name": 1,
                "last_name": 1,
                "company": 1,
                "campaign_status": 1,
                "has_reply": 1,
                "last_reply_message_id": 1,
                "teams_reply_notification_message_id": 1,
                "followup_queue_status": 1,
                "followup_due_at": 1,
                "followup_queued_for": 1,
            }
        ).sort("email_sent_date", -1)

        for prospect in candidates:
            candidate_subject = _normalize_subject_for_match(
                prospect.get("last_email_subject", "")
            )
            if candidate_subject == normalized_subject:
                return prospect, "subject"

    return None, "unmatched"


def check_replies():

    try:
        token = get_access_token()

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

        # Check the last 7 days so older replies are repaired as well.
        since = (
            datetime.now(timezone.utc)
            - timedelta(days=7)
        ).strftime("%Y-%m-%dT%H:%M:%SZ")

        # IMPORTANT: order newest-first and follow @odata.nextLink. Without
        # pagination, an inbox with more than 100 messages can hide the reply.
        url = (
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
            f"?$filter=receivedDateTime ge {since}"
            "&$select=id,subject,receivedDateTime,internetMessageHeaders,from,bodyPreview"
            "&$orderby=receivedDateTime desc"
            "&$top=100"
        )

        pages = 0
        processed_messages = 0

        while url:
            pages += 1
            if pages > 20:
                logger.warning("Stopping reply pagination after 20 Outlook pages.")
                break

            response = requests.get(
                url,
                headers=headers,
                timeout=30,
            )

            if response.status_code != 200:
                logger.warning(
                    "Outlook reply check failed (HTTP %s): %s",
                    response.status_code,
                    response.text[:1200],
                )
                return

            payload = response.json()
            messages = payload.get("value", [])
            logger.info(
                "Outlook reply sync page %s: %s message(s)",
                pages,
                len(messages),
            )

            for msg in messages:
                processed_messages += 1
                message_id = msg.get("id", "")

                sender = (
                    msg.get("from", {})
                    .get("emailAddress", {})
                    .get("address", "")
                    .strip()
                    .lower()
                )

                if not sender or sender == SENDER_EMAIL.lower():
                    continue

                internet_headers = msg.get("internetMessageHeaders") or []
                if not internet_headers:
                    # Fallback only when Graph did not include headers in the list response.
                    try:
                        header_response = requests.get(
                            f"https://graph.microsoft.com/v1.0/me/messages/{message_id}?$select=internetMessageHeaders",
                            headers=headers,
                            timeout=30,
                        )
                        if header_response.status_code == 200:
                            internet_headers = header_response.json().get(
                                "internetMessageHeaders",
                                []
                            )
                    except Exception as header_error:
                        logger.warning(
                            "Failed to fetch headers for Outlook message %s: %s",
                            message_id,
                            header_error,
                        )

                in_reply_to = ""
                references = ""

                for h in internet_headers:
                    name = str(h.get("name", "")).lower()
                    if name == "in-reply-to":
                        in_reply_to = str(h.get("value", "")).strip()
                    elif name == "references":
                        references = str(h.get("value", "")).strip()

                message_text = msg.get("bodyPreview") or ""
                subject = msg.get("subject", "") or ""

                subject_text = _clean_graph_message_text(subject).lower()
                if (
                    not _is_undeliverable_notification(msg, message_text)
                    and any(
                        marker in subject_text
                        for marker in (
                            "undeliverable",
                            "delivery status notification",
                            "delivery has failed",
                            "mail delivery failed",
                        )
                    )
                ):
                    try:
                        body_response = requests.get(
                            f"https://graph.microsoft.com/v1.0/me/messages/{message_id}?$select=subject,body,bodyPreview",
                            headers=headers,
                            timeout=30,
                        )
                        if body_response.status_code == 200:
                            body_data = body_response.json()
                            message_text = (
                                body_data.get("body", {}).get("content")
                                or body_data.get("bodyPreview")
                                or message_text
                            )
                    except Exception as body_error:
                        logger.warning(
                            "Failed to fetch full Outlook NDR body for %s: %s",
                            message_id,
                            body_error,
                        )

                if _process_outlook_undeliverable(msg, message_text):
                    continue

                existing, match_method = _find_prospect_for_reply(
                    sender=sender,
                    in_reply_to=in_reply_to,
                    references=references,
                    subject=subject,
                )

                if not existing:
                    logger.warning(
                        "Outlook reply detected but no prospect matched | sender=%s | subject=%s | in_reply_to=%s | references=%s | message_id=%s",
                        sender,
                        subject,
                        in_reply_to,
                        references,
                        message_id,
                    )
                    continue

                logger.info(
                    "Outlook reply matched prospect | sender=%s | prospect_email=%s | method=%s | message_id=%s",
                    sender,
                    existing.get("email", ""),
                    match_method,
                    message_id,
                )

                incoming_reply_id = message_id
                previous_reply_id = existing.get("last_reply_message_id", "")
                teams_notified_reply_id = existing.get(
                    "teams_reply_notification_message_id", ""
                )

                reply_received = msg.get("receivedDateTime") or ""
                is_new_reply = not (
                    incoming_reply_id
                    and incoming_reply_id == previous_reply_id
                )

                # Always repair the database state whenever a reply is found,
                # even when the same reply was already processed previously.
                # This guarantees Sent Emails never continues to show a follow-up
                # after a customer reply.
                if (
                    existing.get("campaign_status") != "Replied"
                    or existing.get("has_reply") is not True
                    or existing.get("followup_queue_status") == "scheduled"
                    or existing.get("followup_due_at")
                    or existing.get("followup_queued_for")
                    or existing.get("active") is not False
                ):
                    collection.update_one(
                        {"_id": existing["_id"]},
                        {
                            "$set": {
                                "has_reply": True,
                                "campaign_status": "Replied",
                                "reply_subject": subject,
                                "reply_received_date": reply_received,
                                "last_reply_message_id": incoming_reply_id,
                                "followup_due_at": "",
                                "followup_queue_status": "cancelled",
                                "followup_queue_number": 0,
                                "followup_queued_for": "",
                                "sending_followup": False,
                                "active": False,
                                "followup_cancel_reason": "customer_replied",
                                "followup_next_date": "",
                            }
                        }
                    )
                    print(
                        "🔄 REPLY STATUS UPDATED",
                        f"| {sender} -> Replied",
                        "| FOLLOW-UPS STOPPED",
                        f"| match={match_method}",
                    )

                # Teams notification is only needed for a new Outlook reply,
                # or when a previous Teams notification failed.
                if is_new_reply or teams_notified_reply_id != incoming_reply_id:
                    try:
                        notification = send_teams_reply_notification(
                            token=token,
                            sender=sender,
                            subject=subject,
                            reply_received=reply_received,
                            company=existing.get("company", ""),
                            prospect_name=existing.get("full_name")
                            or " ".join(
                                part for part in (
                                    existing.get("first_name", ""),
                                    existing.get("last_name", ""),
                                )
                                if part
                            ),
                            prospect_email=existing.get("email") or sender,
                            followup_due_at="STOPPED - customer replied",
                            reply_preview=message_text,
                        )

                        collection.update_one(
                            {"_id": existing["_id"]},
                            {
                                "$set": {
                                    "teams_reply_notification_message_id": incoming_reply_id,
                                    "teams_reply_notification_status": "sent",
                                    "teams_reply_notification_sent_at": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                                    "teams_reply_notification_error": "",
                                    "teams_reply_chat_id": notification.get("chat_id", ""),
                                    "teams_reply_teams_message_id": notification.get("message_id", ""),
                                    "teams_reply_chat_marked_unread": notification.get("marked_unread", False),
                                }
                            }
                        )

                        print(
                            "🔔 Teams customer-reply notification sent",
                            f"| chat={notification.get('chat_id', '')}",
                            f"| unread={notification.get('marked_unread', False)}",
                            f"| reply={sender}",
                        )

                    except Exception as teams_error:
                        collection.update_one(
                            {"_id": existing["_id"]},
                            {
                                "$set": {
                                    "teams_reply_notification_status": "failed",
                                    "teams_reply_notification_error": str(teams_error),
                                }
                            }
                        )
                        logger.exception(
                            "Teams customer-reply notification failed for %s",
                            sender,
                        )
                        print(
                            "❌ Teams customer-reply notification failed:",
                            str(teams_error),
                        )

                if is_new_reply:
                    print(
                        "✅ Customer reply processed:",
                        sender,
                        "| follow-ups stopped | no follow-up date",
                    )

            url = payload.get("@odata.nextLink")

        logger.info(
            "Outlook reply sync finished | pages=%s | messages=%s",
            pages,
            processed_messages,
        )

    except Exception as e:
        logger.exception("Reply checker failed: %s", e)
        print("Reply checker error:", e)


def clear_stale_initial_locks():

    try:

        cutoff = datetime.now(IST) - timedelta(minutes=30)

        cutoff_string = cutoff.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        result = collection.update_many(
            {
                "sending_initial": True,
                "initial_send_started_at": {
                    "$lt": cutoff_string
                }
            },
            {
                "$set": {
                    "sending_initial": False
                }
            }
        )

        if result.modified_count:
            print(
                f"🔓 Cleared {result.modified_count} stale initial-email locks"
            )

    except Exception as e:

        print(
            "Stale initial lock cleanup error:",
            e
        )

def send_followups():

    print("=" * 70)
    print("send_followups() called:", datetime.now(IST))

    server = None

    try:

        # ============================================================
        # SMTP CONNECTION
        # ============================================================

        server = smtplib.SMTP(
            SMTP_SERVER,
            SMTP_PORT,
            timeout=60
        )

        server.starttls(
            context=ssl.create_default_context()
        )

        server.login(
            SENDER_EMAIL,
            PASSWORD
        )

        now = datetime.now(IST)

        # ============================================================
        # IMPORTANT:
        # Process ONLY 5 profiles per scheduler execution.
        #
        # This prevents the system from trying to process
        # hundreds/thousands of profiles in one execution.
        # ============================================================

        MAX_PROFILES_PER_RUN = 5

        now = datetime.now(IST)

        # The due time is already calculated from the initial email or the
        # previous follow-up send. Follow-ups are sent only while there is NO reply.
        cutoff_string = now.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        print(
            f"Looking for profiles due before or at: "
            f"{cutoff_string} IST"
        )

        # Refresh replies immediately before selecting due follow-ups.
        # This prevents the 10-second follow-up worker from relying only on
        # the periodic reply-check scheduler.
        print("🔎 Refreshing customer replies before selecting follow-ups...")
        check_replies()

        # ============================================================
        # FOLLOW-UP 1-5 WHEN THERE IS NO CUSTOMER REPLY
        #
        # Every follow-up is due 5 minutes after the previous outbound email.
        # ============================================================

        followup_query = {
            "email_sent": True,
            "email_delivery_status": "sent",
            "has_reply": False,
            "campaign_status": "Running",
            "followup_queue_status": "scheduled",
            "sending_followup": {
                "$ne": True
            },
            "followup_count": {
                "$lt": 5
            },
            "followup_due_at": {
                "$exists": True,
                "$nin": ["", None],
                "$lte": cutoff_string
            }
        }

        # ============================================================
        # GET ONLY DUE PROFILES
        # ============================================================

        prospects = list(
            collection.find(followup_query)
            .sort("followup_due_at", 1)
            .limit(MAX_PROFILES_PER_RUN)
        )

        print(
            f"Due no-reply follow-ups selected: {len(prospects)}"
        )

        print(
            f"Due profiles selected: {len(prospects)}"
        )

        if not prospects:

            print(
                "No follow-ups currently eligible."
            )

            return

        # ============================================================
        # PROCESS MAXIMUM 5 PROFILES
        # ============================================================

        for followup_index, person in enumerate(prospects):

            receiver = person.get("email", "")

            if not receiver:
                print(
                    "⚠ Skipping profile without email"
                )
                continue

            try:

                print("=" * 70)
                print(
                    f"Processing follow-up for: {receiver}"
                )

                # ====================================================
                # ATOMIC LOCK
                # ====================================================

                lock_result = collection.update_one(
                    {
                        "_id": person["_id"],

                        "campaign_status": "Running",

                        "has_reply": False,

                        "followup_queue_status": "scheduled",

                        "sending_followup": {
                            "$ne": True
                        }
                    },
                    {
                        "$set": {
                            "sending_followup": True
                        }
                    }
                )

                if lock_result.modified_count == 0:

                    print(
                        f"⚠ Already being processed: {receiver}"
                    )

                    continue

                # Re-check Outlook after claiming the row. A reply may have
                # arrived between the initial query and this send attempt.
                check_replies()
                latest_state = collection.find_one(
                    {"_id": person["_id"]},
                    {"has_reply": 1, "campaign_status": 1}
                )

                if (
                    not latest_state
                    or latest_state.get("has_reply") is True
                    or latest_state.get("campaign_status") != "Running"
                ):
                    collection.update_one(
                        {"_id": person["_id"]},
                        {
                            "$set": {
                                "campaign_status": (
                                    "Replied"
                                    if latest_state and latest_state.get("has_reply") is True
                                    else latest_state.get("campaign_status", "Running")
                                    if latest_state
                                    else "Replied"
                                ),
                                "sending_followup": False,
                                "followup_queue_status": "cancelled"
                            }
                        }
                    )
                    print(
                        f"🛑 Follow-up stopped for {receiver}: customer replied or campaign is no longer Running."
                    )
                    continue

                # ====================================================
                # DETERMINE FOLLOW-UP NUMBER
                # ====================================================

                followup_count = person.get(
                    "followup_count",
                    0
                )

                if followup_count >= 5:

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "campaign_status":
                                    "Completed - Follow-ups Sent",

                                "active": False,

                                "sending_followup": False
                            }
                        }
                    )

                    print(
                        f"Campaign already completed: {receiver}"
                    )

                    continue

                followup_number = followup_count + 1

                # ----------------------------------------------------
                # Ensure a scheduled due time exists.
                # This repairs older "Replied" records created before
                # followup_due_at was stored.
                # ----------------------------------------------------
                due_text = person.get("followup_due_at")

                if not due_text:
                    recovery_base = (
                        person.get("email_sent_date")
                        if followup_count == 0
                        else person.get("last_followup_date")
                    )

                    if recovery_base:
                        try:
                            recovery_due = (
                                get_first_followup_send_time(recovery_base)
                                if followup_count == 0
                                else get_next_followup_queue_slot(
                                    recovery_base,
                                    followup_number
                                )
                            )
                            due_text = recovery_due.strftime(
                                "%Y-%m-%d %H:%M:%S"
                            )

                            collection.update_one(
                                {"_id": person["_id"]},
                                {
                                    "$set": {
                                        "followup_due_at": due_text,
                                        "followup_queue_status": "scheduled",
                                        "followup_queue_number": followup_number,
                                        "followup_queued_for": due_text
                                    }
                                }
                            )

                            person["followup_due_at"] = due_text

                            print(
                                f"🧰 Repaired missing Follow-up #{followup_number} "
                                f"schedule for {receiver}: {due_text}"
                            )

                        except Exception as repair_error:
                            print(
                                f"⚠ Unable to repair follow-up schedule for "
                                f"{receiver}: {repair_error}"
                            )

                # ====================================================
                # DETERMINE LAST EMAIL SENT DATE
                #
                # First follow-up:
                #     initial email date
                #
                # Later follow-ups:
                #     last follow-up date
                #
                # Therefore every follow-up is calculated from the
                # LAST ACTUALLY SENT EMAIL, only while there is no reply.
                # ====================================================

                # ====================================================
                # DETERMINE LAST EMAIL DATE
                #
                # Follow-up #1:
                #     Check email_sent_date
                #
                # Follow-up #2+:
                #     Check last_followup_date
                # ====================================================

                if followup_count == 0:

                    previous_sent = person.get(
                        "email_sent_date"
                    )

                    print(
                        f"Follow-up #1 -> checking initial email sent date: "
                        f"{previous_sent}"
                    )

                else:

                    previous_sent = person.get(
                        "last_followup_date"
                    )

                    print(
                        f"Follow-up #{followup_count + 1} -> checking last follow-up date: "
                        f"{previous_sent}"
                    )

                if not previous_sent:

                    print(
                        f"⚠ No previous sent date for {receiver}"
                    )

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_followup": False
                            }
                        }
                    )

                    continue

                if not previous_sent:

                    print(
                        f"⚠ No previous sent date for {receiver}"
                    )

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_followup": False
                            }
                        }
                    )

                    continue

                if not previous_sent:

                    print(
                        f"⚠ No previous sent date for {receiver}"
                    )

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_followup": False
                            }
                        }
                    )

                    continue

                # ====================================================
                # PARSE PREVIOUS SENT DATE
                # ====================================================

                try:

                    if isinstance(
                        previous_sent,
                        datetime
                    ):

                        if previous_sent.tzinfo is None:

                            previous_sent = (
                                previous_sent.replace(
                                    tzinfo=IST
                                )
                            )

                        else:

                            previous_sent = (
                                previous_sent.astimezone(
                                    IST
                                )
                            )

                    else:

                        previous_text = str(previous_sent)

                        try:
                            previous_sent = datetime.fromisoformat(
                                previous_text.replace("Z", "+00:00")
                            )
                            if previous_sent.tzinfo is None:
                                previous_sent = previous_sent.replace(tzinfo=IST)
                            else:
                                previous_sent = previous_sent.astimezone(IST)
                        except ValueError:
                            previous_sent = datetime.strptime(
                                previous_text,
                                "%Y-%m-%d %H:%M:%S"
                            ).replace(
                                tzinfo=IST
                            )

                except Exception as date_error:

                    print(
                        f"⚠ Invalid date for {receiver}: "
                        f"{previous_sent}"
                    )

                    print(
                        "Date error:",
                        date_error
                    )

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_followup": False
                            }
                        }
                    )

                    continue

                # ====================================================
                # 5-MINUTE CADENCE CHECK
                #
                # The stored followup_due_at is the authoritative schedule.
                # It is created/repaired using the fixed 5-minute cadence.
                # ====================================================

                now = datetime.now(IST)

                scheduled_due = _coerce_ist_datetime(
                    person.get("followup_due_at")
                )

                if scheduled_due is None:
                    scheduled_due = previous_sent + FOLLOWUP_INTERVAL
                    due_text = scheduled_due.strftime("%Y-%m-%d %H:%M:%S")
                    collection.update_one(
                        {"_id": person["_id"]},
                        {
                            "$set": {
                                "followup_due_at": due_text,
                                "followup_queued_for": due_text,
                                "followup_queue_status": "scheduled",
                                "followup_queue_number": followup_number,
                                "sending_followup": False
                            }
                        }
                    )
                    print(
                        f"🧰 Repaired Follow-up #{followup_number} schedule for "
                        f"{receiver}: {due_text}"
                    )
                    collection.update_one(
                        {"_id": person["_id"]},
                        {"$set": {"sending_followup": True}}
                    )
                    scheduled_due = scheduled_due

                print(
                    f"Last email sent: {previous_sent.strftime('%Y-%m-%d %H:%M:%S IST')}"
                )
                print(
                    f"Current time: {now.strftime('%Y-%m-%d %H:%M:%S IST')}"
                )
                print(
                    f"Follow-up #{followup_number} due at: "
                    f"{scheduled_due.strftime('%Y-%m-%d %H:%M:%S IST')}"
                )

                if now < scheduled_due:
                    remaining = scheduled_due - now
                    print(
                        f"⏳ Follow-up #{followup_number} not ready for {receiver}."
                    )
                    print(f"Remaining: {remaining}")
                    collection.update_one(
                        {"_id": person["_id"]},
                        {"$set": {"sending_followup": False}}
                    )
                    continue

                # ====================================================
                # FOLLOW-UPS USE THE 5-MINUTE TEST SCHEDULE.
                # Do not apply the production weekday/daytime window.
                # ====================================================

                # ====================================================
                # REQUIRED CADENCE PASSED + EMAIL WINDOW OPEN
                # ====================================================

                print(
                    f"🚀 Follow-up #{followup_number} "
                    f"is ready for {receiver}"
                )

                # ====================================================
                # GET EMAIL DATA
                # ====================================================

                first_name = person.get(
                    "first_name",
                    ""
                )

                company = person.get(
                    "company",
                    ""
                )

                website = person.get(
                    "website",
                    ""
                )

                title = person.get(
                    "title",
                    ""
                )

                subject = person.get(
                    "last_email_subject",
                    ""
                )

                signature_html = person.get(
                    "email_signature",
                    ""
                )

                # ====================================================
                # THREAD MESSAGE ID
                # ====================================================

                last_message_id = (
                    person.get(
                        "last_message_id",
                        ""
                    )
                )

                thread_id = ""

                if last_message_id:

                    thread_id = (
                        "<"
                        + last_message_id.strip("<>")
                        + ">"
                    )

                # ====================================================
                # FOLLOW-UP TEMPLATE
                # ====================================================

                company_intro = get_followup_template(
                    person,
                    followup_number
                )

                if not company_intro:

                    print(
                        f"⚠ No template for "
                        f"Follow-up #{followup_number}"
                    )

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_followup": False
                            }
                        }
                    )

                    continue

                company_intro = render_followup_template(
                    company_intro,
                    first_name=first_name,
                    company=company,
                    title=title,
                    website=website
                )

                # ====================================================
                # EMAIL BODY
                # ====================================================

                body = f"""
                <html>
                <body style="font-family:Arial,sans-serif;font-size:10.5pt;line-height:1.4;">

                {company_intro}

                {signature_html}

                </body>
                </html>
                """

                # ====================================================
                # CREATE EMAIL
                # ====================================================

                message = MIMEMultipart(
                    "related"
                )

                alternative = MIMEMultipart(
                    "alternative"
                )

                message.attach(
                    alternative
                )

                message["From"] = SENDER_EMAIL

                message["To"] = receiver

                message["Subject"] = (
                    "Re: " + subject
                )

                new_message_id = make_msgid()

                clean_new_message_id = (
                    new_message_id.strip("<>")
                )

                message["Message-ID"] = (
                    new_message_id
                )

                # ====================================================
                # KEEP SAME EMAIL THREAD
                # ====================================================

                if thread_id:

                    message["In-Reply-To"] = (
                        thread_id
                    )

                    message["References"] = (
                        thread_id
                    )

                message["Date"] = formatdate(
                    localtime=True
                )

                html_part = MIMEText(
                    body,
                    "html",
                    "utf-8"
                )

                alternative.attach(
                    html_part
                )

                # ====================================================
                # INLINE IMAGES
                # ====================================================

                images = {
                    "logo": "static/logo.png",
                    "regards": "static/regards.png",
                    "awards": "static/awards.png",
                    "facebook": "static/facebook.png",
                    "instagram": "static/instagram.png",
                    "linkedin": "static/linkedin.png",
                    "youtube": "static/youtube.png",
                    "pinterest": "static/pinterest.png",
                    "x": "static/x.png"
                }

                for cid, path in images.items():

                    with open(
                        path,
                        "rb"
                    ) as f:

                        img = MIMEImage(
                            f.read()
                        )

                        img.add_header(
                            "Content-ID",
                            f"<{cid}>"
                        )

                        img.add_header(
                            "Content-Disposition",
                            "inline",
                            filename=os.path.basename(
                                path
                            )
                        )

                        message.attach(
                            img
                        )

                # ====================================================
                # GLOBAL 30-SECOND OUTBOUND GAP
                # ====================================================

                wait_for_outbound_email_gap()

                # Follow-ups are intentionally allowed at any time because
                # the follow-up schedule is fixed at 5 minutes.

                # ====================================================
                # SEND EMAIL
                # ====================================================

                print(
                    f"📤 Sending Follow-up #{followup_number} "
                    f"to {receiver}"
                )

                server.sendmail(
                    SENDER_EMAIL,
                    receiver,
                    message.as_string()
                )

                # ====================================================
                # ONLY UPDATE DB AFTER SUCCESSFUL SEND
                # ====================================================

                followup_count += 1

                sent_time = (
                    datetime.now(IST)
                    .strftime("%Y-%m-%d %H:%M:%S")
                )

                sent_datetime = datetime.now(IST)

                # Schedule the next follow-up using the configured cadence.
                next_followup_number = followup_count + 1
                if next_followup_number <= 5:
                    next_followup_due_at = (
                        get_next_followup_send_time(
                            sent_datetime,
                            next_followup_number
                        )
                        .strftime("%Y-%m-%d %H:%M:%S")
                    )
                else:
                    next_followup_due_at = ""

                latest_after_send = collection.find_one(
                    {"_id": person["_id"]},
                    {"has_reply": 1}
                ) or {}
                replied_after_send = bool(latest_after_send.get("has_reply"))

                update_data = {

                    "followup_sent": True,

                    "followup_sent_date":
                        sent_time,

                    "last_outbound_sent_date":
                        sent_time,

                    "last_followup_date":
                        sent_time,

                    "followup_count":
                        followup_count,

                    "last_message_id":
                        clean_new_message_id,

                    "last_email_subject":
                        "Re: " + subject,

                    "last_email_body":
                        body,

                    "campaign_status":
                        "Replied" if replied_after_send else "Running",

                    "followup_queue_status":
                        "cancelled" if replied_after_send else "scheduled",

                    "followup_queue_number":
                        0 if replied_after_send else followup_count + 1,

                    "followup_queued_for":
                        "" if replied_after_send else next_followup_due_at,

                    "followup_due_at":
                        "" if replied_after_send else next_followup_due_at,

                    "sending_followup":
                        False,

                    "email_error":
                        "",

                    "active":
                        False if replied_after_send else True
                }

                # ====================================================
                # 5 FOLLOW-UPS COMPLETED
                # ====================================================

                if followup_count >= 5 and not replied_after_send:

                    update_data[
                        "campaign_status"
                    ] = "Completed - Follow-ups Sent"

                    update_data[
                        "followup_queue_status"
                    ] = "completed"

                    update_data[
                        "followup_queue_number"
                    ] = 5

                    update_data[
                        "followup_queued_for"
                    ] = ""

                    update_data[
                        "followup_due_at"
                    ] = ""

                    update_data[
                        "active"
                    ] = False

                # ====================================================
                # DATABASE UPDATE
                # ====================================================

                collection.update_one(
                    {
                        "_id": person["_id"]
                    },
                    {
                        "$set": update_data,

                        "$push": {
                            "email_thread": {

                                "type":
                                    "followup",

                                "number":
                                    followup_count,

                                "subject":
                                    "Re: " + subject,

                                "body":
                                    body,

                                "message_id":
                                    clean_new_message_id,

                                "date":
                                    sent_time
                            }
                        }
                    }
                )

                print(
                    f"✓ Follow-up #{followup_count} "
                    f"sent successfully to {receiver}"
                )

                # Keep the same 30-second gap between every outbound email.
                if followup_index < len(prospects) - 1:
                    print(
                        f"⏳ Waiting {EMAIL_INTERVAL_SECONDS} seconds before next follow-up"
                    )
                    time.sleep(EMAIL_INTERVAL_SECONDS)
                    print(
                            "▶ 30-second gap completed."
                        )   
            # ========================================================
            # IMPORTANT:
            # ERROR FOR ONE PROFILE SHOULD NOT STOP OTHER PROFILES
            # ========================================================

            except Exception as e:

                print("=" * 70)
                print(
                    f"✗ Follow-up failed for {receiver}"
                )

                print(
                    "Error:",
                    str(e)
                )

                traceback.print_exc()

                try:

                    collection.update_one(
                        {
                            "_id": person["_id"]
                        },
                        {
                            "$set": {
                                "sending_followup":
                                    False,

                                "email_error":
                                    str(e)
                            }
                        }
                    )

                except Exception as db_error:

                    print(
                        "Failed to release follow-up lock:",
                        db_error
                    )

                # IMPORTANT:
                # Continue to next profile.
                # One bad profile must NOT terminate the whole batch.

                continue

    # ================================================================
    # GLOBAL ERROR
    # ================================================================

    except Exception as e:

        print("=" * 70)
        print(
            "Follow-up scheduler error:",
            str(e)
        )

        traceback.print_exc()

    # ================================================================
    # CLOSE SMTP
    # ================================================================

    finally:

        if server:

            try:
                server.quit()
            except Exception:
                pass

        print(
            "send_followups() finished:",
            datetime.now(IST)
        )


@app.get("/queued-emails", response_class=HTMLResponse)
def queued_emails_page(request: Request):
    """
    Show both:
      1) initial emails waiting for the next sending window
      2) customer-reply follow-ups waiting for their scheduled time
    """
    base_projection = {
        "_id": 0,
        "first_name": 1,
        "last_name": 1,
        "title": 1,
        "email": 1,
        "company": 1,
        "email_queue_subject": 1,
        "email_queued_for": 1,
        "email_queued_at": 1,
        "email_queue_reason": 1,
        "email_delivery_status": 1,
        "followup_count": 1,
        "followup_queue_status": 1,
        "followup_queue_number": 1,
        "followup_queued_for": 1,
        "followup_due_at": 1,
        "reply_received_date": 1,
        "last_email_subject": 1,
        "email_sent": 1
    }

    query = {
        "$or": [
            {
                "email_queue_status": "queued",
                "email_sent": {"$ne": True}
            },
            {
                "followup_queue_status": "scheduled",
                "email_sent": True,
                "followup_count": {"$lt": 5}
            }
        ],
        "email": {
            "$exists": True,
            "$nin": ["", None]
        }
    }

    rows = list(collection.find(query, base_projection))

    data = []

    for row in rows:
        if row.get("followup_queue_status") == "scheduled":
            followup_number = row.get("followup_queue_number") or (
                row.get("followup_count", 0) + 1
            )
            due_for = (
                row.get("followup_queued_for")
                or row.get("followup_due_at")
                or ""
            )

            row["queue_type"] = f"Follow-up #{followup_number}"
            row["email_queue_subject"] = (
                row.get("last_email_subject")
                or row.get("email_queue_subject")
                or ""
            )
            row["email_queued_for"] = due_for
            row["email_queued_at"] = (
                row.get("reply_received_date")
                or row.get("last_followup_date")
                or row.get("email_sent_date")
                or ""
            )
            row["email_queue_reason"] = "Customer replied - follow-up scheduled"
            row["email_delivery_status"] = "scheduled"

        else:
            row["queue_type"] = "Initial Email"

        data.append(row)

    data.sort(key=lambda x: x.get("email_queued_for") or "")

    return templates.TemplateResponse(
        request=request,
        name="queued_emails.html",
        context={
            "data": data,
            "queued_count": len(data)
        }
    )


@app.get(
    "/api/queued-emails",
    tags=["Queue"],
    summary="List queued emails",
    description="Return initial queued emails and reply-triggered follow-ups currently scheduled."
)
def queued_emails_api():
    """JSON API for initial queued emails + reply-triggered follow-ups."""
    query = {
        "$or": [
            {
                "email_queue_status": "queued",
                "email_sent": {"$ne": True}
            },
            {
                "followup_queue_status": "scheduled",
                "email_sent": True,
                "followup_count": {"$lt": 5}
            }
        ],
        "email": {
            "$exists": True,
            "$nin": ["", None]
        }
    }

    projection = {
        "_id": 0,
        "first_name": 1,
        "last_name": 1,
        "title": 1,
        "email": 1,
        "company": 1,
        "email_queue_subject": 1,
        "email_queued_for": 1,
        "email_queued_at": 1,
        "email_queue_reason": 1,
        "email_delivery_status": 1,
        "followup_count": 1,
        "followup_queue_status": 1,
        "followup_queue_number": 1,
        "followup_queued_for": 1,
        "followup_due_at": 1,
        "reply_received_date": 1,
        "last_email_subject": 1,
        "email_sent": 1
    }

    rows = list(collection.find(query, projection))

    data = []

    for row in rows:
        if row.get("followup_queue_status") == "scheduled":
            followup_number = row.get("followup_queue_number") or (
                row.get("followup_count", 0) + 1
            )
            due_for = (
                row.get("followup_queued_for")
                or row.get("followup_due_at")
                or ""
            )

            row["queue_type"] = f"Follow-up #{followup_number}"
            row["email_queue_subject"] = (
                row.get("last_email_subject")
                or row.get("email_queue_subject")
                or ""
            )
            row["email_queued_for"] = due_for
            row["email_queued_at"] = (
                row.get("reply_received_date")
                or row.get("email_sent_date")
                or ""
            )
            row["email_queue_reason"] = "Customer replied - follow-up scheduled"
            row["email_delivery_status"] = "scheduled"

        else:
            row["queue_type"] = "Initial Email"

        data.append(row)

    data.sort(key=lambda x: x.get("email_queued_for") or "")

    return {
        "success": True,
        "count": len(data),
        "data": data
    }


@app.get("/sent-emails", response_class=HTMLResponse)
def sent_emails_page(
    request: Request,
    from_date: str = None,
    to_date: str = None
):
    query = {
        "$and": [
            {
                "$or": [
                    {"email_sent": True},
                    {"email_delivery_status": {"$in": ["failed", "bounced", "quarantined", "smtp_failed", "undeliverable"]}}
                ]
            },
            {"email": {"$exists": True, "$nin": ["", None]}},
            {"domain": {"$exists": True, "$nin": ["", None]}}
        ]
    }

    # Date filtering
    if from_date or to_date:

        date_conditions = []

        if from_date:
            date_conditions.append({
                "$gte": [
                    {"$substr": ["$email_sent_date", 0, 10]},
                    from_date
                ]
            })

        if to_date:
            date_conditions.append({
                "$lte": [
                    {"$substr": ["$email_sent_date", 0, 10]},
                    to_date
                ]
            })

        query["$expr"] = {
            "$and": date_conditions
        }

    # IMPORTANT:
    # Only load fields actually required by sent_emails.html
    projection = {
        "_id": 0,
        "first_name": 1,
        "last_name": 1,
        "title": 1,
        "email": 1,
        "email_status": 1,
        "linkedin": 1,
        "country": 1,
        "company": 1,
        "website": 1,
        "domain": 1,
        "email_sent_date": 1,
        "followup_count": 1,
        "has_reply": 1,
        "campaign_status": 1,
        "email_delivery_status": 1,
        "email_error": 1,
        "email_failed_date": 1,
        "email_retry_after": 1,
        "last_email_subject": 1,
        "followup_due_at": 1,
        "followup_queue_status": 1,
        "followup_queue_number": 1,
        "followup_queued_for": 1
    }

    data = list(
        collection.find(
            query,
            projection
        ).sort(
            "email_sent_date",
            -1
        )
    )

    # For replied prospects, Sent Emails should show only the replied state.
    # Do not expose a stale/pending follow-up date, queue number, or schedule.
    for row in data:
        if row.get("has_reply") is True or row.get("campaign_status") == "Replied":
            row["followup_due_at"] = ""
            row["followup_queued_for"] = ""
            row["followup_queue_status"] = ""
            row["followup_queue_number"] = ""
            row["followup_count"] = ""

    return templates.TemplateResponse(
        request=request,
        name="sent_emails.html",
        context={
            "data": data,
            "from_date": from_date,
            "to_date": to_date
        }
    )
    
scheduler = BackgroundScheduler()


@app.post(
    "/blacklist-prospects",
    tags=["Prospects"],
    summary="Delete selected prospects and add them to the blacklist",
    description=(
        "Move selected prospect records to the MongoDB Blacklist collection "
        "and remove them from the active prospects collection. Blacklisted "
        "emails will be skipped on future searches."
    )
)
async def blacklist_prospects(request: Request):
    try:
        payload = await request.json()
    except Exception:
        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error": "Invalid JSON request."
            }
        )

    emails = payload.get("emails") if isinstance(payload, dict) else None

    if not isinstance(emails, list):
        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error": "emails must be a list."
            }
        )

    # Normalize and remove duplicates.
    emails = list(dict.fromkeys(
        str(email).strip().lower()
        for email in emails
        if str(email).strip()
    ))

    if not emails:
        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error": "Please select at least one prospect."
            }
        )

    blacklisted_count = 0
    deleted_count = 0
    skipped_count = 0
    details = []

    for email in emails:
        prospect = collection.find_one({"email": email})

        if not prospect:
            skipped_count += 1
            details.append({
                "email": email,
                "status": "not_found"
            })
            continue

        # Build a blacklist record from the current prospect data.
        blacklist_record = dict(prospect)
        blacklist_record.pop("_id", None)

        blacklist_record["blacklisted_email"] = email
        blacklist_record["blacklisted_at"] = datetime.now(IST).strftime(
            "%Y-%m-%d %H:%M:%S"
        )
        blacklist_record["blacklist_reason"] = "Deleted by user from Prospect Results"
        blacklist_record["deletion_status"] = "pending"

        # Keep one blacklist record per email.
        blacklist_collection.update_one(
            {"email": email},
            {
                "$set": blacklist_record
            },
            upsert=True
        )
        blacklisted_count += 1

        delete_result = collection.delete_one({"_id": prospect["_id"]})

        if delete_result.deleted_count:
            deleted_count += 1

            # Record the deletion timestamp/status so the dedicated
            # blacklist page can show both blacklist and deletion history.
            collection_updated = blacklist_collection.update_one(
                {"email": email},
                {
                    "$set": {
                        "deletion_status": "deleted",
                        "deleted_at": datetime.now(IST).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    }
                }
            )

            details.append({
                "email": email,
                "status": "blacklisted_and_deleted"
            })
        else:
            skipped_count += 1

            blacklist_collection.update_one(
                {"email": email},
                {
                    "$set": {
                        "deletion_status": "not_deleted"
                    }
                }
            )

            details.append({
                "email": email,
                "status": "blacklisted_but_not_deleted"
            })

    return {
        "success": True,
        "blacklisted": blacklisted_count,
        "deleted": deleted_count,
        "skipped": skipped_count,
        "details": details,
        "message": (
            f"{deleted_count} prospect(s) deleted and added to the blacklist."
            if deleted_count
            else "No prospect records were deleted."
        )
    }


@app.get(
    "/blacklist",
    response_class=HTMLResponse,
    tags=["Prospects"],
    summary="View blacklisted and deleted prospects",
    description=(
        "Display prospect records stored in the MongoDB Blacklist collection, "
        "including blacklist and deletion timestamps."
    )
)
def blacklist_page(request: Request):
    blacklist_data = list(
        blacklist_collection.find(
            {},
            {"_id": 0}
        ).sort("blacklisted_at", -1)
    )

    deleted_count = sum(
        1
        for record in blacklist_data
        if record.get("deletion_status") == "deleted"
        or (
            not record.get("deletion_status")
            and record.get("blacklist_reason") == "Deleted by user from Prospect Results"
        )
    )

    return templates.TemplateResponse(
        request=request,
        name="blacklist.html",
        context={
            "data": blacklist_data,
            "total_blacklisted": len(blacklist_data),
            "total_deleted": deleted_count
        }
    )


@app.get(
    "/email-thread/{email}",
    tags=["Email Thread"],
    summary="Get email conversation thread",
    description="Return the stored initial email, sent follow-ups, next follow-up content, reply state, and campaign status for a prospect."
)
def get_email_thread(email: str):

    receiver = (email or "").strip()

    if not receiver:
        return {
            "success": False,
            "message": "Email is required"
        }

    # Match the prospect case-insensitively so the frontend can always
    # retrieve the same MongoDB record even when the table email casing differs.
    prospect = collection.find_one(
        {
            "email": {
                "$regex": f"^{re.escape(receiver)}$",
                "$options": "i"
            }
        },
        {
            "_id": 0,
            "first_name": 1,
            "company": 1,
            "title": 1,
            "website": 1,
            "email_signature": 1,
            "email_thread": 1,
            "last_email_subject": 1,
            "last_email_body": 1,
            "campaign_status": 1,
            "followup_count": 1,
            "followup_queue_number": 1,
            "followup_queue_status": 1,
            "followup_due_at": 1,
            "followup_queued_for": 1,
            "has_reply": 1,
            "reply_subject": 1,
            "reply_received_date": 1,
            "followup_templates": 1
        }
    )

    if not prospect:
        return {
            "success": False,
            "message": "Thread not found"
        }

    messages = prospect.get("email_thread") or []

    # Older records may not have email_thread populated. If at least one
    # outbound email was already sent, expose the latest stored body so the
    # portal can still show useful history instead of a blank panel.
    if not messages and prospect.get("followup_count", 0) > 0:
        latest_subject = prospect.get("last_email_subject", "")
        latest_body = prospect.get("last_email_body", "")
        messages = [{
            "type": "followup",
            "number": prospect.get("followup_count", 0),
            "subject": latest_subject,
            "body": latest_body,
            "date": prospect.get("last_followup_date", "")
        }]

    next_number = int(prospect.get("followup_queue_number") or 0)
    if next_number < 1 and prospect.get("campaign_status") == "Running":
        next_number = int(prospect.get("followup_count") or 0) + 1

    # Build the complete 6-step sequence for the Send Email modal:
    # 1 initial/send email + Follow-up #1 ... #5.
    # These are previews only; the scheduler remains responsible for actual follow-up sending.
    email_sequence = []

    initial_message = next((
        message for message in messages
        if isinstance(message, dict) and message.get("type") == "initial"
    ), None)

    initial_subject = (
        (initial_message or {}).get("subject")
        or prospect.get("last_email_subject")
        or ""
    )
    initial_body = (
        (initial_message or {}).get("body")
        or prospect.get("last_email_body")
        or ""
    )

    email_sequence.append({
        "number": 0,
        "label": "Send Email",
        "type": "initial",
        "subject": initial_subject,
        "body": initial_body,
        "sent": any(
            isinstance(message, dict) and message.get("type") == "initial"
            for message in messages
        )
    })

    for sequence_number in range(1, 6):
        template = get_followup_template(prospect, sequence_number)
        if not template:
            continue

        company_intro = render_followup_template(
            template,
            first_name=prospect.get("first_name", ""),
            company=prospect.get("company", ""),
            title=prospect.get("title", ""),
            website=prospect.get("website", "")
        )

        sequence_body = f"""
        <html>
        <body style="font-family:Arial,sans-serif;font-size:10.5pt;line-height:1.4;">
        {company_intro}
        {prospect.get("email_signature", "")}
        </body>
        </html>
        """

        sent_message = next((
            message for message in messages
            if isinstance(message, dict)
            and message.get("type") == "followup"
            and int(message.get("number") or 0) == sequence_number
        ), None)

        email_sequence.append({
            "number": sequence_number,
            "label": f"Follow-up {sequence_number}",
            "type": "followup",
            "subject": (
                (sent_message or {}).get("subject")
                or "Re: " + initial_subject
            ),
            "body": (
                (sent_message or {}).get("body")
                or sequence_body
            ),
            "sent": bool(sent_message),
            "sent_date": (sent_message or {}).get("date", "")
        })

    next_followup = None

    if (
        1 <= next_number <= 5
        and prospect.get("campaign_status") == "Running"
        and not prospect.get("has_reply", False)
    ):
        template = get_followup_template(prospect, next_number)

        if template:
            company_intro = render_followup_template(
                template,
                first_name=prospect.get("first_name", ""),
                company=prospect.get("company", ""),
                title=prospect.get("title", ""),
                website=prospect.get("website", "")
            )

            next_body = f"""
            <html>
            <body style="font-family:Arial,sans-serif;font-size:10.5pt;line-height:1.4;">
            {company_intro}
            {prospect.get("email_signature", "")}
            </body>
            </html>
            """

            next_followup = {
                "number": next_number,
                "subject": "Re: " + (prospect.get("last_email_subject", "") or ""),
                "body": next_body,
                "due_at": prospect.get("followup_due_at") or prospect.get("followup_queued_for") or ""
            }

    return {
        "success": True,
        "campaign_status": prospect.get("campaign_status", ""),
        "followup_count": prospect.get("followup_count", 0),
        "followup_queue_number": prospect.get("followup_queue_number", 0),
        "followup_queue_status": prospect.get("followup_queue_status", ""),
        "followup_due_at": prospect.get("followup_due_at", ""),
        "followup_queued_for": prospect.get("followup_queued_for", ""),
        "has_reply": prospect.get("has_reply", False),
        "reply_subject": prospect.get("reply_subject", ""),
        "reply_received_date": prospect.get("reply_received_date", ""),
        "messages": messages,
        "email_sequence": email_sequence,
        # Unrendered templates (placeholders intact) for the compose editor.
        "followup_templates": {
            str(number): get_followup_template(prospect, number) or ""
            for number in range(1, 6)
        },
        "next_followup": next_followup
    }

class FollowupTemplatesRequest(BaseModel):
    emails: list[str] = Field(
        ...,
        description="Prospect emails whose follow-up sequence should be updated.",
        min_length=1
    )
    followup_templates: dict[str, str] = Field(
        ...,
        description=(
            "Follow-up HTML templates keyed by follow-up number (\"1\"-\"5\"). "
            "Supported placeholders: {first_name}, {company}, {title}, {website}."
        )
    )


@app.get(
    "/followup-templates",
    tags=["Email Campaign"],
    summary="Get default follow-up templates"
)
def default_followup_templates():
    return {
        "success": True,
        "followup_templates": {
            str(number): template
            for number, template in FOLLOWUP_TEMPLATES.items()
        }
    }


@app.post(
    "/followup-templates",
    tags=["Email Campaign"],
    summary="Save follow-up templates for selected prospects",
    description=(
        "Store edited Follow-up #1-#5 templates on every selected prospect. "
        "The follow-up scheduler sends these instead of the defaults. "
        "Follow-ups that were already sent are not affected."
    )
)
def save_followup_templates(payload: FollowupTemplatesRequest):
    # Match both the submitted and lowercase spelling of each address.
    emails = sorted({
        variant
        for email in payload.emails
        if (email or "").strip()
        for variant in ((email or "").strip(), (email or "").strip().lower())
    })

    templates = {}
    for key, value in (payload.followup_templates or {}).items():
        try:
            number = int(key)
        except (TypeError, ValueError):
            continue
        if 1 <= number <= 5 and isinstance(value, str) and value.strip():
            templates[str(number)] = value

    if not emails or not templates:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "Emails and at least one follow-up template are required"}
        )

    update_fields = {
        f"followup_templates.{number}": html
        for number, html in templates.items()
    }
    update_fields["followup_templates_updated_at"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

    result = collection.update_many(
        {
            "email": {
                "$in": emails
            }
        },
        {
            "$set": update_fields
        }
    )

    logger.info(
        "Saved follow-up templates %s for %s prospect(s)",
        sorted(templates.keys()),
        result.matched_count
    )

    return {
        "success": True,
        "matched": result.matched_count,
        "updated": result.modified_count,
        "followups": sorted(templates.keys())
    }


def process_overdue_queue_on_startup():
    """
    When the project starts, immediately process emails
    whose scheduled time has already passed.
    """

    now = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

    queued_emails = collection.find({
        "email_queue_status": "queued",
        "email_sent": {"$ne": True},
        "email_sending": {"$ne": True},
        "email_queued_for": {
            "$exists": True,
            "$lte": now
        }
    }).sort("email_queued_for", 1)

    for person in queued_emails:
        print(
            f"📤 Processing overdue queued email: "
            f"{person.get('email')} | "
            f"scheduled: {person.get('email_queued_for')}"
        )

    # Let the normal worker send them.
# Repair any pending queue records that still use the previous daytime 09:00-17:00 schedule.
repair_existing_initial_queue_slots()


scheduler.add_job(
    check_replies,
    "interval",
    minutes=5,
    max_instances=1
)

scheduler.add_job(
    send_followups,
    "interval",
    seconds=10,
    max_instances=1
)
process_overdue_queue_on_startup()

scheduler.add_job(
    send_queued_initial_emails,
    "interval",
    seconds=10,
    max_instances=1
)

scheduler.add_job(
    clear_stale_initial_locks,
    "interval",
    minutes=10,
    max_instances=1
)


# Stop any follow-up sequence that already has a customer reply.
repair_followups_after_replies()
# Repair pending Follow-up #1 records from the initial email date.
repair_existing_first_followup_slots()
# Repair pending Follow-up #2-#5 records using the configured follow-up cadence.
repair_existing_followup_slots()
# Repair any legacy records where all 5 follow-ups were sent but the
# campaign status was not stored/displayed as completed.
repair_completed_followup_statuses()

logger.info(
    "Follow-up cadence enabled: #1-#3 every 48 hours, #4-#5 every 120 hours when no reply is found."
)

scheduler.start()