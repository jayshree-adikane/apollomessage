import os
from dotenv import load_dotenv

# Secrets and credentials come from the local .env file (see .env.example).
load_dotenv()

# 
from fastapi import FastAPI, Form, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from pymongo import MongoClient, ReturnDocument
import requests
import time
from datetime import datetime, timedelta, timezone
from datetime import datetime, timedelta, time
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
import re
from urllib.parse import urlparse
from urllib.parse import urlparse, urlunparse
from fastapi.staticfiles import StaticFiles
from email.mime.image import MIMEImage
from google import genai
import traceback
from zoneinfo import ZoneInfo
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")

#Apollo ID
API_KEY = os.getenv("APOLLO_API_KEY_2", "")

headers = {
    "X-Api-Key": API_KEY,
    "Content-Type": "application/json"
}

# # MongoDB local
client = MongoClient("mongodb://localhost:27017")
db = client["Apollo_Email"]
collection = db["prospects"]
blacklist_collection = db["Blacklist"]

# MongoDB online
# client = MongoClient(os.getenv("MONGODB_URI", ""))
# db = client["Apollo_Email"]
# collection = db["prospects"]
# blacklist_collection = db["Blacklist"]

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
SCOPES = ["User.Read", "Mail.Read"]

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

FOLLOWUP_INTERVALS = {
    1: timedelta(hours=48),
    2: timedelta(hours=48),
    3: timedelta(hours=48),
    4: timedelta(hours=48),
    5: timedelta(hours=48),
}

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

    if accounts:
        result = msal_app.acquire_token_silent_with_error(
            SCOPES,
            account=accounts[0]
        )

        print("Silent token result:", result)

    if not result:
        print("Silent authentication failed. Opening login...")

        result = msal_app.acquire_token_interactive(
            scopes=SCOPES
        )

    if "access_token" not in result:
        raise Exception(
            f"Authentication failed: {result}"
        )

    if cache.has_state_changed:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())

        print("Token cache updated:", CACHE_FILE)

    return result["access_token"]


@app.get("/", response_class=HTMLResponse)
def home(request: Request):

    query = {
        "$and": [
            {
                "$or": [
                    {"email_sent": {"$exists": False}},
                    {"email_sent": False},
                    {"email_sent": None}
                ]
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


@app.post("/", response_class=HTMLResponse)
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
                {
                    "$or": [
                        {"email_sent": {"$exists": False}},
                        {"email_sent": False},
                        {"email_sent": None}
                    ]
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
        blacklist_match = blacklist_collection.find_one({
            "full_name": record.get("full_name"),
            "title": record.get("title"),
            "email": record.get("email"),
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
            {
                "$or": [
                    {"email_sent": {"$exists": False}},
                    {"email_sent": False},
                    {"email_sent": None}
                ]
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


@app.post("/send-emails")
async def send_emails(request: Request):
    try:
        # Get JSON data from request
        data = await request.json()

        recipients = data.get('recipients', [])
        subject_template = data.get('subject', '')
        body_template = data.get('body', '')
        signature_html = data.get("signature", "")

        # Remove duplicate recipients
        unique_recipients = []
        seen_emails = set()

        for recipient in recipients:

            email = (recipient.get("email") or "").strip().lower()

            if not email:
                continue

            if email in seen_emails:
                print(f"⚠ Duplicate recipient removed: {email}")
                continue

            seen_emails.add(email)
            unique_recipients.append(recipient)

        recipients = unique_recipients
        
        if not recipients:
            return JSONResponse(
                status_code=400,
                content={'error': 'No recipients provided'}
            )
        
        # Connect to SMTP server
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls(context=ssl.create_default_context())
        server.login(SENDER_EMAIL, PASSWORD)
        
        sent_count = 0
        failed_count = 0
        results = []


        subject_patterns = [
            "How is {company} finding in-market accounts?",
            "How does {company} spot ready-to-buy accounts?"
        ]
        
        # Send emails to each recipient
        for i, recipient in enumerate(recipients):
            receiver = (recipient.get('email') or '').strip().lower()

            try:

                if not receiver:
                    print("⚠ Skipping recipient without email")
                    continue

                # ============================================================
                # ATOMIC INITIAL EMAIL LOCK
                # ============================================================

                claim_result = collection.update_one(
                    {
                        "email": receiver,

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
                            "initial_send_started_at": datetime.now(IST).strftime(
                                "%Y-%m-%d %H:%M:%S"
                            )
                        }
                    }
                )

                if claim_result.modified_count == 0:

                    print(
                        f"⚠ SKIPPED duplicate/concurrent initial send: {receiver}"
                    )

                    results.append({
                        "email": receiver,
                        "status": "already_sent_or_in_progress"
                    })

                    continue

                print(
                    f"🔒 Initial email claimed for {receiver}"
                )

                first_name = recipient.get('first_name', '')
                title = recipient.get('title', '')
                company = recipient.get('company', '')
                person = collection.find_one(
                    {"email": receiver},
                    {"website": 1}
                )

                website = ""

                if person:
                    website = person.get("website", "")

                company_about = generate_company_about(
                    company,
                    website,
                    title
                )
                company_about = company_about or ""
                
                if not receiver:
                    continue
                
                ## Rotate subject for each recipient
                subject_template = subject_patterns[i % len(subject_patterns)]
                # Replace placeholders in subject
                subject = subject_template.replace('{company}', company)
                subject = subject.replace('{first_name}', first_name)
                subject = subject.replace('{title}', title)
                
                # Replace placeholders in body
                body = body_template.replace('{company}', company)
                body = body.replace('{first_name}', first_name)
                body = body.replace('{title}', title)
                body = body.replace('{company_about}', company_about)

                # Append the same signature used in index.html
                body = body + "<br><br>" + signature_html

                # Insert common signature
                body = body.replace(
                    "{signature}",
                    get_signature()
                )
                
                # Create message
                message = MIMEMultipart("related")

                alternative = MIMEMultipart("alternative")

                html_part = MIMEText(body, "html", "utf-8")
                alternative.attach(html_part)

                message.attach(alternative)

                message["From"] = SENDER_EMAIL
                message["To"] = receiver
                message["Subject"] = subject

                message_id = make_msgid()

                clean_message_id = message_id.strip("<>")

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

                for cid, path in images.items():
                    with open(path, "rb") as f:
                        img = MIMEImage(f.read())
                        img.add_header("Content-ID", f"<{cid}>")
                        img.add_header("Content-Disposition", "inline", filename=path)
                        message.attach(img)
                                
                # Send email
                # Prevent duplicate initial email
                lock = collection.update_one(
                    {
                        "_id": person["_id"],
                        "email_sent": {"$ne": True},
                        "email_sending": {"$ne": True}
                    },
                    {
                        "$set": {"email_sending": True}
                    }
                )

                if lock.modified_count == 0:
                    print(f"Already sent/being sent: {receiver}")
                    continue

                # Send email
                server.sendmail(
                    SENDER_EMAIL,
                    receiver,
                    message.as_string()
                )

                collection.update_one(
                    {"_id": person["_id"]},
                    {
                        "$set": {
                            "email_sent": True,
                            "email_sending": False,
                            "email_sent_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                        }
                    }
                )
                
                current_time = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                
                collection.update_one(
                    {
                        "email": receiver,
                        "sending_initial": True
                    },
                    {
                        "$set": {

                            "email_sent": True,
                            "email_sent_date": current_time,
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

                            "campaign_status": "Running",
                            "has_reply": False,

                            "reply_subject": "",
                            "reply_received_date": "",

                            # Release initial-email lock
                            "sending_initial": False,

                            # Clear previous error if any
                            "email_error": "",

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
                    
                sent_count += 1
                results.append({"email": receiver, "status": "sent"})
                print(f"✓ Sent to {receiver} and updated database")
                
            except Exception as e:
                failed_count += 1
                
                # ❌ UPDATE DATABASE - Mark as failed
                current_time = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                
                collection.update_one(
                    {
                        "email": receiver,
                        "sending_initial": True
                    },
                    {
                        "$set": {
                            "email_sent": False,
                            "email_delivery_status": "failed",
                            "email_error": str(e),
                            "email_failed_date": current_time,

                            # Release lock so it can be retried
                            "sending_initial": False
                        }
                    }
                )
                
                results.append({"email": receiver, "status": "failed", "error": str(e)})
                print(f"✗ Failed: {receiver} - {str(e)}")
        
        server.quit()
        
        return JSONResponse(
            content={
                'success': True,
                'sent': sent_count,
                'failed': failed_count,
                'results': results,
                'message': f'Successfully sent {sent_count} emails. Failed: {failed_count}'
            }
        )
        
    except Exception as e:
        print(f"Error in send_emails route: {str(e)}")
        return JSONResponse(
            status_code=500,
            content={'error': str(e)}
        )
    
def check_replies():

    try:
        token = get_access_token()

        headers = {
            "Authorization": f"Bearer {token}"
        }

        since = (
            datetime.now(timezone.utc)
            - timedelta(hours=24)
        ).strftime("%Y-%m-%dT%H:%M:%SZ")

        url = (
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
            f"?$filter=receivedDateTime ge {since}"
            "&$select=id,subject,receivedDateTime,internetMessageHeaders,from"
            "&$top=100"
        )

        response = requests.get(
            url,
            headers=headers
        )

        if response.status_code == 200:

            messages = response.json().get("value", [])

            for msg in messages:
                message_id = msg["id"]

                header_response = requests.get(

                    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}?$select=internetMessageHeaders",

                    headers=headers

                )

                internet_headers = []

                if header_response.status_code == 200:

                    internet_headers = header_response.json().get(

                        "internetMessageHeaders",

                        []

                    )

                sender = (
                    msg.get("from", {})
                    .get("emailAddress", {})
                    .get("address", "")
                    .lower()
                )

                if sender == SENDER_EMAIL.lower():
                    continue

                print("=" * 50)
                print("Subject:", msg.get("subject"))
                print("Sender :", sender)
                print("Received:", msg.get("receivedDateTime"))

                email_headers = internet_headers

                in_reply_to = ""
                references = ""

                for h in email_headers:

                    name = h.get("name", "").lower()

                    if name == "in-reply-to":
                        in_reply_to = h.get("value", "").strip("<>")

                    elif name == "references":
                        references = h.get("value", "")

                existing = None

                # -------------------------------------------------
                # First try matching by Message-ID
                # -------------------------------------------------

                if sender != SENDER_EMAIL.lower() and in_reply_to:
                    existing = collection.find_one({
                        "thread_message_id": in_reply_to,
                        "campaign_status": "Running"
                    })

                # -------------------------------------------------
                # If not found, search inside References header
                # -------------------------------------------------

                if not existing and references:

                    prospects = collection.find({

                        "campaign_status": "Running"

                    })

                    for p in prospects:

                        thread_id = p.get("thread_message_id", "")

                        if thread_id and thread_id in references:

                            existing = p

                            break

                # -------------------------------------------------
                # Final fallback = sender email
                # -------------------------------------------------

                if not existing:

                    existing = collection.find_one({

                        "email": sender,

                        "campaign_status": "Running"

                    })

                if existing:

                    collection.update_one(

                        {"_id": existing["_id"]},

                        {

                            "$set": {

                                "has_reply": True,

                                "campaign_status": "Replied",

                                "reply_subject": msg.get("subject"),

                                "reply_received_date": msg.get("receivedDateTime")

                            }

                        }

                    )
                    print("Reply received from", sender)

    except Exception as e:
        print(e)


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

        cutoff_time = now - timedelta(hours=48)

        cutoff_string = cutoff_time.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        print(
            f"Looking for profiles due before or at: "
            f"{cutoff_string} IST"
        )

        # ============================================================
        # FOLLOW-UP 1
        # Initial email was sent 48+ hours ago
        # ============================================================

        first_followup_query = {
            "email_sent": True,
            "email_delivery_status": "sent",
            "has_reply": False,
            "campaign_status": "Running",
            "sending_followup": {
                "$ne": True
            },
            "followup_count": 0,
            "email_sent_date": {
                "$lte": cutoff_string
            }
        }

        # ============================================================
        # FOLLOW-UP 2, 3, 4, 5
        # Previous follow-up was sent 48+ hours ago
        # ============================================================

        later_followup_query = {
            "email_sent": True,
            "email_delivery_status": "sent",
            "has_reply": False,
            "campaign_status": "Running",
            "sending_followup": {
                "$ne": True
            },
            "followup_count": {
                "$gte": 1,
                "$lt": 5
            },
            "last_followup_date": {
                "$lte": cutoff_string
            }
        }

        # ============================================================
        # GET ONLY DUE PROFILES
        # ============================================================

        first_followups = list(
            collection.find(first_followup_query)
            .sort("email_sent_date", 1)
            .limit(MAX_PROFILES_PER_RUN)
        )

        remaining_slots = (
            MAX_PROFILES_PER_RUN - len(first_followups)
        )

        later_followups = []

        if remaining_slots > 0:

            later_followups = list(
                collection.find(later_followup_query)
                .sort("last_followup_date", 1)
                .limit(remaining_slots)
            )

        # Combine both types
        prospects = first_followups + later_followups

        print(
            f"Due profiles selected: {len(prospects)}"
        )

        print(
            f"First follow-ups due: {len(first_followups)}"
        )

        print(
            f"Later follow-ups due: {len(later_followups)}"
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

        for person in prospects:

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
                                    "Completed - No Reply",

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
                # LAST ACTUALLY SENT EMAIL.
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
                        f"Follow-up #1 -> checking initial email date: "
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

                        previous_sent = datetime.strptime(
                            str(previous_sent),
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
                # 48-HOUR CHECK
                #
                # NO WEEKEND LOGIC
                # NO 11 AM LOGIC
                # NO 6 PM LOGIC
                #
                # Simply compare current time against last sent time.
                # ====================================================

                now = datetime.now(IST)

                elapsed_time = (
                    now - previous_sent
                )

                required_interval = timedelta(
                    hours=48
                )

                print(
                    f"Last email sent: "
                    f"{previous_sent.strftime('%Y-%m-%d %H:%M:%S IST')}"
                )

                print(
                    f"Current time: "
                    f"{now.strftime('%Y-%m-%d %H:%M:%S IST')}"
                )

                print(
                    f"Elapsed: "
                    f"{elapsed_time}"
                )

                # ====================================================
                # NOT YET 48 HOURS
                # ====================================================

                if elapsed_time < required_interval:

                    remaining = (
                        required_interval
                        - elapsed_time
                    )

                    print(
                        f"⏳ Follow-up #{followup_number} "
                        f"not ready for {receiver}."
                    )

                    print(
                        f"Remaining: {remaining}"
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
                # 48 HOURS PASSED
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

                company_intro = FOLLOWUP_TEMPLATES.get(
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

                company_intro = company_intro.format(
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

                update_data = {

                    "followup_sent": True,

                    "followup_sent_date":
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
                        "Running",

                    "sending_followup":
                        False,

                    "email_error":
                        ""
                }

                # ====================================================
                # 5 FOLLOW-UPS COMPLETED
                # ====================================================

                if followup_count >= 5:

                    update_data[
                        "campaign_status"
                    ] = "Completed - No Reply"

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


@app.get("/sent-emails", response_class=HTMLResponse)
def sent_emails_page(
    request: Request,
    from_date: str = None,
    to_date: str = None
):
    query = {
        "email_sent": True,
        "email": {"$exists": True, "$nin": ["", None]},
        "domain": {"$exists": True, "$nin": ["", None]}
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
        "last_email_subject": 1
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

@app.get("/email-thread/{email}")
def get_email_thread(email:str):

    prospect = collection.find_one(
        {
            "email":email
        },
        {
            "_id":0,
            "email_thread":1,
            "campaign_status":1,
            "followup_count":1,
            "has_reply":1,
            "reply_subject":1,
            "reply_received_date":1
        }
    )


    if not prospect:

        return {
            "success":False,
            "message":"Thread not found"
        }


    return {

        "success":True,

        "campaign_status":
            prospect.get(
                "campaign_status",
                ""
            ),

        "followup_count":
            prospect.get(
                "followup_count",
                0
            ),

        "has_reply":
            prospect.get(
                "has_reply",
                False
            ),

        "reply_subject":
            prospect.get(
                "reply_subject",
                ""
            ),

        "reply_received_date":
            prospect.get(
                "reply_received_date",
                ""
            ),

        "messages":
            prospect.get(
                "email_thread",
                []
            )

    }


scheduler.add_job(
    check_replies,
    "interval",
    minutes=5
)

scheduler.add_job(
    send_followups,
    "interval",
    minutes=1,
    max_instances=1
)

scheduler.add_job(
    clear_stale_initial_locks,
    "interval",
    minutes=10,
    max_instances=1
)
scheduler.start()