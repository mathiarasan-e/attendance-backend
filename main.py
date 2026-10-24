from __future__ import annotations

import hashlib
import logging
import math
import os
import re
import secrets
import smtplib
import ssl
import time
from contextlib import asynccontextmanager, contextmanager
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path
from typing import Optional

import psycopg
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from pydantic import BaseModel, Field

try:  # psycopg_pool is optional; the app falls back to direct connections.
    from psycopg_pool import ConnectionPool
except Exception:  # pragma: no cover
    ConnectionPool = None


logging.basicConfig(level=logging.INFO)
log = logging.getLogger("attendance")


def _read_env_file(path: str) -> dict:
    """Parse simple KEY=VALUE pairs from a local .env file.

    Written by hand to avoid adding a dependency to the AppSail bundle for a
    few lines of parsing. Returns an empty mapping when the file is absent.
    """
    values: dict = {}
    try:
        with open(path, encoding="utf-8") as handle:
            lines = handle.readlines()
    except OSError:
        return values

    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        value = value.strip()
        # Strip the surrounding quotes that `neon link` writes.
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key.strip()] = value
    return values


def _load_env_files(*paths: str) -> None:
    """Apply local .env files without ever overriding the real environment.

    Later files win over earlier ones, so .env.local overrides .env. Anything
    already set in the process environment beats all of them, which keeps a
    deployed container authoritative and lets a one-off command such as
    `$env:DATABASE_URL=...; python clear_staff_passwords.py` win locally.
    """
    real_env = set(os.environ)
    merged: dict = {}
    for path in paths:
        merged.update(_read_env_file(path))

    for key, value in merged.items():
        if key not in real_env:
            os.environ[key] = value


# .env carries the configuration the deployed service runs with. .env.local
# overrides it for local development (typically pointing DATABASE_URL at a Neon
# dev branch) and is never committed.
_load_env_files(".env", ".env.local")


# ============================================================
# CONFIGURATION (all values come from environment variables)
# ============================================================

DATABASE_URL = os.getenv("DATABASE_URL")

CODE_VALIDITY_MINUTES = int(os.getenv("CODE_VALIDITY_MINUTES", "15"))

# How far a student may be from the classroom and still mark attendance.
# Widened from 60 m because browser geolocation is less precise than native.
MAX_ATTENDANCE_DISTANCE_METERS = float(
    os.getenv("MAX_ATTENDANCE_DISTANCE_METERS", "120")
)

LOCATION_REQUIRED = os.getenv("LOCATION_REQUIRED", "true").lower() in {
    "1",
    "true",
    "yes",
}

CLASSROOM_LATITUDE = float(os.getenv("CLASSROOM_LATITUDE", "12.9869153"))
CLASSROOM_LONGITUDE = float(os.getenv("CLASSROOM_LONGITUDE", "79.9724422"))

# Semester window (used only for the export header / reporting defaults).
SEMESTER_START = os.getenv("SEMESTER_START")
SEMESTER_END = os.getenv("SEMESTER_END")

# ERP delivery. mode = "file" (download + email) or "api" (push to ERP endpoint).
ERP_MODE = os.getenv("ERP_MODE", "file").lower()
ERP_UPLOAD_EMAIL = os.getenv("ERP_UPLOAD_EMAIL")
ERP_API_URL = os.getenv("ERP_API_URL")
ERP_API_KEY = os.getenv("ERP_API_KEY")

# Outbound mail (Catalyst Mail / any SMTP provider).
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_FROM = os.getenv("SMTP_FROM") or SMTP_USER

# The destructive reset endpoint is disabled unless explicitly enabled.
RESET_DB_ENABLED = os.getenv("RESET_DB_ENABLED", "").lower() in {
    "1",
    "true",
    "yes",
}

# Staff choose their own password the first time they sign in, and must type this
# code to do it. The college hands the code out to staff; changing it here simply
# retires the old one. While it is empty, self-service setup is switched off and
# only an administrator can clear an account.
STAFF_SETUP_CODE = os.getenv("STAFF_SETUP_CODE", "").strip()

ALLOWED_ORIGINS = [
    origin.strip()
    for origin in os.getenv("ALLOWED_ORIGINS", "*").split(",")
    if origin.strip()
]

# Compiled Flutter web app. build_appsail.ps1 copies it into build/web, so it
# exists in the deployed bundle and is absent during local development.
# Serving the UI from the same origin as the API avoids CORS altogether.
# WEB_DIR overrides the location, which is handy for local testing.
_web_dir_env = os.getenv("WEB_DIR", "").strip()
WEB_DIR = (
    Path(_web_dir_env)
    if _web_dir_env
    else Path(__file__).resolve().parent / "web"
)

PBKDF2_ITERATIONS = int(os.getenv("PBKDF2_ITERATIONS", "200000"))

MIN_PASSWORD_LENGTH = 4

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ============================================================
# STUDENTS
# ============================================================

STUDENTS: dict[str, str] = {
    "123": "JOEL S",
    "124": "K DEVADHARSHINI",
    "125": "K DHANANJAY",
    "126": "K M LEKHAA VARSHINI",
    "127": "K R HARSHINI",
    "128": "K SHAMITHA",
    "129": "KAILASH SURYA P",
    "130": "KAMALA KANNAN S",
    "131": "KAMALI S B",
    "132": "KANISHKA A",
    "133": "KAPILVENKAT K",
    "134": "KARTHICK S",
    "135": "KARTHIGA D",
    "136": "KARTHIGA S",
    "137": "KARUNYA G",
    "138": "KAVINAYA J",
    "139": "KAVIYA G",
    "140": "KAVIYA J",
    "141": "KEERTHANA R",
    "142": "KEERTHANA S",
    "143": "KESAVARAM P",
    "144": "KIRUPANITHI S",
    "145": "KISHORE K",
    "146": "KRISHITHA J",
    "147": "KRISHNA A M",
    "148": "KRITHICK S",
    "149": "LAKSHAYA G",
    "150": "LAKSHAYA NAGARAJAN",
    "151": "LAKSHIKHAA SRI B",
    "152": "LAKSHMAN M",
    "153": "LAKSHMI PRIYA S",
    "154": "LALITH CHANDAR M",
    "155": "LAURA SUSAN C",
    "156": "LAVANYA R",
    "157": "LIVISHARAN E",
    "158": "LOKESH D",
    "159": "LOKESHWARI V",
    "160": "M GUNADHARSHINI",
    "161": "M S ABISHIEK",
    "162": "M YASHWANTH SAI",
    "163": "MADHAN RAJ D",
    "164": "MALIK KANI M K",
    "165": "MANASA S",
    "166": "MANSI R",
    "167": "MATHANGI P",
    "168": "MATHAVAN R D",
    "169": "MATHAVAN S",
    "170": "MATHIARASAN E",
    "171": "MEENAKSHI A",
    "172": "MITHRAA A",
    "173": "MOHAMED ANEESUR RAHMAN A",
    "174": "MOHAMED JASIM M R",
    "175": "MOHAMMED NADEEM K S",
    "176": "MOHANKUMAR M",
    "177": "MOKSHITA K G",
    "178": "MONISH V",
    "179": "MONISHA V",
    "180": "MOONISHHA T",
    "181": "MOURIYA M",
    "182": "MUGUNDHASARAVANAN D",
    "183": "MUHAMMED JIYAD A",
    "184": "MUKUL V P",
    "185": "NAREN V",
}

# Maps the last four digits of a student's admission number
# (2026CSxxxx) to their s.no. This is what students type to log in,
# while the s.no is the three-digit number shown in the app.
ADMISSION_LAST4_TO_SNO: dict[str, str] = {
    "1302": "123",
    "0398": "124",
    "0683": "125",
    "0707": "126",
    "1219": "127",
    "0517": "128",
    "0448": "129",
    "0184": "130",
    "0475": "131",
    "0215": "132",
    "0776": "133",
    "0504": "134",
    "0500": "135",
    "0167": "136",
    "1026": "137",
    "0777": "138",
    "0921": "139",
    "0991": "140",
    "0005": "141",
    "1466": "142",
    "1470": "143",
    "1354": "144",
    "0756": "145",
    "0491": "146",
    "0209": "147",
    "0718": "148",
    "1340": "149",
    "0652": "150",
    "1452": "151",
    "1402": "152",
    "1353": "153",
    "0898": "154",
    "0380": "155",
    "1050": "156",
    "0359": "157",
    "1261": "158",
    "0164": "159",
    "0358": "160",
    "0351": "161",
    "0365": "162",
    "0218": "163",
    "0395": "164",
    "0345": "165",
    "0724": "166",
    "0045": "167",
    "0378": "168",
    "0582": "169",
    "0559": "170",
    "0021": "171",
    "0779": "172",
    "0082": "173",
    "0079": "174",
    "0828": "175",
    "0927": "176",
    "1443": "177",
    "0354": "178",
    "0174": "179",
    "1372": "180",
    "0469": "181",
    "0393": "182",
    "1436": "183",
    "0125": "184",
    "0883": "185",
}


# ============================================================
# SUBJECTS + STAFF
# ============================================================

SUBJECTS: dict[str, str] = {
    "Tamil": "Ms. S. Saranya",
    "Communicative English": "Ms. S. Siva Prathesha",
    "Algebra and Calculus": "Dr. G. SatheeshKumar",
    "Applied Chemistry": "Dr. A.R. Karthiga",
    "Computer Programming 1": "Ms. N. Suriya",
    "CAD Based Engineering graphics": "Dr. A. Saravanan",
    "Computer Programming laboratory I": "Ms. N. Suriya",
    "Makerspace": "Ms. R. Priyadharshini",
    "Seminar 1": "Ms. V. Ariyamala",
    "Seminar 2": "Ms. N. Suriya",
    "Seminar 3": "Ms. V. Ariyamala",
}


# ============================================================
# APP
# ============================================================

@asynccontextmanager
async def lifespan(_app: FastAPI):
    """Startup and shutdown wiring.

    Replaces the deprecated ``on_event`` hooks. ``init_database`` and ``_pool``
    are defined further down: they only need to exist by the time the app
    actually starts serving requests, which is well after this module is
    imported.
    """
    global _pool

    if DATABASE_URL:
        init_database()
    else:
        log.error("DATABASE_URL is not set. Database calls will fail.")

    try:
        yield
    finally:
        # Release pooled connections so the app shuts down cleanly.
        if _pool is not None:
            _pool.close()
            _pool = None


app = FastAPI(
    title="College Attendance Backend",
    version="3.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# DATABASE
# ============================================================

_pool = None


def _get_pool():
    global _pool
    if ConnectionPool is None or not DATABASE_URL:
        return None
    if _pool is None:
        _pool = ConnectionPool(
            DATABASE_URL,
            min_size=1,
            max_size=5,
            kwargs={"row_factory": psycopg.rows.dict_row},
            open=True,
        )
    return _pool


@contextmanager
def get_connection():
    pool = _get_pool()
    if pool is not None:
        with pool.connection() as conn:
            yield conn
        return

    conn = psycopg.connect(DATABASE_URL, row_factory=psycopg.rows.dict_row)
    try:
        yield conn
    finally:
        conn.close()


def init_database() -> None:
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS students (
                    sno VARCHAR(10) PRIMARY KEY,
                    name TEXT NOT NULL
                )
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS student_passwords (
                    sno VARCHAR(10) PRIMARY KEY
                        REFERENCES students(sno)
                        ON DELETE CASCADE,
                    password_hash TEXT NOT NULL,
                    email TEXT,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )
            # Upgrade path for databases created before the email column existed.
            cur.execute(
                "ALTER TABLE student_passwords ADD COLUMN IF NOT EXISTS email TEXT"
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS staff_accounts (
                    staff_name TEXT PRIMARY KEY,
                    password_hash TEXT,
                    must_change BOOLEAN NOT NULL DEFAULT TRUE,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )

            # A staff account starts with no password: the staff member sets
            # their own the first time they sign in. Deployments that predate
            # that created the column NOT NULL, so drop the constraint. Running
            # this again is a no-op and no data is touched.
            cur.execute(
                "ALTER TABLE staff_accounts ALTER COLUMN password_hash DROP NOT NULL"
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance_sessions (
                    id BIGSERIAL PRIMARY KEY,
                    subject TEXT NOT NULL,
                    staff_name TEXT NOT NULL,
                    period TEXT NOT NULL,
                    started_at TIMESTAMPTZ NOT NULL,
                    expires_at TIMESTAMPTZ NOT NULL,
                    closed_at TIMESTAMPTZ,
                    active BOOLEAN NOT NULL DEFAULT TRUE
                )
                """
            )

            cur.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS
                one_active_attendance_session_per_subject
                ON attendance_sessions (subject)
                WHERE active = TRUE
                """
            )

            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS
                attendance_sessions_active_expires
                ON attendance_sessions (active, expires_at)
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance_codes (
                    id BIGSERIAL PRIMARY KEY,
                    session_id BIGINT NOT NULL
                        REFERENCES attendance_sessions(id)
                        ON DELETE CASCADE,
                    sno VARCHAR(10) NOT NULL
                        REFERENCES students(sno)
                        ON DELETE CASCADE,
                    code VARCHAR(20) NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    UNIQUE(session_id, sno),
                    UNIQUE(session_id, code)
                )
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance_records (
                    id BIGSERIAL PRIMARY KEY,
                    session_id BIGINT NOT NULL
                        REFERENCES attendance_sessions(id)
                        ON DELETE CASCADE,
                    sno VARCHAR(10) NOT NULL
                        REFERENCES students(sno)
                        ON DELETE CASCADE,
                    subject TEXT NOT NULL,
                    period TEXT NOT NULL,
                    attendance_date DATE NOT NULL,
                    marked_at TIMESTAMPTZ,
                    status VARCHAR(20) NOT NULL,
                    latitude DOUBLE PRECISION,
                    longitude DOUBLE PRECISION,
                    distance_meters DOUBLE PRECISION,
                    UNIQUE(session_id, sno)
                )
                """
            )

            cur.execute(
                """
                CREATE INDEX IF NOT EXISTS
                attendance_records_subject_sno
                ON attendance_records (subject, sno)
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance_device_locks (
                    id BIGSERIAL PRIMARY KEY,
                    session_id BIGINT NOT NULL
                        REFERENCES attendance_sessions(id)
                        ON DELETE CASCADE,
                    device_id TEXT NOT NULL,
                    sno VARCHAR(10) NOT NULL
                        REFERENCES students(sno)
                        ON DELETE CASCADE,
                    marked_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    UNIQUE(session_id, device_id)
                )
                """
            )

            # Students do not reset a password themselves - remember your
            # password is the rule, and the college clears one that is forgotten
            # (see clear_student_password.py). The table that held emailed reset
            # codes is dropped so no stale codes linger in a live database.
            cur.execute("DROP TABLE IF EXISTS verification_codes")

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS erp_sync_log (
                    id BIGSERIAL PRIMARY KEY,
                    subject TEXT NOT NULL,
                    mode TEXT NOT NULL,
                    period_start DATE,
                    period_end DATE,
                    student_count INTEGER NOT NULL DEFAULT 0,
                    recipients TEXT,
                    status TEXT NOT NULL,
                    detail TEXT,
                    generated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )

            for sno, name in STUDENTS.items():
                cur.execute(
                    """
                    INSERT INTO students (sno, name)
                    VALUES (%s, %s)
                    ON CONFLICT (sno)
                    DO UPDATE SET name = EXCLUDED.name
                    """,
                    (sno, name),
                )

            # Drop students who are no longer on the roster (a student who left
            # the class, say) so they stop appearing in staff lists. Anyone with
            # attendance history is kept: attendance_records cascades on delete,
            # and silently erasing records would be worse than a stale name.
            cur.execute(
                """
                DELETE FROM students
                WHERE NOT (sno = ANY(%s))
                  AND NOT EXISTS (
                      SELECT 1 FROM attendance_records
                      WHERE attendance_records.sno = students.sno
                  )
                RETURNING sno
                """,
                (list(STUDENTS),),
            )
            for row in cur.fetchall():
                log.warning(
                    "[roster] removed s.no %s - no longer on the roster", row["sno"]
                )

            cur.execute(
                """
                SELECT sno FROM students
                WHERE NOT (sno = ANY(%s))
                ORDER BY sno
                """,
                (list(STUDENTS),),
            )
            for row in cur.fetchall():
                log.warning(
                    "[roster] s.no %s is off the roster but has attendance "
                    "records, so it was kept",
                    row["sno"],
                )

            # One account per staff member, created with no password at all.
            # Each staff member sets their own the first time they sign in, so
            # nothing secret is ever written to the logs. DO NOTHING leaves an
            # account that has already been claimed exactly as it is.
            for staff_name in sorted(set(SUBJECTS.values())):
                cur.execute(
                    """
                    INSERT INTO staff_accounts
                        (staff_name, password_hash, must_change)
                    VALUES (%s, NULL, TRUE)
                    ON CONFLICT (staff_name) DO NOTHING
                    """,
                    (staff_name,),
                )

        conn.commit()


# ============================================================
# HELPERS
# ============================================================


def now_utc() -> datetime:
    return datetime.now(timezone.utc)


# ---- password hashing (PBKDF2-HMAC-SHA256, with legacy SHA-256 support) ----


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        PBKDF2_ITERATIONS,
    )
    return f"pbkdf2_sha256${PBKDF2_ITERATIONS}${salt}${digest.hex()}"


def _legacy_sha256(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password: str, stored: str) -> tuple[bool, bool]:
    """Return (is_valid, needs_rehash)."""
    if not stored:
        return False, False

    if stored.startswith("pbkdf2_sha256$"):
        try:
            _, iterations, salt, digest = stored.split("$")
            computed = hashlib.pbkdf2_hmac(
                "sha256",
                password.encode("utf-8"),
                bytes.fromhex(salt),
                int(iterations),
            )
            return secrets.compare_digest(computed.hex(), digest), False
        except Exception:
            return False, False

    # Legacy unsalted SHA-256 from an earlier version.
    valid = secrets.compare_digest(_legacy_sha256(password), stored)
    return valid, valid


# ---- simple in-process rate limiting ----


_rate_buckets: dict[str, list[float]] = {}


def rate_limit(key: str, limit: int, window_seconds: int) -> None:
    now = time.monotonic()
    bucket = [t for t in _rate_buckets.get(key, []) if now - t < window_seconds]
    if len(bucket) >= limit:
        raise HTTPException(
            status_code=429,
            detail="Too many attempts. Please wait a moment and try again.",
        )
    bucket.append(now)
    _rate_buckets[key] = bucket


# ---- email ----


def send_email(
    to_address: Optional[str],
    subject: str,
    body: str,
    attachment: Optional[tuple[str, bytes]] = None,
) -> bool:
    if not to_address:
        return False

    if not (SMTP_HOST and SMTP_FROM):
        log.warning(
            "[email-disabled] SMTP is not configured. Would send %r to %s. Body:\n%s",
            subject,
            to_address,
            body,
        )
        return False

    message = EmailMessage()
    message["From"] = SMTP_FROM
    message["To"] = to_address
    message["Subject"] = subject
    message.set_content(body)

    if attachment is not None:
        filename, data = attachment
        message.add_attachment(
            data,
            maintype="application",
            subtype=(
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            filename=filename,
        )

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=25) as server:
            server.starttls(context=ssl.create_default_context())
            if SMTP_USER:
                server.login(SMTP_USER, SMTP_PASSWORD or "")
            server.send_message(message)
        return True
    except Exception as exc:  # pragma: no cover - network dependent
        log.error("Failed to send email to %s: %s", to_address, exc)
        return False


def sno_from_admission_last4(last4: str) -> str:
    value = (last4 or "").strip()
    if not re.fullmatch(r"\d{4}", value):
        raise HTTPException(
            status_code=400,
            detail="Enter exactly the last 4 digits of your admission number.",
        )
    sno = ADMISSION_LAST4_TO_SNO.get(value)
    if sno is None:
        raise HTTPException(
            status_code=404,
            detail="Admission number not found.",
        )
    return sno


def make_code() -> str:
    return f"{secrets.randbelow(1_000_000):06d}"


def distance_meters(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6_371_000.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = (
        math.sin(dp / 2) ** 2
        + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    )
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


# ---- session helpers ----


def close_expired_sessions() -> None:
    """Close expired sessions and record Absent for anyone not Present."""
    current = now_utc()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE active = TRUE
                  AND expires_at <= %s
                FOR UPDATE
                """,
                (current,),
            )
            sessions = cur.fetchall()

            for session in sessions:
                cur.execute(
                    """
                    UPDATE attendance_sessions
                    SET active = FALSE, closed_at = %s
                    WHERE id = %s
                    """,
                    (current, session["id"]),
                )

                cur.execute(
                    """
                    INSERT INTO attendance_records
                    (
                        session_id, sno, subject, period,
                        attendance_date, marked_at, status
                    )
                    SELECT
                        %s, s.sno, %s, %s, %s, NULL, 'Absent'
                    FROM students s
                    WHERE NOT EXISTS (
                        SELECT 1
                        FROM attendance_records ar
                        WHERE ar.session_id = %s
                          AND ar.sno = s.sno
                          AND ar.status = 'Present'
                    )
                    ON CONFLICT (session_id, sno) DO NOTHING
                    """,
                    (
                        session["id"],
                        session["subject"],
                        session["period"],
                        session["started_at"].date(),
                        session["id"],
                    ),
                )

        conn.commit()


def get_active_session():
    """Return the most recently started usable session (any subject)."""
    current = now_utc()
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE active = TRUE
                  AND expires_at > %s
                ORDER BY started_at DESC
                LIMIT 1
                """,
                (current,),
            )
            return cur.fetchone()


def get_active_session_for_subject(subject: str):
    current = now_utc()
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE subject = %s
                  AND active = TRUE
                  AND expires_at > %s
                ORDER BY started_at DESC
                LIMIT 1
                """,
                (subject, current),
            )
            return cur.fetchone()


def get_latest_session_for_subject(subject: str):
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE subject = %s
                ORDER BY started_at DESC
                LIMIT 1
                """,
                (subject,),
            )
            return cur.fetchone()


# ============================================================
# BASIC
# ============================================================


def _build_info() -> dict:
    return {
        "message": "College Attendance Backend is running",
        "students": len(STUDENTS),
        "subjects": len(SUBJECTS),
        "code_validity_minutes": CODE_VALIDITY_MINUTES,
        "max_distance_meters": MAX_ATTENDANCE_DISTANCE_METERS,
        "semester_start": SEMESTER_START,
        "semester_end": SEMESTER_END,
        "erp_mode": ERP_MODE,
    }


@app.get("/api/info")
def api_info():
    return _build_info()


@app.get("/")
def root():
    # In the deployed bundle this origin also serves the Flutter app, so "/"
    # returns the UI. Locally, where no web build has been copied in, it falls
    # back to the service summary so the endpoint stays useful for smoke tests.
    index = WEB_DIR / "index.html"
    if index.is_file():
        return FileResponse(index)
    return _build_info()


@app.get("/health")
def health():
    try:
        with get_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1 AS ok")
                cur.fetchone()
    except Exception as exc:
        raise HTTPException(
            status_code=503, detail=f"Database unavailable: {exc}"
        )
    return {"status": "ok", "database": "connected"}


@app.get("/subjects")
def list_subjects():
    return [
        {"subject": subject, "staff": staff}
        for subject, staff in SUBJECTS.items()
    ]


# ============================================================
# STAFF AUTHENTICATION (one password per staff member)
# ============================================================


class StaffVerifyRequest(BaseModel):
    subject: str
    password: str


class StaffChangePasswordRequest(BaseModel):
    subject: str
    current_password: str
    new_password: str


class StaffSetupRequest(BaseModel):
    subject: str
    setup_code: str
    new_password: str


def _staff_account(staff_name: str):
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT * FROM staff_accounts WHERE staff_name = %s",
                (staff_name,),
            )
            return cur.fetchone()


@app.post("/staff/set-password")
def set_staff_password(request: StaffSetupRequest):
    """First-time setup: a staff member chooses their own password.

    The college hands out ``STAFF_SETUP_CODE`` to staff, and the account must
    still have no password, so this cannot be used to take over an account that
    someone is already using.
    """
    subject = request.subject.strip()
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    if not STAFF_SETUP_CODE:
        raise HTTPException(
            status_code=503,
            detail="Staff setup is switched off. Ask the administrator.",
        )

    if len(request.new_password) < MIN_PASSWORD_LENGTH + 2:
        raise HTTPException(
            status_code=400,
            detail="Password must contain at least 6 characters.",
        )

    staff_name = SUBJECTS[subject]
    rate_limit(f"staff-setup:{staff_name}", limit=10, window_seconds=300)

    if not secrets.compare_digest(
        request.setup_code.strip().encode("utf-8"), STAFF_SETUP_CODE.encode("utf-8")
    ):
        raise HTTPException(status_code=401, detail="Incorrect setup code.")

    account = _staff_account(staff_name)
    if not account:
        raise HTTPException(
            status_code=500, detail="Staff account is not configured."
        )

    if account["password_hash"]:
        raise HTTPException(
            status_code=409,
            detail=(
                "This staff account already has a password. "
                "Sign in instead, or ask the administrator to reset it."
            ),
        )

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE staff_accounts
                SET password_hash = %s,
                    must_change = FALSE,
                    updated_at = NOW()
                WHERE staff_name = %s
                """,
                (hash_password(request.new_password), staff_name),
            )
        conn.commit()

    log.info("[staff-setup] %s chose their own password", staff_name)
    return {"success": True, "subject": subject, "staff": staff_name}


@app.post("/staff/verify-subject")
def verify_subject_password(request: StaffVerifyRequest):
    subject = request.subject.strip()
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    staff_name = SUBJECTS[subject]
    rate_limit(f"staff-login:{staff_name}", limit=10, window_seconds=300)

    account = _staff_account(staff_name)
    if not account:
        raise HTTPException(
            status_code=500,
            detail="Staff account is not configured.",
        )

    if not account["password_hash"]:
        raise HTTPException(
            status_code=409,
            detail=(
                "No password is set for this staff account yet. "
                "Create one first, then sign in."
            ),
        )

    valid, needs_rehash = verify_password(
        request.password, account["password_hash"]
    )
    if not valid:
        raise HTTPException(status_code=401, detail="Incorrect password.")

    if needs_rehash:
        with get_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE staff_accounts SET password_hash = %s WHERE staff_name = %s",
                    (hash_password(request.password), staff_name),
                )
            conn.commit()

    return {
        "success": True,
        "subject": subject,
        "staff": staff_name,
        "must_change_password": bool(account["must_change"]),
    }


@app.post("/staff/change-password")
def change_staff_password(request: StaffChangePasswordRequest):
    subject = request.subject.strip()
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    if len(request.new_password) < MIN_PASSWORD_LENGTH + 2:
        raise HTTPException(
            status_code=400,
            detail="New password must contain at least 6 characters.",
        )

    staff_name = SUBJECTS[subject]
    account = _staff_account(staff_name)
    if not account:
        raise HTTPException(
            status_code=500, detail="Staff account is not configured."
        )

    if not account["password_hash"]:
        raise HTTPException(
            status_code=409,
            detail=(
                "No password is set for this staff account yet. "
                "Create one first, then sign in."
            ),
        )

    valid, _ = verify_password(request.current_password, account["password_hash"])
    if not valid:
        raise HTTPException(
            status_code=401, detail="Current password is incorrect."
        )

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE staff_accounts
                SET password_hash = %s,
                    must_change = FALSE,
                    updated_at = NOW()
                WHERE staff_name = %s
                """,
                (hash_password(request.new_password), staff_name),
            )
        conn.commit()

    return {"success": True, "subject": subject, "staff": staff_name}


# ============================================================
# START ATTENDANCE
# ============================================================


@app.post("/attendance/start")
def start_attendance(
    subject: str = Query(...),
    period: Optional[str] = Query(None),
    password: Optional[str] = Query(None),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    if password is not None:
        account = _staff_account(SUBJECTS[subject])
        if account and not verify_password(password, account["password_hash"])[0]:
            raise HTTPException(status_code=401, detail="Incorrect password.")

    close_expired_sessions()

    current = now_utc()
    expires = current + timedelta(minutes=CODE_VALIDITY_MINUTES)

    if period is None or not period.strip():
        with get_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT COUNT(*) AS count
                    FROM attendance_sessions
                    WHERE subject = %s
                      AND started_at::date = %s
                    """,
                    (subject, current.date()),
                )
                number = int(cur.fetchone()["count"]) + 1
        period = f"Session {number}"
    else:
        period = period.strip()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT pg_advisory_xact_lock(%s)", (918273645,))

            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE subject = %s
                  AND active = TRUE
                  AND expires_at > %s
                ORDER BY started_at DESC
                LIMIT 1
                """,
                (subject, current),
            )
            active = cur.fetchone()

            if active:
                raise HTTPException(
                    status_code=409,
                    detail=(
                        f"An attendance session is already active for "
                        f"{active['subject']} ({active['period']})."
                    ),
                )

            cur.execute(
                """
                INSERT INTO attendance_sessions
                (subject, staff_name, period, started_at, expires_at, active)
                VALUES (%s, %s, %s, %s, %s, TRUE)
                RETURNING id
                """,
                (subject, SUBJECTS[subject], period, current, expires),
            )
            session_id = cur.fetchone()["id"]

            used_codes: set[str] = set()
            for sno in sorted(STUDENTS, key=lambda value: int(value)):
                code = make_code()
                while code in used_codes:
                    code = make_code()
                used_codes.add(code)
                cur.execute(
                    """
                    INSERT INTO attendance_codes (session_id, sno, code)
                    VALUES (%s, %s, %s)
                    """,
                    (session_id, sno, code),
                )

        conn.commit()

    return {
        "success": True,
        "session_id": session_id,
        "subject": subject,
        "staff": SUBJECTS[subject],
        "period": period,
        "started_at": current.isoformat(),
        "expires_at": expires.isoformat(),
        "valid_for_minutes": CODE_VALIDITY_MINUTES,
    }


# ============================================================
# STUDENT'S OWN CODE
# ============================================================


@app.get("/attendance/my-code/{sno}")
def get_my_code(sno: str, subject: Optional[str] = Query(None)):
    if sno not in STUDENTS:
        raise HTTPException(status_code=404, detail="Student S.No. not found.")

    session = (
        get_active_session_for_subject(subject)
        if subject
        else get_active_session()
    )
    if not session:
        raise HTTPException(
            status_code=404, detail="No active attendance session."
        )

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT code
                FROM attendance_codes
                WHERE session_id = %s AND sno = %s
                """,
                (session["id"], sno),
            )
            result = cur.fetchone()

    if not result:
        raise HTTPException(status_code=404, detail="Attendance code not found.")

    return {
        "sno": sno,
        "name": STUDENTS[sno],
        "code": result["code"],
        "subject": session["subject"],
        "period": session["period"],
        "expires_at": session["expires_at"].isoformat(),
    }


# ============================================================
# MARK ATTENDANCE
# ============================================================


class MarkAttendanceRequest(BaseModel):
    sno: str
    code: str
    device_id: str = Field(min_length=8, max_length=200)
    latitude: Optional[float] = Field(default=None)
    longitude: Optional[float] = Field(default=None)


@app.post("/attendance/mark")
def mark_attendance(request: MarkAttendanceRequest):
    sno = request.sno.strip()
    code = request.code.strip()

    if sno not in STUDENTS:
        raise HTTPException(status_code=404, detail="Student S.No. not found.")

    session = get_active_session()
    if not session:
        raise HTTPException(status_code=410, detail="Attendance session has ended.")

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id
                FROM attendance_codes
                WHERE session_id = %s AND sno = %s AND code = %s
                """,
                (session["id"], sno, code),
            )
            if not cur.fetchone():
                raise HTTPException(
                    status_code=401, detail="Invalid attendance code."
                )

            device_id = request.device_id.strip()
            if not device_id:
                raise HTTPException(
                    status_code=400, detail="Device identification is required."
                )

            cur.execute(
                """
                SELECT sno
                FROM attendance_device_locks
                WHERE session_id = %s AND device_id = %s
                FOR UPDATE
                """,
                (session["id"], device_id),
            )
            existing_device_lock = cur.fetchone()
            if existing_device_lock:
                if existing_device_lock["sno"] == sno:
                    raise HTTPException(
                        status_code=409,
                        detail="Attendance is already marked from this device for this session.",
                    )
                raise HTTPException(
                    status_code=409,
                    detail="This device has already been used to mark attendance for this session.",
                )

            cur.execute(
                """
                SELECT id
                FROM attendance_records
                WHERE session_id = %s AND sno = %s AND status = 'Present'
                """,
                (session["id"], sno),
            )
            if cur.fetchone():
                raise HTTPException(
                    status_code=409, detail="Attendance is already marked."
                )

            distance = None
            if LOCATION_REQUIRED:
                if request.latitude is None or request.longitude is None:
                    raise HTTPException(
                        status_code=400, detail="Location is required."
                    )

                distance = distance_meters(
                    request.latitude,
                    request.longitude,
                    CLASSROOM_LATITUDE,
                    CLASSROOM_LONGITUDE,
                )

                if distance > MAX_ATTENDANCE_DISTANCE_METERS:
                    raise HTTPException(
                        status_code=403,
                        detail=(
                            "You are outside the classroom area. "
                            f"Distance: {distance:.1f} m."
                        ),
                    )

            marked_at = now_utc()

            cur.execute(
                """
                INSERT INTO attendance_records
                (
                    session_id, sno, subject, period,
                    attendance_date, marked_at, status,
                    latitude, longitude, distance_meters
                )
                VALUES (%s, %s, %s, %s, %s, %s, 'Present', %s, %s, %s)
                """,
                (
                    session["id"],
                    sno,
                    session["subject"],
                    session["period"],
                    marked_at.date(),
                    marked_at,
                    request.latitude,
                    request.longitude,
                    distance,
                ),
            )

            cur.execute(
                """
                INSERT INTO attendance_device_locks
                (session_id, device_id, sno, marked_at)
                VALUES (%s, %s, %s, %s)
                """,
                (session["id"], device_id, sno, marked_at),
            )

        conn.commit()

    return {
        "success": True,
        "sno": sno,
        "name": STUDENTS[sno],
        "subject": session["subject"],
        "period": session["period"],
        "date": marked_at.date().isoformat(),
        "status": "Present",
    }


# ============================================================
# ABSENTEES
# ============================================================


@app.get("/attendance/absentees")
def get_absentees(subject: str = Query(...)):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()

    active = get_active_session_for_subject(subject)
    session = active or get_latest_session_for_subject(subject)

    if not session:
        return {
            "subject": subject,
            "period": None,
            "date": None,
            "session_id": None,
            "active": False,
            "absentees": [],
        }

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT s.sno, s.name
                FROM students s
                WHERE NOT EXISTS (
                    SELECT 1
                    FROM attendance_records ar
                    WHERE ar.session_id = %s
                      AND ar.sno = s.sno
                      AND ar.status = 'Present'
                )
                ORDER BY s.sno::INTEGER ASC
                """,
                (session["id"],),
            )
            absentees = cur.fetchall()

    return {
        "subject": session["subject"],
        "period": session["period"],
        "date": session["started_at"].date().isoformat(),
        "session_id": session["id"],
        "active": session["active"],
        "expires_at": session["expires_at"].isoformat(),
        "absentee_count": len(absentees),
        "absentees": absentees,
    }


# ============================================================
# ABSENTEE HISTORY (kept for the staff "just finished" view)
# ============================================================


@app.get("/attendance/absentee-history")
def get_absentee_history(subject: str = Query(...)):
    """Latest session's absentee list for a subject.

    Unlike the original version this does NOT hide the list after a fixed
    window; attendance is meant to stay visible for the whole semester.
    A newer active session still suppresses the older list.
    """
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE subject = %s
                ORDER BY started_at DESC
                LIMIT 1
                """,
                (subject,),
            )
            session = cur.fetchone()

            if not session:
                return {
                    "subject": subject,
                    "date": None,
                    "session_id": None,
                    "visible_until": None,
                    "absentees": [],
                }

            if session["active"]:
                return {
                    "subject": subject,
                    "date": None,
                    "session_id": session["id"],
                    "visible_until": None,
                    "absentees": [],
                }

            cur.execute(
                """
                SELECT s.sno, s.name
                FROM students s
                WHERE NOT EXISTS (
                    SELECT 1
                    FROM attendance_records ar
                    WHERE ar.session_id = %s
                      AND ar.sno = s.sno
                      AND ar.status = 'Present'
                )
                ORDER BY s.sno::INTEGER ASC
                """,
                (session["id"],),
            )
            rows = cur.fetchall()

    return {
        "subject": subject,
        "date": session["started_at"].date().isoformat(),
        "session_id": session["id"],
        "visible_until": None,
        "absentees": rows,
    }


# ============================================================
# SEMESTER REPORTS
# ============================================================


def _parse_date(value: Optional[str], field: str):
    if value is None or not value.strip():
        return None
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail=f"{field} must be in YYYY-MM-DD format.",
        )


def _parse_month(value: Optional[str], field: str = "month"):
    if value is None or not value.strip():
        return None
    try:
        return datetime.strptime(value.strip(), "%Y-%m").date()
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail=f"{field} must be in YYYY-MM format.",
        )


def _next_month(month_start: date) -> date:
    if month_start.month == 12:
        return date(month_start.year + 1, 1, 1)
    return date(month_start.year, month_start.month + 1, 1)


def _month_label(month_start: date) -> str:
    return month_start.strftime("%B %Y")


def _semester_rows(
    cur,
    subject: str,
    start_date=None,
    end_date=None,
) -> list[dict]:
    query = """
        SELECT
            s.sno,
            s.name,
            COALESCE(SUM(CASE WHEN ar.status = 'Present' THEN 1 ELSE 0 END), 0)
                AS present,
            COALESCE(SUM(CASE WHEN ar.status = 'Absent' THEN 1 ELSE 0 END), 0)
                AS absent
        FROM students s
        LEFT JOIN attendance_records ar
            ON ar.sno = s.sno
           AND ar.subject = %s
    """
    params: list = [subject]
    if start_date is not None:
        query += " AND ar.attendance_date >= %s"
        params.append(start_date)
    if end_date is not None:
        query += " AND ar.attendance_date <= %s"
        params.append(end_date)

    query += """
        GROUP BY s.sno, s.name
        ORDER BY s.sno::INTEGER ASC
    """
    cur.execute(query, params)
    rows = cur.fetchall()

    result = []
    for row in rows:
        present = int(row["present"])
        absent = int(row["absent"])
        result.append(
            {
                "sno": row["sno"],
                "name": row["name"],
                "present": present,
                "absent": absent,
            }
        )
    return result


@app.get("/attendance/sessions")
def list_subject_sessions(
    subject: str = Query(...),
    from_date: Optional[str] = Query(None, alias="from"),
    to_date: Optional[str] = Query(None, alias="to"),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()

    start = _parse_date(from_date, "from")
    end = _parse_date(to_date, "to")

    query = """
        SELECT
            ses.id,
            ses.period,
            ses.started_at,
            ses.expires_at,
            ses.closed_at,
            ses.active,
            COUNT(ar.id) FILTER (WHERE ar.status = 'Present') AS present,
            COUNT(ar.id) FILTER (WHERE ar.status = 'Absent') AS absent
        FROM attendance_sessions ses
        LEFT JOIN attendance_records ar ON ar.session_id = ses.id
        WHERE ses.subject = %s
    """
    params: list = [subject]
    if start is not None:
        query += " AND ses.started_at::date >= %s"
        params.append(start)
    if end is not None:
        query += " AND ses.started_at::date <= %s"
        params.append(end)
    query += " GROUP BY ses.id ORDER BY ses.started_at DESC"

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, params)
            sessions = cur.fetchall()

    return {
        "subject": subject,
        "staff": SUBJECTS[subject],
        "sessions": [
            {
                "session_id": s["id"],
                "period": s["period"],
                "started_at": s["started_at"].isoformat(),
                "expires_at": s["expires_at"].isoformat(),
                "closed_at": s["closed_at"].isoformat()
                if s["closed_at"]
                else None,
                "active": s["active"],
                "present": int(s["present"]),
                "absent": int(s["absent"]),
            }
            for s in sessions
        ],
    }


@app.get("/attendance/summary")
def attendance_summary(
    subject: str = Query(...),
    from_date: Optional[str] = Query(None, alias="from"),
    to_date: Optional[str] = Query(None, alias="to"),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()

    start = _parse_date(from_date, "from")
    end = _parse_date(to_date, "to")

    with get_connection() as conn:
        with conn.cursor() as cur:
            rows = _semester_rows(cur, subject, start, end)

    return {
        "subject": subject,
        "staff": SUBJECTS[subject],
        "from": start.isoformat() if start else None,
        "to": end.isoformat() if end else None,
        "students": rows,
    }


# ============================================================
# MONTHLY ABSENTEES
# ============================================================


def _require_subject_password(subject: str, password: Optional[str]) -> None:
    if password is None:
        return
    account = _staff_account(SUBJECTS[subject])
    if account and not verify_password(password, account["password_hash"])[0]:
        raise HTTPException(status_code=401, detail="Incorrect password.")


def _monthly_absentee_data(
    subject: str,
    month_start: Optional[date] = None,
) -> dict:
    """Absentees for one subject, session by session, for a single month.

    A student counts as absent for a session when there is no Present record
    for it. That also covers a session that is still running, because its
    Absent rows are only written when it closes.
    """
    close_expired_sessions()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    to_char(started_at, 'YYYY-MM') AS month,
                    COUNT(*) AS sessions
                FROM attendance_sessions
                WHERE subject = %s
                GROUP BY 1
                ORDER BY 1 DESC
                """,
                (subject,),
            )
            available_months = [
                {
                    "month": row["month"],
                    "sessions": int(row["sessions"]),
                }
                for row in cur.fetchall()
            ]

            if month_start is None:
                month_start = (
                    datetime.strptime(
                        available_months[0]["month"], "%Y-%m"
                    ).date()
                    if available_months
                    else now_utc().date().replace(day=1)
                )

            month_end = _next_month(month_start)

            cur.execute("SELECT COUNT(*) AS count FROM students")
            student_count = int(cur.fetchone()["count"])

            cur.execute(
                """
                SELECT
                    ses.id,
                    ses.period,
                    ses.started_at,
                    ses.closed_at,
                    ses.active,
                    COUNT(ar.id) FILTER (WHERE ar.status = 'Present')
                        AS present
                FROM attendance_sessions ses
                LEFT JOIN attendance_records ar ON ar.session_id = ses.id
                WHERE ses.subject = %s
                  AND ses.started_at::date >= %s
                  AND ses.started_at::date < %s
                GROUP BY ses.id
                ORDER BY ses.started_at ASC
                """,
                (subject, month_start, month_end),
            )
            session_rows = cur.fetchall()

            cur.execute(
                """
                SELECT ses.id AS session_id, s.sno, s.name
                FROM attendance_sessions ses
                CROSS JOIN students s
                WHERE ses.subject = %s
                  AND ses.started_at::date >= %s
                  AND ses.started_at::date < %s
                  AND NOT EXISTS (
                      SELECT 1
                      FROM attendance_records ar
                      WHERE ar.session_id = ses.id
                        AND ar.sno = s.sno
                        AND ar.status = 'Present'
                  )
                ORDER BY ses.id ASC, s.sno::INTEGER ASC
                """,
                (subject, month_start, month_end),
            )
            absentees_by_session: dict[int, list[dict]] = {}
            for row in cur.fetchall():
                absentees_by_session.setdefault(row["session_id"], []).append(
                    {"sno": row["sno"], "name": row["name"]}
                )

            cur.execute(
                """
                WITH month_sessions AS (
                    SELECT id
                    FROM attendance_sessions
                    WHERE subject = %s
                      AND started_at::date >= %s
                      AND started_at::date < %s
                )
                SELECT
                    s.sno,
                    s.name,
                    COUNT(ms.id) AS sessions,
                    COUNT(ar.session_id) AS present
                FROM students s
                CROSS JOIN month_sessions ms
                LEFT JOIN attendance_records ar
                    ON ar.session_id = ms.id
                   AND ar.sno = s.sno
                   AND ar.status = 'Present'
                GROUP BY s.sno, s.name
                ORDER BY s.sno::INTEGER ASC
                """,
                (subject, month_start, month_end),
            )
            totals = [
                {
                    "sno": row["sno"],
                    "name": row["name"],
                    "sessions": int(row["sessions"]),
                    "present": int(row["present"]),
                    "absent": int(row["sessions"]) - int(row["present"]),
                }
                for row in cur.fetchall()
            ]

    sessions = []
    absent_total = 0
    for row in session_rows:
        absentees = absentees_by_session.get(row["id"], [])
        absent = len(absentees)
        absent_total += absent
        sessions.append(
            {
                "session_id": row["id"],
                "date": row["started_at"].date().isoformat(),
                "period": row["period"],
                "started_at": row["started_at"].isoformat(),
                "closed_at": row["closed_at"].isoformat()
                if row["closed_at"]
                else None,
                "active": row["active"],
                "present": int(row["present"]),
                "absent": absent,
                "absentees": absentees,
            }
        )

    return {
        "subject": subject,
        "staff": SUBJECTS[subject],
        "month": month_start.strftime("%Y-%m"),
        "month_label": _month_label(month_start),
        "student_count": student_count,
        "session_count": len(sessions),
        "absent_total": absent_total,
        "available_months": available_months,
        "sessions": sessions,
        "totals": totals,
    }


@app.get("/attendance/monthly-absentees")
def monthly_absentees(
    subject: str = Query(...),
    month: Optional[str] = Query(None),
    password: Optional[str] = Query(None),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    _require_subject_password(subject, password)

    return _monthly_absentee_data(subject, _parse_month(month))


# ============================================================
# CLOSE ATTENDANCE MANUALLY
# ============================================================


@app.post("/attendance/close")
def close_attendance(
    subject: str = Query(...),
    password: Optional[str] = Query(None),
):
    if password is not None:
        account = _staff_account(SUBJECTS.get(subject, ""))
        if account and not verify_password(password, account["password_hash"])[0]:
            raise HTTPException(status_code=401, detail="Incorrect password.")

    close_expired_sessions()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE subject = %s AND active = TRUE
                ORDER BY started_at DESC
                LIMIT 1
                FOR UPDATE
                """,
                (subject,),
            )
            session = cur.fetchone()
            if not session:
                raise HTTPException(
                    status_code=404, detail="No active session found."
                )

            current = now_utc()
            cur.execute(
                """
                UPDATE attendance_sessions
                SET active = FALSE, closed_at = %s
                WHERE id = %s
                """,
                (current, session["id"]),
            )

            cur.execute(
                """
                INSERT INTO attendance_records
                (
                    session_id, sno, subject, period,
                    attendance_date, marked_at, status
                )
                SELECT %s, s.sno, %s, %s, %s, NULL, 'Absent'
                FROM students s
                WHERE NOT EXISTS (
                    SELECT 1
                    FROM attendance_records ar
                    WHERE ar.session_id = %s
                      AND ar.sno = s.sno
                      AND ar.status = 'Present'
                )
                ON CONFLICT (session_id, sno) DO NOTHING
                """,
                (
                    session["id"],
                    session["subject"],
                    session["period"],
                    session["started_at"].date(),
                    session["id"],
                ),
            )

        conn.commit()

    return {
        "success": True,
        "session_id": session["id"],
        "subject": session["subject"],
        "period": session["period"],
        "closed_at": current.isoformat(),
    }


# ============================================================
# STUDENT HISTORY
# ============================================================


@app.get("/attendance/history/{sno}")
def attendance_history(sno: str):
    if sno not in STUDENTS:
        raise HTTPException(status_code=404, detail="Student not found.")

    close_expired_sessions()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT sno, subject, period, attendance_date, marked_at, status
                FROM attendance_records
                WHERE sno = %s
                ORDER BY attendance_date DESC, session_id DESC
                """,
                (sno,),
            )
            records = cur.fetchall()

    for record in records:
        record["attendance_date"] = record["attendance_date"].isoformat()
        if record["marked_at"] is not None:
            record["marked_at"] = record["marked_at"].isoformat()

    return {"sno": sno, "name": STUDENTS[sno], "records": records}


# ============================================================
# STUDENT ACCOUNTS
# ============================================================


class StudentSetupRequest(BaseModel):
    admission_last4: str
    password: str
    # An email only ever existed to receive a reset code, and nothing sends one
    # now. It stays optional rather than blocking sign-up on unused data.
    email: str = ""


class StudentLoginRequest(BaseModel):
    admission_last4: str
    password: str


@app.post("/student/set-password")
def set_student_password(request: StudentSetupRequest):
    sno = sno_from_admission_last4(request.admission_last4)

    if len(request.password) < MIN_PASSWORD_LENGTH:
        raise HTTPException(
            status_code=400,
            detail=f"Password must contain at least {MIN_PASSWORD_LENGTH} characters.",
        )

    email = (request.email or "").strip().lower()
    if email and not EMAIL_RE.match(email):
        raise HTTPException(status_code=400, detail="Enter a valid email address.")

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT sno FROM student_passwords WHERE sno = %s", (sno,))
            if cur.fetchone():
                raise HTTPException(
                    status_code=409,
                    detail="Password already created for this student.",
                )
            cur.execute(
                """
                INSERT INTO student_passwords (sno, password_hash, email)
                VALUES (%s, %s, %s)
                """,
                (sno, hash_password(request.password), email or None),
            )
        conn.commit()

    return {"success": True, "sno": sno, "name": STUDENTS[sno]}


@app.post("/student/login")
def student_login(request: StudentLoginRequest):
    sno = sno_from_admission_last4(request.admission_last4)
    rate_limit(f"student-login:{sno}", limit=10, window_seconds=300)

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT password_hash FROM student_passwords WHERE sno = %s",
                (sno,),
            )
            record = cur.fetchone()

    if not record:
        raise HTTPException(status_code=404, detail="Password not created yet.")

    valid, needs_rehash = verify_password(
        request.password, record["password_hash"]
    )
    if not valid:
        raise HTTPException(status_code=401, detail="Incorrect password.")

    if needs_rehash:
        with get_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE student_passwords SET password_hash = %s WHERE sno = %s",
                    (hash_password(request.password), sno),
                )
            conn.commit()

    return {"success": True, "sno": sno, "name": STUDENTS[sno]}




# ============================================================
# EXPORT (per-student semester summary)
# ============================================================


def _build_summary_workbook(subject: str, start=None, end=None) -> BytesIO:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Attendance Summary"

    sheet.append(
        [
            "S.No.",
            "Name",
            "Present",
            "Absent",
        ]
    )

    with get_connection() as conn:
        with conn.cursor() as cur:
            rows = _semester_rows(cur, subject, start, end)

    for row in rows:
        sheet.append(
            [
                int(row["sno"]),
                row["name"],
                row["present"],
                row["absent"],
            ]
        )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def _safe_filename(subject: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", subject).strip("_").lower()
    return f"attendance_{cleaned or 'subject'}.xlsx"


def _safe_month_filename(subject: str, month: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", subject).strip("_").lower()
    return f"absentees_{cleaned or 'subject'}_{month}.xlsx"


def _build_monthly_absentee_workbook(
    subject: str,
    month_start: Optional[date],
) -> tuple[BytesIO, dict]:
    data = _monthly_absentee_data(subject, month_start)

    workbook = Workbook()
    daily = workbook.active
    daily.title = "Daily absentees"
    daily.append(["Date", "Period", "S.No.", "Name"])

    for session in data["sessions"]:
        for person in session["absentees"]:
            daily.append(
                [
                    session["date"],
                    session["period"],
                    int(person["sno"]),
                    person["name"],
                ]
            )

    summary = workbook.create_sheet("Monthly summary")
    summary.append(["S.No.", "Name", "Sessions", "Present", "Absent"])

    for row in data["totals"]:
        summary.append(
            [
                int(row["sno"]),
                row["name"],
                row["sessions"],
                row["present"],
                row["absent"],
            ]
        )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, data


@app.get("/attendance/monthly-absentees/export")
def export_monthly_absentees(
    subject: str = Query(...),
    month: Optional[str] = Query(None),
    password: Optional[str] = Query(None),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    _require_subject_password(subject, password)

    output, data = _build_monthly_absentee_workbook(
        subject, _parse_month(month)
    )
    filename = _safe_month_filename(subject, data["month"])

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/attendance/export")
def export_attendance(
    subject: str = Query(...),
    from_date: Optional[str] = Query(None, alias="from"),
    to_date: Optional[str] = Query(None, alias="to"),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()

    start = _parse_date(from_date, "from")
    end = _parse_date(to_date, "to")

    output = _build_summary_workbook(subject, start, end)
    filename = _safe_filename(subject)

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ============================================================
# ERP UPDATE
# ============================================================


def _push_to_erp(subject: str, filename: str, data: bytes) -> tuple[bool, str]:
    if ERP_MODE == "api" and ERP_API_URL:
        import json
        import urllib.request

        request = urllib.request.Request(
            ERP_API_URL,
            data=data,
            method="POST",
            headers={
                "Content-Type": (
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                "X-Subject": subject,
                "X-Filename": filename,
                "Authorization": f"Bearer {ERP_API_KEY}" if ERP_API_KEY else "",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                body = response.read().decode("utf-8", "replace")
                return 200 <= response.status < 300, body[:500]
        except Exception as exc:  # pragma: no cover - network dependent
            return False, str(exc)

    # file mode: email the workbook to the college ERP address.
    if not ERP_UPLOAD_EMAIL:
        return False, "ERP_UPLOAD_EMAIL is not configured."

    sent = send_email(
        ERP_UPLOAD_EMAIL,
        f"Attendance update - {subject}",
        (
            f"Attached is the semester attendance summary for {subject}.\n\n"
            "Generated by the College Attendance app."
        ),
        attachment=(filename, data),
    )
    return sent, "emailed" if sent else "email failed"


@app.post("/attendance/erp-update")
def erp_update(
    subject: str = Query(...),
    from_date: Optional[str] = Query(None, alias="from"),
    to_date: Optional[str] = Query(None, alias="to"),
):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()

    start = _parse_date(from_date, "from")
    end = _parse_date(to_date, "to")

    output = _build_summary_workbook(subject, start, end)
    data = output.getvalue()
    filename = _safe_filename(subject)

    ok, detail = _push_to_erp(subject, filename, data)

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO erp_sync_log
                (
                    subject, mode, period_start, period_end,
                    student_count, recipients, status, detail
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    subject,
                    ERP_MODE,
                    start,
                    end,
                    len(STUDENTS),
                    ERP_UPLOAD_EMAIL if ERP_MODE != "api" else ERP_API_URL,
                    "success" if ok else "failed",
                    detail,
                ),
            )
        conn.commit()

    if not ok:
        raise HTTPException(
            status_code=502,
            detail=f"Could not deliver to the ERP ({detail}). "
            "You can still download the Excel file.",
        )

    return {
        "success": True,
        "subject": subject,
        "mode": ERP_MODE,
        "delivered_to": ERP_UPLOAD_EMAIL
        if ERP_MODE != "api"
        else ERP_API_URL,
        "filename": filename,
        "students": len(STUDENTS),
    }


# ============================================================
# DEVELOPMENT-ONLY RESET (disabled unless RESET_DB_ENABLED=true)
# ============================================================


@app.delete("/development/reset-database")
def development_reset_database():
    if not RESET_DB_ENABLED:
        raise HTTPException(
            status_code=403,
            detail="Reset is disabled. Set RESET_DB_ENABLED=true to allow it.",
        )

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "TRUNCATE attendance_records, attendance_codes, "
                "attendance_sessions, attendance_device_locks, "
                "student_passwords, erp_sync_log "
                "RESTART IDENTITY CASCADE"
            )
        conn.commit()

    return {"success": True, "message": "Development attendance data reset."}


# ============================================================
# STATIC FLUTTER WEB APP
# ============================================================
#
# Mounted last so every API route above keeps priority. Serving the UI and the
# API from one origin removes cross-origin requests entirely, which matters
# because the Catalyst gateway answers CORS preflights itself and never passes
# them through to this app.

if (WEB_DIR / "index.html").is_file():
    app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
    log.info("Serving the Flutter web app from %s", WEB_DIR)
else:
    log.info("No Flutter web build at %s - serving the API only", WEB_DIR)


# ============================================================
# DIRECT EXECUTION (used by Zoho Catalyst AppSail)
# ============================================================

if __name__ == "__main__":
    # AppSail starts the service with a plain command and tells us which port
    # to bind through X_ZOHO_CATALYST_LISTEN_PORT. Handling that here avoids
    # having to interpolate the variable inside a shell-wrapped startup
    # command, which is easy to get wrong and awkward to debug.
    import uvicorn

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.getenv("X_ZOHO_CATALYST_LISTEN_PORT", "9000")),
    )
