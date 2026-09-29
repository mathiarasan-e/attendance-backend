from __future__ import annotations

import hashlib
import math
import re
import secrets
from datetime import datetime, timedelta, timezone
from io import BytesIO
from typing import Optional

import psycopg
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from pydantic import BaseModel, Field
import os


# ============================================================
# DATABASE / APP CONFIG
# ============================================================

# CHANGE ONLY YOUR POSTGRESQL PASSWORD.

DATABASE_URL = os.getenv("DATABASE_URL")

CODE_VALIDITY_MINUTES = 1
ABSENTEE_VISIBILITY_MINUTES = 45
MAX_ATTENDANCE_DISTANCE_METERS = 20.0

# Keep False while we are testing the backend before the
# classroom GPS coordinates are configured.
# Change to True after Flutter sends GPS coordinates.
LOCATION_REQUIRED = True

# Put the actual classroom coordinates here before enabling
# LOCATION_REQUIRED.
CLASSROOM_LATITUDE = 12.9868169
CLASSROOM_LONGITUDE = 79.9724955


# ============================================================
# STUDENTS
# ============================================================

STUDENTS: dict[str, str] = {
    "120": 'JOAN MARSALINE A',
    "121": 'JOEL S',
    "122": 'K DEVADHARSHINI',
    "123": 'K DHANANJAY',
    "124": 'K M LEKHAA VARSHINI',
    "125": 'K R HARSHINI',
    "126": 'K SHAMITHA',
    "127": 'KAILASH SURYA P',
    "128": 'KAMALA KANNAN S',
    "129": 'KAMALI S B',
    "130": 'KANISHKA A',
    "131": 'KAPILVENKAT K',
    "132": 'KARTHICK S',
    "133": 'KARTHIGA D',
    "134": 'KARTHIGA S',
    "135": 'KARUNYA G',
    "136": 'KAVINAYA J',
    "137": 'KAVIYA G',
    "138": 'KAVIYA J',
    "139": 'KEERTHANA R',
    "140": 'KEERTHANA S',
    "141": 'KESAVARAM P',
    "142": 'KIRUPANITHI S',
    "143": 'KISHORE K',
    "144": 'KRISHITHA J',
    "145": 'KRISHNA A M',
    "146": 'KRITHICK S',
    "147": 'LAKSHAYA G',
    "148": 'LAKSHAYA NAGARAJAN',
    "149": 'LAKSHIKHAA SRI B',
    "150": 'LAKSHMAN M',
    "151": 'LAKSHMI PRIYA S',
    "152": 'LALITH CHANDAR M',
    "153": 'LAURA SUSAN C',
    "154": 'LAVANYA R',
    "155": 'LIVISHARAN E',
    "156": 'LOKESH D',
    "157": 'LOKESHWARI V',
    "158": 'M GUNADHARSHINI',
    "159": 'M S ABISHIEK',
    "160": 'M YASHWANTH SAI',
    "161": 'MADHAN RAJ D',
    "162": 'MALIK KANI M K',
    "163": 'MANASA S',
    "164": 'MANSI R',
    "165": 'MATHANGI P',
    "166": 'MATHAVAN R D',
    "167": 'MATHAVAN S',
    "168": 'MATHIARASAN E',
    "169": 'MEENAKSHI A',
    "170": 'MITHRAA A',
    "171": 'MOHAMED ANEESUR RAHMAN A',
    "172": 'MOHAMED JASIM M R',
    "173": 'MOHAMMED NADEEM K S',
    "174": 'MOHANKUMAR M',
    "175": 'MOKSHITA K G',
    "176": 'MONISH V',
    "177": 'MONISHA V',
    "178": 'MOONISHHA T',
    "179": 'MOURIYA M',
    "180": 'MUGUNDHASARAVANAN D',
    "181": 'MUHAMMED JIYAD A',
    "182": 'MUKUL V P',
    "183": 'NAREN V',
}


REGISTRATION_LAST4_TO_SNO: dict[str, str] = {
    "0553": "120",
    "1302": "121",
    "0398": "122",
    "0683": "123",
    "0707": "124",
    "1219": "125",
    "0517": "126",
    "0448": "127",
    "0184": "128",
    "0475": "129",
    "0215": "130",
    "0776": "131",
    "0504": "132",
    "0500": "133",
    "0167": "134",
    "1026": "135",
    "0777": "136",
    "0921": "137",
    "0991": "138",
    "0005": "139",
    "1466": "140",
    "1470": "141",
    "1354": "142",
    "0756": "143",
    "0491": "144",
    "0209": "145",
    "0718": "146",
    "1340": "147",
    "0652": "148",
    "1452": "149",
    "1402": "150",
    "1353": "151",
    "0898": "152",
    "0380": "153",
    "1050": "154",
    "0359": "155",
    "1261": "156",
    "0164": "157",
    "0358": "158",
    "0351": "159",
    "0365": "160",
    "0218": "161",
    "0395": "162",
    "0345": "163",
    "0724": "164",
    "0045": "165",
    "0378": "166",
    "0582": "167",
    "0559": "168",
    "0021": "169",
    "0779": "170",
    "0082": "171",
    "0079": "172",
    "0828": "173",
    "0927": "174",
    "1443": "175",
    "0354": "176",
    "0174": "177",
    "1372": "178",
    "0469": "179",
    "0393": "180",
    "1436": "181",
    "0125": "182",
    "0883": "183",
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


# NOTE:
# Makerspace password was not supplied by you.
# This is a temporary placeholder and MUST be changed before
# real deployment.
SUBJECT_PASSWORDS: dict[str, str] = {
    "Tamil": "Staff@tamil",
    "Communicative English": "Staff@english",
    "Algebra and Calculus": "Staff@maths",
    "Applied Chemistry": "Staff@chemistry",
    "Computer Programming 1": "Staff@cs",
    "CAD Based Engineering graphics": "Staff@eg",
    "Computer Programming laboratory I": "Staff@cs",
    "Makerspace": "Staff@makerspace",
    "Seminar 1": "Staff@seminar",
    "Seminar 2": "Staff@cs",
    "Seminar 3": "Staff@seminar",
}


app = FastAPI(
    title="College Attendance Backend",
    version="2.0.0",
)


# ============================================================
# DATABASE
# ============================================================

def get_connection():
    return psycopg.connect(
        DATABASE_URL,
        row_factory=psycopg.rows.dict_row,
    )


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
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
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
                one_active_attendance_session
                ON attendance_sessions (active)
                WHERE active = TRUE
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
                CREATE TABLE IF NOT EXISTS verification_codes (
                    id BIGSERIAL PRIMARY KEY,
                    user_type VARCHAR(20) NOT NULL,
                    user_id TEXT NOT NULL,
                    verification_code VARCHAR(10) NOT NULL,
                    expires_at TIMESTAMPTZ NOT NULL,
                    used BOOLEAN NOT NULL DEFAULT FALSE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
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

        conn.commit()


@app.on_event("startup")
def startup_event():
    init_database()


# ============================================================
# HELPERS
# ============================================================

def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def hash_password(password: str) -> str:
    return hashlib.sha256(
        password.encode("utf-8")
    ).hexdigest()


def sno_from_registration_last4(last4: str) -> str:
    value = last4.strip()
    if not re.fullmatch(r"\d{4}", value):
        raise HTTPException(status_code=400, detail="Enter exactly the last 4 digits of your registration number.")
    sno = REGISTRATION_LAST4_TO_SNO.get(value)
    if sno is None:
        raise HTTPException(status_code=404, detail="Registration number suffix not found.")
    return sno


def make_code() -> str:
    return f"{secrets.randbelow(1_000_000):06d}"


def distance_meters(
    lat1: float,
    lon1: float,
    lat2: float,
    lon2: float,
) -> float:
    radius = 6_371_000.0

    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)

    a = (
        math.sin(dp / 2) ** 2
        + math.cos(p1)
        * math.cos(p2)
        * math.sin(dl / 2) ** 2
    )

    return radius * 2 * math.atan2(
        math.sqrt(a),
        math.sqrt(1 - a),
    )


def close_expired_sessions() -> None:
    """
    Closes expired sessions and permanently records Absent
    for every student who never marked attendance.

    IMPORTANT:
    This does NOT put Present students in the absentee list.
    It records Absent only for students without a Present record.
    """
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
                    SET active = FALSE,
                        closed_at = %s
                    WHERE id = %s
                    """,
                    (current, session["id"]),
                )

                # Finalize only students who did NOT mark Present.
                cur.execute(
                    """
                    INSERT INTO attendance_records
                    (
                        session_id,
                        sno,
                        subject,
                        period,
                        attendance_date,
                        marked_at,
                        status
                    )
                    SELECT
                        %s,
                        s.sno,
                        %s,
                        %s,
                        %s,
                        NULL,
                        'Absent'
                    FROM students s
                    WHERE NOT EXISTS (
                        SELECT 1
                        FROM attendance_records ar
                        WHERE ar.session_id = %s
                          AND ar.sno = s.sno
                          AND ar.status = 'Present'
                    )
                    ON CONFLICT (session_id, sno)
                    DO NOTHING
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
    """Return the currently usable session without taking expiry locks.

    Expired sessions are finalized separately by the endpoints that need
    absentee records. Keeping this lookup read-only prevents repeated
    student requests from competing for the same expired-session lock.
    """
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

@app.get("/")
def root():
    return {
        "message": "College Attendance Backend is running",
        "students": len(STUDENTS),
        "subjects": len(SUBJECTS),
        "code_validity_minutes": CODE_VALIDITY_MINUTES,
    }


@app.get("/health")
def health():
    return {"status": "ok", "database": "connected"}


# ============================================================
# SUBJECTS
# ============================================================

@app.get("/subjects")
def list_subjects():
    return [
        {
            "subject": subject,
            "staff": staff,
        }
        for subject, staff in SUBJECTS.items()
    ]


@app.post("/staff/verify-subject")
def verify_subject_password(
    subject: str,
    password: str,
):
    if subject not in SUBJECTS:
        raise HTTPException(
            status_code=404,
            detail="Subject not found.",
        )

    expected = SUBJECT_PASSWORDS.get(subject)

    if expected is None:
        raise HTTPException(
            status_code=500,
            detail="Password is not configured for this subject.",
        )

    if not secrets.compare_digest(password, expected):
        raise HTTPException(
            status_code=401,
            detail="Incorrect subject password.",
        )

    return {
        "success": True,
        "subject": subject,
        "staff": SUBJECTS[subject],
    }


# ============================================================
# START ATTENDANCE
# ============================================================

@app.post("/attendance/start")
def start_attendance(
    subject: str = Query(...),
    period: Optional[str] = Query(None),
):
    if subject not in SUBJECTS:
        raise HTTPException(
            status_code=404,
            detail="Subject not found.",
        )

    # Finalize an expired session before starting another one.
    # This is intentionally done before the transaction below so an old
    # 1-minute session can never block the next Generate Code operation.
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
                    (
                        subject,
                        current.date(),
                    ),
                )
                number = int(cur.fetchone()["count"]) + 1

        period = f"Session {number}"
    else:
        period = period.strip()

    with get_connection() as conn:
        with conn.cursor() as cur:
            # Serialize concurrent Generate Code requests.
            # PostgreSQL releases this lock automatically when the transaction ends.
            cur.execute("SELECT pg_advisory_xact_lock(%s)", (918273645,))

            # Sessions are independent by subject. A Maths session must not
            # prevent a Tamil session from being started, and vice versa.
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
                (
                    subject,
                    staff_name,
                    period,
                    started_at,
                    expires_at,
                    active
                )
                VALUES (%s, %s, %s, %s, %s, TRUE)
                RETURNING id
                """,
                (
                    subject,
                    SUBJECTS[subject],
                    period,
                    current,
                    expires,
                ),
            )

            session_id = cur.fetchone()["id"]

            used_codes: set[str] = set()

            for sno in sorted(
                STUDENTS,
                key=lambda value: int(value),
            ):
                code = make_code()

                while code in used_codes:
                    code = make_code()

                used_codes.add(code)

                cur.execute(
                    """
                    INSERT INTO attendance_codes
                    (
                        session_id,
                        sno,
                        code
                    )
                    VALUES (%s, %s, %s)
                    """,
                    (
                        session_id,
                        sno,
                        code,
                    ),
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
def get_my_code(sno: str):
    if sno not in STUDENTS:
        raise HTTPException(
            status_code=404,
            detail="Student S.No. not found.",
        )

    session = get_active_session()

    if not session:
        raise HTTPException(
            status_code=404,
            detail="No active attendance session.",
        )

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT code
                FROM attendance_codes
                WHERE session_id = %s
                  AND sno = %s
                """,
                (
                    session["id"],
                    sno,
                ),
            )

            result = cur.fetchone()

    if not result:
        raise HTTPException(
            status_code=404,
            detail="Attendance code not found.",
        )

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
    latitude: Optional[float] = Field(default=None)
    longitude: Optional[float] = Field(default=None)


@app.post("/attendance/mark")
def mark_attendance(request: MarkAttendanceRequest):
    sno = request.sno.strip()
    code = request.code.strip()

    if sno not in STUDENTS:
        raise HTTPException(
            status_code=404,
            detail="Student S.No. not found.",
        )

    session = get_active_session()

    if not session:
        raise HTTPException(
            status_code=410,
            detail="Attendance session has ended.",
        )

    with get_connection() as conn:
        with conn.cursor() as cur:
            # Verify that the code belongs specifically to this student
            # and specifically to the current session.
            cur.execute(
                """
                SELECT id
                FROM attendance_codes
                WHERE session_id = %s
                  AND sno = %s
                  AND code = %s
                """,
                (
                    session["id"],
                    sno,
                    code,
                ),
            )

            code_record = cur.fetchone()

            if not code_record:
                raise HTTPException(
                    status_code=401,
                    detail="Invalid attendance code.",
                )

            # Duplicate prevention.
            cur.execute(
                """
                SELECT id
                FROM attendance_records
                WHERE session_id = %s
                  AND sno = %s
                  AND status = 'Present'
                """,
                (
                    session["id"],
                    sno,
                ),
            )

            if cur.fetchone():
                raise HTTPException(
                    status_code=409,
                    detail="Attendance is already marked.",
                )

            # Location verification.
            distance = None

            if LOCATION_REQUIRED:
                if (
                    request.latitude is None
                    or request.longitude is None
                ):
                    raise HTTPException(
                        status_code=400,
                        detail="Location is required.",
                    )

                if (
                    CLASSROOM_LATITUDE == 0.0
                    and CLASSROOM_LONGITUDE == 0.0
                ):
                    raise HTTPException(
                        status_code=500,
                        detail="Classroom coordinates are not configured.",
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
                    session_id,
                    sno,
                    subject,
                    period,
                    attendance_date,
                    marked_at,
                    status,
                    latitude,
                    longitude,
                    distance_meters
                )
                VALUES
                (
                    %s, %s, %s, %s, %s,
                    %s, 'Present', %s, %s, %s
                )
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
def get_absentees(
    subject: str = Query(...),
):
    if subject not in SUBJECTS:
        raise HTTPException(
            status_code=404,
            detail="Subject not found.",
        )

    # If the 1-minute session has ended, finalize absent records
    # before calculating the list.
    close_expired_sessions()

    active = get_active_session()

    if active and active["subject"] == subject:
        session = active
    else:
        session = get_latest_session_for_subject(subject)

    if not session:
        return {
            "subject": subject,
            "period": None,
            "date": None,
            "session_id": None,
            "active": False,
            "absentees": [],
        }

    # THIS IS THE IMPORTANT PART:
    #
    # Start with the COMPLETE class list.
    # Remove everyone who has PRESENT attendance in this session.
    #
    # Therefore the response contains ONLY students who did not
    # mark attendance.
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    s.sno,
                    s.name
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
# ABSENTEE HISTORY
# ============================================================

@app.get("/attendance/absentee-history")
def get_absentee_history(subject: str = Query(...)):
    """Return the latest subject's absentee list for 45 minutes after expiry.

    Attendance records remain stored permanently, but the staff UI exposes
    the absentee list only for ABSENTEE_VISIBILITY_MINUTES after that
    specific subject session expires. A later session for the same subject
    replaces the earlier visible list.
    """
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Subject not found.")

    close_expired_sessions()
    current = now_utc()

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

            # Do not expose an absentee list while a new session is active.
            if session["active"] or session["expires_at"] > current:
                return {
                    "subject": subject,
                    "date": None,
                    "session_id": session["id"],
                    "visible_until": None,
                    "absentees": [],
                }

            visible_until = session["expires_at"] + timedelta(
                minutes=ABSENTEE_VISIBILITY_MINUTES
            )

            if current >= visible_until:
                return {
                    "subject": subject,
                    "date": None,
                    "session_id": session["id"],
                    "visible_until": visible_until.isoformat(),
                    "absentees": [],
                }

            # Calculate absentees directly from the class roster minus
            # students who were marked Present for this exact session.
            # This makes the list reliable even if the finalization request
            # and the UI expiry timer happen at slightly different moments.
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
        "visible_until": visible_until.isoformat(),
        "absentees": rows,
    }


# ============================================================
# CLOSE ATTENDANCE MANUALLY
# ============================================================

@app.post("/attendance/close")
def close_attendance(
    subject: str = Query(...),
):
    close_expired_sessions()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM attendance_sessions
                WHERE subject = %s
                  AND active = TRUE
                ORDER BY started_at DESC
                LIMIT 1
                FOR UPDATE
                """,
                (subject,),
            )

            session = cur.fetchone()

            if not session:
                raise HTTPException(
                    status_code=404,
                    detail="No active session found.",
                )

            current = now_utc()

            cur.execute(
                """
                UPDATE attendance_sessions
                SET active = FALSE,
                    closed_at = %s
                WHERE id = %s
                """,
                (
                    current,
                    session["id"],
                ),
            )

            # Immediately finalize everyone who never marked Present.
            cur.execute(
                """
                INSERT INTO attendance_records
                (
                    session_id,
                    sno,
                    subject,
                    period,
                    attendance_date,
                    marked_at,
                    status
                )
                SELECT
                    %s,
                    s.sno,
                    %s,
                    %s,
                    %s,
                    NULL,
                    'Absent'
                FROM students s
                WHERE NOT EXISTS (
                    SELECT 1
                    FROM attendance_records ar
                    WHERE ar.session_id = %s
                      AND ar.sno = s.sno
                      AND ar.status = 'Present'
                )
                ON CONFLICT (session_id, sno)
                DO NOTHING
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
        raise HTTPException(
            status_code=404,
            detail="Student not found.",
        )

    close_expired_sessions()

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    sno,
                    subject,
                    period,
                    attendance_date,
                    marked_at,
                    status
                FROM attendance_records
                WHERE sno = %s
                ORDER BY
                    attendance_date DESC,
                    session_id DESC
                """,
                (sno,),
            )

            records = cur.fetchall()

    for record in records:
        record["attendance_date"] = (
            record["attendance_date"].isoformat()
        )

        if record["marked_at"] is not None:
            record["marked_at"] = (
                record["marked_at"].isoformat()
            )

    return {
        "sno": sno,
        "name": STUDENTS[sno],
        "records": records,
    }


# ============================================================
# STUDENT PASSWORD
# ============================================================

class StudentPasswordRequest(BaseModel):
    registration_last4: str
    password: str


@app.post("/student/set-password")
def set_student_password(request: StudentPasswordRequest):
    sno = sno_from_registration_last4(request.registration_last4)
    if len(request.password) < 4:
        raise HTTPException(status_code=400, detail="Password must contain at least 4 characters.")

    password_hash = hash_password(request.password)

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT sno FROM student_passwords WHERE sno = %s", (sno,))
            if cur.fetchone():
                raise HTTPException(status_code=409, detail="Password already created for this student.")
            cur.execute(
                "INSERT INTO student_passwords (sno, password_hash) VALUES (%s, %s)",
                (sno, password_hash),
            )
        conn.commit()

    return {"success": True, "sno": sno, "name": STUDENTS[sno]}


@app.post("/student/login")
def student_login(request: StudentPasswordRequest):
    sno = sno_from_registration_last4(request.registration_last4)

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT password_hash FROM student_passwords WHERE sno = %s", (sno,))
            record = cur.fetchone()

    if not record:
        raise HTTPException(status_code=404, detail="Password not created yet.")

    if not secrets.compare_digest(hash_password(request.password), record["password_hash"]):
        raise HTTPException(status_code=401, detail="Incorrect password.")

    return {"success": True, "sno": sno, "name": STUDENTS[sno]}


# ============================================================
# EXCEL EXPORT
# ============================================================

@app.get("/attendance/export")
def export_attendance(
    subject: Optional[str] = None,
    date: Optional[str] = None,
):
    close_expired_sessions()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Attendance"

    sheet.append(
        [
            "S.No.",
            "Name",
            "Date",
            "Subject",
            "Period",
            "Status",
            "Marked At",
        ]
    )

    with get_connection() as conn:
        with conn.cursor() as cur:
            query = """
                SELECT
                    ar.sno,
                    s.name,
                    ar.attendance_date,
                    ar.subject,
                    ar.period,
                    ar.status,
                    ar.marked_at
                FROM attendance_records ar
                JOIN students s
                  ON s.sno = ar.sno
                WHERE 1 = 1
            """

            params: list[str] = []

            if subject:
                query += " AND ar.subject = %s"
                params.append(subject)

            if date:
                query += " AND ar.attendance_date = %s"
                params.append(date)

            query += """
                ORDER BY
                    ar.sno::INTEGER ASC,
                    ar.attendance_date ASC,
                    ar.session_id ASC
            """

            cur.execute(query, params)
            rows = cur.fetchall()

    for row in rows:
        sheet.append(
            [
                int(row["sno"]),
                row["name"],
                row["attendance_date"].isoformat(),
                row["subject"],
                row["period"],
                row["status"],
                (
                    row["marked_at"].isoformat()
                    if row["marked_at"]
                    else ""
                ),
            ]
        )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition":
                'attachment; filename="attendance.xlsx"'
        },
    )


# ============================================================
# DEVELOPMENT-ONLY RESET
# ============================================================

@app.delete("/development/reset-database")
def development_reset_database():
    """
    Development helper only.
    DO NOT expose this endpoint in a real deployment.
    """
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "TRUNCATE attendance_records, "
                "attendance_codes, attendance_sessions, "
                "student_passwords, verification_codes "
                "RESTART IDENTITY CASCADE"
            )

        conn.commit()

    return {
        "success": True,
        "message": "Development attendance data reset.",
    }
