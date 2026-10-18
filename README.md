# College Attendance — Backend

FastAPI + PostgreSQL service for the college attendance app.

## What changed in v3

- **Passwords are hashed with PBKDF2-HMAC-SHA256** (salted). Legacy unsalted
  SHA-256 hashes are still accepted and are transparently upgraded on the next
  successful login.
- **One unique password per staff member** (`staff_accounts`), instead of
  subject passwords shared by several people.
- **Attendance data is never hidden after a time window.** A session's records
  stay for the whole semester. There is a new semester report.
- **Email password reset** with a 6-digit code (`student_passwords.email`).
- **Per-student semester summary** (`Present`, `Absent`, `Total`, `Percentage`).
- **Pluggable ERP delivery**: generate an Excel file and email it, or (later)
  push it to an ERP endpoint.
- Rate limiting on login and reset endpoints; the destructive reset endpoint is
  disabled unless `RESET_DB_ENABLED=true`.
- Optional connection pooling via `psycopg_pool`; a per-subject active session
  lookup; and a wider (120 m) classroom radius for browser GPS.

## Endpoints

| Method | Path | Purpose |
| --- | --- | --- |
| GET | `/health` | Service + database check |
| GET | `/subjects` | Subject and staff list |
| POST | `/staff/verify-subject` | Staff login (subject + password) |
| POST | `/staff/change-password` | Staff changes own password |
| POST | `/attendance/start` | Start a timed session for a subject |
| GET | `/attendance/my-code/{sno}` | Student's own code |
| POST | `/attendance/mark` | Mark present (GPS checked) |
| GET | `/attendance/absentees` | Live absentee list for a subject |
| GET | `/attendance/absentee-history` | Latest session's absentee list |
| GET | `/attendance/sessions` | All sessions for a subject |
| GET | `/attendance/summary` | Per-student semester summary |
| POST | `/attendance/close` | End a session now |
| GET | `/attendance/history/{sno}` | One student's records |
| POST | `/student/set-password` | First-time student setup (+ email) |
| POST | `/student/login` | Student login |
| POST | `/student/request-password-reset` | Email a 6-digit reset code |
| POST | `/student/reset-password` | Set a new password with the code |
| GET | `/attendance/export` | Download the summary as `.xlsx` |
| POST | `/attendance/erp-update` | Generate + deliver the summary to the ERP |

## Configuration

Copy `.env.example` and fill in the values, or set them in the Catalyst AppSail
environment. Everything has a safe default except `DATABASE_URL`.

| Variable | Purpose |
| --- | --- |
| `DATABASE_URL` | Postgres connection string (Neon). Required. |
| `CODE_VALIDITY_MINUTES` | How long a session stays open (default 15) |
| `MAX_ATTENDANCE_DISTANCE_METERS` | Classroom radius (default 120) |
| `LOCATION_REQUIRED` | Require GPS coordinates (default true) |
| `CLASSROOM_LATITUDE` / `CLASSROOM_LONGITUDE` | Classroom location |
| `SEMESTER_START` / `SEMESTER_END` | Reporting window (informational) |
| `ALLOWED_ORIGINS` | Comma-separated web origins (default `*`) |
| `SMTP_HOST` / `SMTP_PORT` / `SMTP_USER` / `SMTP_PASSWORD` / `SMTP_FROM` | Outbound mail (Catalyst Mail) |
| `ERP_MODE` | `file` (email the Excel) or `api` (POST to an endpoint) |
| `ERP_UPLOAD_EMAIL` | Where the ERP summary is emailed |
| `ERP_API_URL` / `ERP_API_KEY` | Used when `ERP_MODE=api` |
| `RESET_DB_ENABLED` | Enables `DELETE /development/reset-database` |

If `SMTP_HOST` is not set, emails are written to the log instead of sent, and
`request-password-reset` still returns success (so the flow stays testable).

## Staff passwords

On first startup the backend creates one account per staff member with a
random password and logs it:

```
WARNING:attendance:[staff-init] Ms. S. Saranya initial password: Svce@67fa3e (must be changed on first login)
```

Collect those from the AppSail logs and hand them out. Each staff member is
asked to set their own password the first time they sign in.

## Local development

```bash
python -m venv venv
venv\Scripts\pip install -r requirements.txt
# set DATABASE_URL in .env, then
venv\Scripts\uvicorn main:app --reload --port 8000
```

## Deploy to Zoho Catalyst (AppSail + Docker)

1. Create a Neon Postgres database and copy its connection string.
2. Create a Catalyst AppSail service from this repository; it builds `Dockerfile`.
3. Set the environment variables above (at minimum `DATABASE_URL`).
4. AppSail passes the port through `X_ZOHO_CATALYST_LISTEN_PORT`; the
   `Dockerfile` already binds to it on `0.0.0.0`.
5. After the first boot, read the staff passwords from the logs.
6. Health check: `GET /health` should return `{"status":"ok"}`.

## ERP integration

`ERP_MODE=file` (the default) builds a one-sheet Excel summary — S.No, Name,
Subject, Present, Absent, Total, Percentage — and emails it to
`ERP_UPLOAD_EMAIL`. Staff can also download the same file from the app.

When the college ERP interface is known, set `ERP_MODE=api`, `ERP_API_URL` and
`ERP_API_KEY`; no app change is needed. Every attempt is recorded in
`erp_sync_log`.
