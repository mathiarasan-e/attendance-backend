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
- **No student self-service reset.** Students remember their own password; if one
  is forgotten the college clears it and the student sets a new one
  (`clear_student_password.py`).
- **Per-student semester summary** (`Present`, `Absent`). The counts live in the
  app, on the subject's report page; there is no percentage column anywhere.
- **Month-wise absentee database.** For a subject, staff can browse any month,
  see the absentees for every session and per-student totals, and download that
  month as an `.xlsx` workbook.
- **Pluggable ERP delivery**: generate an Excel file and email it, or (later)
  push it to an ERP endpoint. The app's "Update to ERP" button is hidden until
  the college interface is agreed.
- Rate limiting on login endpoints; the destructive reset endpoint is
  disabled unless `RESET_DB_ENABLED=true`.
- Optional connection pooling via `psycopg_pool`; a per-subject active session
  lookup; and a campus-wide attendance circle (200 m) for browser GPS.

## Endpoints

| Method | Path | Purpose |
| --- | --- | --- |
| GET | `/` | The Flutter web app (falls back to the info JSON locally) |
| GET | `/api/info` | Service summary: students, subjects, radius, ERP mode |
| GET | `/health` | Service + database check |
| GET | `/subjects` | Subject and staff list |
| POST | `/staff/verify-subject` | Staff login (subject + password) |
| POST | `/staff/change-password` | Staff changes own password |
| POST | `/attendance/start` | Start a timed session for a subject |
| GET | `/attendance/my-code/{sno}` | Student's own code |
| POST | `/attendance/mark` | Mark present (campus GPS checked) |
| GET | `/attendance/absentees` | Live absentee list for a subject |
| GET | `/attendance/absentee-history` | Latest session's absentee list |
| GET | `/attendance/sessions` | All sessions for a subject |
| GET | `/attendance/summary` | Per-student semester summary (present/absent counts) |
| GET | `/attendance/monthly-absentees` | Month-wise absentees (sessions + per-student totals) |
| GET | `/attendance/monthly-absentees/export` | Download one month of absentees as `.xlsx` |
| POST | `/attendance/close` | End a session now |
| GET | `/attendance/history/{sno}` | One student's records |
| POST | `/student/set-password` | First-time student setup |
| POST | `/student/login` | Student login |
| GET | `/attendance/export` | Download the summary as `.xlsx` (S.No, Name, Present, Absent) |
| POST | `/attendance/erp-update` | Generate + deliver the summary to the ERP |

## Serving the web app

AppSail serves the compiled Flutter app **and** the API from one origin: `GET /`
returns the UI, and every API path above keeps priority over it. The browser
therefore never makes a cross-origin request.

That matters here. The Catalyst gateway answers CORS preflights itself and does
not pass them through to this app, so it returns `200` with no
`Access-Control-Allow-*` headers. A browser would reject that and block every
JSON `POST` — login, start session, mark attendance, ERP update — from a
different origin. Same-origin serving sidesteps it entirely.

`build_appsail.ps1` copies the Flutter build into `build/web`. `WEB_DIR`
overrides the location, which is useful for local runs. When no build is present
the API still works and `GET /` returns the summary JSON instead.

## Configuration

Copy `.env.example` and fill in the values, or set them in the Catalyst AppSail
environment. Everything has a safe default except `DATABASE_URL`.

| Variable | Purpose |
| --- | --- |
| `DATABASE_URL` | Postgres connection string (Neon). Required. |
| `CODE_VALIDITY_MINUTES` | How long a session stays open (default 15) |
| `CAMPUS_CENTER_LATITUDE` / `CAMPUS_CENTER_LONGITUDE` | Centre of the campus attendance circle (default 12.987929, 79.972311) |
| `CAMPUS_RADIUS_METERS` | Campus radius (default 200) |
| `CAMPUS_ACCURACY_TOLERANCE_METERS` | Extra metres allowed outside the boundary for a noisy fix (default 0) |
| `MAX_GPS_ACCURACY_METERS` | Reject fixes worse than this (default 100) |
| `LOCATION_REQUIRED` | Require GPS coordinates (default true) |
| `SEMESTER_START` / `SEMESTER_END` | Reporting window (informational) |
| `ALLOWED_ORIGINS` | Comma-separated web origins (default `*`; unused when the UI is served from this origin) |
| `WEB_DIR` | Directory of the compiled Flutter app (default `./web`) |
| `SMTP_HOST` / `SMTP_PORT` / `SMTP_USER` / `SMTP_PASSWORD` / `SMTP_FROM` | Outbound mail (Catalyst Mail) |
| `ERP_MODE` | `file` (email the Excel) or `api` (POST to an endpoint) |
| `ERP_UPLOAD_EMAIL` | Where the ERP summary is emailed |
| `ERP_API_URL` / `ERP_API_KEY` | Used when `ERP_MODE=api` |
| `RESET_DB_ENABLED` | Enables `DELETE /development/reset-database` |

`SMTP_HOST` and `SMTP_FROM` are only needed by `ERP_MODE=file`, which emails the
semester summary. Nothing else sends mail. If they are unset the message is
written to the log instead of sent.

## Students and login

The class list lives in two places that **must be kept identical**:

| File | What |
| --- | --- |
| `main.py` | `STUDENTS` (s.no → name) and `ADMISSION_LAST4_TO_SNO` |
| `lib/main.dart` | the `students` map |

The API validates the roster from `main.py`; the app renders the welcome name and
s.no from its own copy, so a mismatch shows the wrong student a wrong name.

Students sign in with the **last four digits of their admission number**
(`2026CSxxxx`). `ADMISSION_LAST4_TO_SNO` maps those digits to the s.no, so
`2026CS0398` signs in as s.no `124`, which the app then displays.

Every startup syncs the roster: anything added to `STUDENTS` is inserted, and
anything no longer listed is deleted so it stops showing in staff lists. A
student who already has attendance records is never deleted — `attendance_records`
cascades on delete, so a stale name is preferable to silently losing history —
and a warning is logged instead.

Because every attendance table is keyed by `s.no`, changing the roster renumbers
the data with it: a row written for s.no `121` now describes whoever holds `121`.
So when a student leaves, or the class is renumbered, clear the attendance data
and let the sync finish the job:

```powershell
venv\Scripts\python clear_attendance_data.py                              # asks to confirm
venv\Scripts\python clear_attendance_data.py --yes
venv\Scripts\python clear_attendance_data.py --yes --student-passwords
```

It deletes the attendance tables first (that is what unblocks the delete above),
then calls the same `init_database()` the service uses, so a student who has left
really is removed. It prints the target host and the row counts before touching
anything. Pass `--student-passwords` to also delete every student password, so
all students create their own again.

## Student passwords

Students create their own password the first time they sign in, and are told
plainly to remember it: there is no reset email and no reset code. When one is
forgotten, clear it and the student sets a new one from the same page:

```powershell
venv\Scripts\python clear_student_password.py 0398     # by admission last 4 digits
venv\Scripts\python clear_student_password.py 124      # by s.no
venv\Scripts\python clear_student_password.py --all --yes
```

## Staff passwords

Staff choose their own password. The backend creates one account per staff
member **with no password at all**, so nothing secret is ever written to the
logs — the only staff line you will see is:

```
[staff-setup] Ms. S. Saranya chose their own password
```

A staff member opens the portal, picks *First time? Create your password*,
chooses their subject, and enters the college's setup code plus a password of
their own. That code is `STAFF_SETUP_CODE`; set it in the AppSail environment
variables (and in `.env` for local runs) and hand it out to staff. While it is
unset, self-service is switched off and the page answers **503**.

The code only works while an account has no password, so it cannot be used to
take over an account that is already in use. Signing in to an account that has no
password yet returns **409**, which is how the app knows to offer the setup page.

If a staff member forgets their password, clear it and they can set a new one
with the same code:

```powershell
venv\Scripts\python clear_staff_passwords.py            # prints the host, asks to confirm
venv\Scripts\python clear_staff_passwords.py --yes
venv\Scripts\python clear_staff_passwords.py --staff "Ms. S. Saranya"
```

It resolves the database the same way the app does (`.env`, then `.env.local`,
then real environment variables). Check the host it prints before confirming.

## Local development

`.env` is loaded automatically at startup, then `.env.local` is loaded on top of
it and takes precedence. Real environment variables override both, so a
deployed container's configuration is never overridden by a stray file.

Keep the shared configuration in `.env` and machine-specific overrides (for
example a Neon **dev** branch) in `.env.local`, which is git-ignored. Pointing
`.env` at production means local testing writes real attendance data.

```powershell
python -m venv venv
venv\Scripts\pip install -r requirements.txt

# .env.local, for example:
#   DATABASE_URL=<pooled dev branch connection string>
# Optional: serve the real UI. Relative paths resolve from this directory.
$env:WEB_DIR = "..\attendance_app\build\web"

venv\Scripts\python main.py
```

`python main.py` binds `0.0.0.0` on `X_ZOHO_CATALYST_LISTEN_PORT` (default
`9000`), the same path AppSail uses, so it is the closest thing to production.
For autoreload use `venv\Scripts\uvicorn main:app --reload --port 8000`.

## Database (Neon)

`DATABASE_URL` takes a Neon connection string. Use the **pooled** one (the host
contains `-pooler`) for application traffic, because the app keeps a small
connection pool; use the direct string for migrations and `pg_dump`.
`neon link` writes both into `.env` as `DATABASE_URL` and
`DATABASE_URL_UNPOOLED`.

## Deploy to Zoho Catalyst AppSail (managed Python runtime)

AppSail does **not** run `pip install` for managed runtimes, and it runs Linux,
so every dependency must be vendored into the build folder as a Linux wheel.
`build_appsail.ps1` does that, pinned to the Python 3.12 wheels that match the
`python_3_12` stack:

```powershell
# 1. Compile the web UI (run in the attendance_app repo). The API base URL must
#    match the AppSail origin the UI will be served from.
flutter build web --release --dart-define=API_BASE_URL=https://<appsail-url>

# 2. Vendor the Linux dependencies and bundle the UI
.\build_appsail.ps1     # produces build/ = main.py + Linux wheels + web/
```

Without step 1 the script still works and deploys the API alone.

Then register and deploy the service:

```bash
catalyst appsail:add -ni --name attendance-api --stack python_3_12 \
  --source "<repo dir>" --build build --command "python3 main.py"
catalyst deploy -ni --only appsail
```

`main.py` has a `__main__` block that binds to `0.0.0.0` on
`X_ZOHO_CATALYST_LISTEN_PORT`, so the startup command stays a plain command
with no shell quoting or variable interpolation to get wrong.

Set the environment variables above (at minimum `DATABASE_URL`) in the AppSail
service configuration. After the first boot, collect the staff passwords from
the logs, then confirm `GET /health` returns
`{"status":"ok","database":"connected"}`.

`app-config.json` carries the service's environment variables, so it is
git-ignored: it holds the database password. `catalyst appsail:add`
regenerates it. The same variables can be managed in the Catalyst console
instead, which is the better home for secrets on a shared machine.

`Dockerfile` is kept for a possible future container-based deployment; the
managed runtime path above does not use it.


## ERP integration

`ERP_MODE=file` (the default) builds a one-sheet Excel summary — S.No, Name,
Present, Absent — and emails it to `ERP_UPLOAD_EMAIL`. Staff can also download
the same file from the app, from the subject's report page.

The report page is where the semester record is kept, per subject. Its Download
button only unlocks once at least one session has finished and none is still
running, so a sheet cannot be pulled halfway through a class; while it locked
the page re-checks every 30 seconds and unlocks itself.

When the college ERP interface is known, set `ERP_MODE=api`, `ERP_API_URL` and
`ERP_API_KEY`; no app change is needed. Every attempt is recorded in
`erp_sync_log`.
