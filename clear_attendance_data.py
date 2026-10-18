"""Clear attendance data, and remove students who are no longer in the roster.

Two things get in the way when the class roster is renumbered or someone leaves:

* ``init_database()`` deliberately refuses to delete a student who has
  attendance records, so it only logs a warning and leaves the row behind.
* Every attendance table is keyed by ``sno``. If s.no values shift, old rows
  silently describe a *different* student.

This tool deletes the attendance data first, then calls the app's own
``init_database()`` so the normal roster sync can finish the job. History is
gone either way, which is the point: it is meant for clearing test runs and for
applying a renumbering, not for routine use.

Usage, from the attendance_backend directory:

    venv\\Scripts\\python clear_attendance_data.py                        # asks to confirm
    venv\\Scripts\\python clear_attendance_data.py --yes                  # no confirmation
    venv\\Scripts\\python clear_attendance_data.py --yes --student-passwords

Which database is targeted follows the normal rules: ``.env`` is loaded first,
then ``.env.local`` on top of it, then real environment variables. The script
prints the target host before doing anything, so you can confirm you are not
about to touch production by accident.
"""

from __future__ import annotations

import argparse
import logging
import sys

# Importing main turns on app-level INFO logging. Silence it so this command's
# output is only the confirmation prompt and the result. Warnings still show,
# which is how a roster-sync problem would surface.
logging.disable(logging.INFO)

import main  # noqa: E402 - must follow the logging tweak above

# Deleted in this order: rows in each table point at the ones before it.
ATTENDANCE_TABLES = (
    "attendance_codes",
    "attendance_device_locks",
    "attendance_records",
    "attendance_sessions",
)


def clear(args: argparse.Namespace) -> int:
    with main.get_connection() as conn:
        with conn.cursor() as cur:
            counts = {}
            for table in ATTENDANCE_TABLES:
                cur.execute(f"SELECT COUNT(*) AS n FROM {table}")
                counts[table] = cur.fetchone()["n"]

            cur.execute("SELECT sno FROM students ORDER BY sno::int")
            roster = [row["sno"] for row in cur.fetchall()]
            stale = [sno for sno in roster if sno not in main.STUDENTS]

            passwords = 0
            if args.student_passwords:
                cur.execute("SELECT COUNT(*) AS n FROM student_passwords")
                passwords = cur.fetchone()["n"]

    print()
    print("About to delete:")
    for table in ATTENDANCE_TABLES:
        print(f"  {counts[table]:>6}  {table}")
    if stale:
        print(f"  {len(stale):>6}  students no longer in the roster: {', '.join(stale)}")
    if args.student_passwords:
        print(f"  {passwords:>6}  student passwords (and pending reset codes)")

    if not args.yes:
        print()
        answer = input("Type 'yes' to continue: ").strip().lower()
        if answer != "yes":
            print("Aborted. Nothing changed.")
            return 1

    with main.get_connection() as conn:
        with conn.cursor() as cur:
            for table in ATTENDANCE_TABLES:
                cur.execute(f"DELETE FROM {table}")
            if args.student_passwords:
                cur.execute("DELETE FROM student_passwords")
                cur.execute("DELETE FROM verification_codes")
        conn.commit()

    print()
    print("Attendance data cleared.")

    # Now that nothing references those students, the app's own roster sync can
    # delete them. This is the same code path the service runs at startup.
    main.init_database()

    with main.get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT sno FROM students ORDER BY sno::int")
            remaining = [row["sno"] for row in cur.fetchall()]
            left_over = [sno for sno in remaining if sno not in main.STUDENTS]

    if remaining:
        print(
            f"Roster now holds {len(remaining)} student(s), "
            f"s.no {remaining[0]}-{remaining[-1]}."
        )
    else:
        print("Roster is empty.")

    if left_over:
        print(
            f"Still present but not in the roster: {', '.join(left_over)}",
            file=sys.stderr,
        )
        return 1

    if args.student_passwords:
        print("Student passwords cleared: every student creates their own again.")
    return 0


def run() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--yes", action="store_true", help="skip the confirmation prompt")
    parser.add_argument(
        "--student-passwords",
        action="store_true",
        help=(
            "also delete every student password and pending reset code, so all "
            "students create their own again"
        ),
    )
    args = parser.parse_args()

    if not main.DATABASE_URL:
        print("DATABASE_URL is not set. Nothing to do.", file=sys.stderr)
        return 1

    host = main.DATABASE_URL.rsplit("@", 1)[-1].split("/")[0]
    print(f"Target database host : {host}")

    try:
        return clear(args)
    finally:
        # The connection pool starts worker threads; close it so the script
        # exits promptly instead of warning about a lingering thread.
        if main._pool is not None:
            main._pool.close()
            main._pool = None


if __name__ == "__main__":
    raise SystemExit(run())
