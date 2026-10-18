"""Clear a student password so the student can create a new one.

Students remember their own password: there is no reset email and no reset code.
When one is forgotten, the college clears it here and the student sets a new one
from the portal using the same *First time? Create your password* page they used
the first time.

Usage, from the attendance_backend directory:

    venv\\Scripts\\python clear_student_password.py 0398        # by admission last 4 digits
    venv\\Scripts\\python clear_student_password.py 124         # by s.no
    venv\\Scripts\\python clear_student_password.py --yes 0398
    venv\\Scripts\\python clear_student_password.py --all --yes

A four-digit argument is an admission-number suffix (``2026CS0398`` -> ``0398``)
and a three-digit one is the s.no, so you can use whichever is on hand. Use
``--all`` to clear every student password, for example before the first real use.

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
# output is only the confirmation prompt and the result.
logging.disable(logging.INFO)

import main  # noqa: E402 - must follow the logging tweak above


def resolve_sno(arg: str) -> str:
    """Accept either an admission-number suffix (4 digits) or an s.no (3 digits)."""
    value = arg.strip()
    if not value.isdigit():
        raise ValueError(f"{arg!r} is not a number.")
    if len(value) == 4:
        if value not in main.ADMISSION_LAST4_TO_SNO:
            raise ValueError(f"admission number ending {value} is not on the roster.")
        return main.ADMISSION_LAST4_TO_SNO[value]
    if value in main.STUDENTS:
        return value
    raise ValueError(f"s.no {value} is not on the roster (123-185).")


def clear(args: argparse.Namespace) -> int:
    if args.all and args.student:
        print("Pass either a student or --all, not both.", file=sys.stderr)
        return 1
    if not args.all and not args.student:
        print(
            "Name a student (admission last 4 digits or s.no) or pass --all.\n"
            "Run with --help for examples.",
            file=sys.stderr,
        )
        return 1

    if args.all:
        targets = sorted(main.STUDENTS, key=int)
    else:
        try:
            targets = [resolve_sno(args.student)]
        except ValueError as exc:
            print(exc, file=sys.stderr)
            return 1

    with main.get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT sno FROM student_passwords ORDER BY sno::int")
            have_password = {row["sno"] for row in cur.fetchall()}

    clearable = [sno for sno in targets if sno in have_password]
    for sno in targets:
        state = "has a password" if sno in have_password else "no password set"
        print(f"  {sno:>4}  {main.STUDENTS[sno]:<30} {state}")

    if not clearable:
        print()
        print("Nothing to clear.")
        return 0

    if not args.yes:
        print()
        answer = input(f"Type 'yes' to clear {len(clearable)} student password(s): ")
        if answer.strip().lower() != "yes":
            print("Aborted. Nothing changed.")
            return 1

    with main.get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM student_passwords WHERE sno = ANY(%s)",
                (clearable,),
            )
        conn.commit()

    print()
    print(f"Cleared {len(clearable)} student password(s).")
    print()
    print("Each of them signs in, chooses 'Create your password', and picks a new one.")
    return 0


def run() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "student",
        nargs="?",
        help="admission number last 4 digits (0398) or s.no (124)",
    )
    parser.add_argument(
        "--all", action="store_true", help="clear every student password"
    )
    parser.add_argument("--yes", action="store_true", help="skip the confirmation prompt")
    args = parser.parse_args()

    if not main.DATABASE_URL:
        print("DATABASE_URL is not set. Nothing to do.", file=sys.stderr)
        return 1

    host = main.DATABASE_URL.rsplit("@", 1)[-1].split("/")[0]
    print(f"Target database host : {host}")
    print()

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
