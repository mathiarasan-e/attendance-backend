"""Reset every staff password to a fresh random value and print it once.

Use this when the one-time ``[staff-init]`` lines from the first AppSail boot are
no longer in the console logs. Behaviour matches the original seeding flow: each
staff member gets a unique password and is still forced to choose their own on
the next login.

Usage, from the attendance_backend directory:

    venv\\Scripts\\python reset_staff_passwords.py            # asks to confirm
    venv\\Scripts\\python reset_staff_passwords.py --yes      # no confirmation
    venv\\Scripts\\python reset_staff_passwords.py --staff "Ms. S. Saranya"

Which database is targeted follows the normal rules: ``.env`` is loaded first,
then ``.env.local`` on top of it, then real environment variables. The script
prints the target host before doing anything, so you can confirm you are not
about to rewrite production passwords by accident.
"""

from __future__ import annotations

import argparse
import logging
import secrets
import sys

# Importing main turns on app-level INFO logging. Silence it so this command's
# output is only the confirmation prompt and the passwords.
logging.disable(logging.INFO)

import main  # noqa: E402 - must follow the logging tweak above


def run() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--yes", action="store_true", help="skip the confirmation prompt")
    parser.add_argument(
        "--staff",
        help="reset only this staff member (must match staff_accounts.staff_name exactly)",
    )
    args = parser.parse_args()

    if not main.DATABASE_URL:
        print("DATABASE_URL is not set. Nothing to do.", file=sys.stderr)
        return 1

    host = main.DATABASE_URL.rsplit("@", 1)[-1].split("/")[0]
    print(f"Target database host : {host}")

    if not args.yes:
        answer = input("Type 'yes' to reset staff password(s): ").strip().lower()
        if answer != "yes":
            print("Aborted. Nothing changed.")
            return 1

    try:
        with main.get_connection() as conn:
            with conn.cursor() as cur:
                if args.staff:
                    cur.execute(
                        "SELECT staff_name FROM staff_accounts WHERE staff_name = %s",
                        (args.staff,),
                    )
                else:
                    cur.execute("SELECT staff_name FROM staff_accounts ORDER BY staff_name")

                names = [row["staff_name"] for row in cur.fetchall()]
                if not names:
                    print(
                        "No matching staff accounts. Has the database been initialised?",
                        file=sys.stderr,
                    )
                    return 1

                results = []
                for staff_name in names:
                    password = "Svce@" + secrets.token_hex(3)
                    cur.execute(
                        """
                        UPDATE staff_accounts
                        SET password_hash = %s, must_change = TRUE
                        WHERE staff_name = %s
                        """,
                        (main.hash_password(password), staff_name),
                    )
                    results.append((staff_name, password))
            conn.commit()
    finally:
        # The connection pool starts worker threads; close it so the script
        # exits promptly instead of warning about a lingering thread.
        if main._pool is not None:
            main._pool.close()
            main._pool = None

    print()
    print(f"Reset {len(results)} staff password(s).")
    print("Each staff member sets their own password on the next login.")
    print()
    width = max(len(name) for name, _ in results)
    for staff_name, password in results:
        print(f"  {staff_name:<{width}}  {password}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
