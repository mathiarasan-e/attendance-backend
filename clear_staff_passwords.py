"""Clear a staff password so the staff member can choose a new one.

Staff own their passwords: the backend never generates one and never prints one.
If a staff member forgets theirs, clear it here and they set a new one from the
portal using the college's setup code (``STAFF_SETUP_CODE``).

Usage, from the attendance_backend directory:

    venv\\Scripts\\python clear_staff_passwords.py            # asks to confirm
    venv\\Scripts\\python clear_staff_passwords.py --yes      # no confirmation
    venv\\Scripts\\python clear_staff_passwords.py --staff "Ms. S. Saranya"

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


def run() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--yes", action="store_true", help="skip the confirmation prompt")
    parser.add_argument(
        "--staff",
        help="clear only this staff member (must match staff_accounts.staff_name exactly)",
    )
    args = parser.parse_args()

    if not main.DATABASE_URL:
        print("DATABASE_URL is not set. Nothing to do.", file=sys.stderr)
        return 1

    host = main.DATABASE_URL.rsplit("@", 1)[-1].split("/")[0]
    print(f"Target database host : {host}")

    # Bring the schema up to date before touching any rows. An older database
    # still has password_hash NOT NULL, which would otherwise reject the clear
    # below. init_database() is idempotent and is the migration the app runs.
    main.init_database()

    if not args.yes:
        answer = input("Type 'yes' to clear staff password(s): ").strip().lower()
        if answer != "yes":
            print("Aborted. Nothing changed.")
            return 1

    try:
        with main.get_connection() as conn:
            with conn.cursor() as cur:
                if args.staff:
                    cur.execute(
                        """
                        SELECT staff_name, password_hash
                        FROM staff_accounts
                        WHERE staff_name = %s
                        """,
                        (args.staff,),
                    )
                else:
                    cur.execute(
                        """
                        SELECT staff_name, password_hash
                        FROM staff_accounts
                        ORDER BY staff_name
                        """
                    )

                accounts = cur.fetchall()
                if not accounts:
                    print(
                        "No matching staff accounts. Has the database been initialised?",
                        file=sys.stderr,
                    )
                    return 1

                cleared = []
                already_clear = []
                for row in accounts:
                    if row["password_hash"] is None:
                        already_clear.append(row["staff_name"])
                        continue
                    cur.execute(
                        """
                        UPDATE staff_accounts
                        SET password_hash = NULL,
                            must_change = TRUE,
                            updated_at = NOW()
                        WHERE staff_name = %s
                        """,
                        (row["staff_name"],),
                    )
                    cleared.append(row["staff_name"])
            conn.commit()
    finally:
        # The connection pool starts worker threads; close it so the script
        # exits promptly instead of warning about a lingering thread.
        if main._pool is not None:
            main._pool.close()
            main._pool = None

    print()
    print(f"Cleared {len(cleared)} staff password(s).")
    if cleared:
        for staff_name in cleared:
            print(f"  {staff_name}")
        print()
        print("Each staff member signs in, chooses 'Create your password', and picks")
        print("their own using the setup code.")
    if already_clear:
        print()
        print("Already had no password:")
        for staff_name in already_clear:
            print(f"  {staff_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
