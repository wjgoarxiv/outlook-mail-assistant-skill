from __future__ import annotations

import argparse
from pathlib import Path

from outlook_mail_assistant.storage import initialize_database


def main() -> None:
    parser = argparse.ArgumentParser(description="Initialize an Outlook mail SQLite database.")
    parser.add_argument("path", help="Path to the SQLite database file")
    args = parser.parse_args()

    db_path = Path(args.path)
    initialize_database(db_path)
    print(db_path)


if __name__ == "__main__":
    main()
