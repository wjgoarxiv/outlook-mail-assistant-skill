from __future__ import annotations

import argparse
from pathlib import Path

from outlook_mail_assistant.storage import bootstrap_workspace


def main() -> None:
    parser = argparse.ArgumentParser(description="Initialize an Outlook mail workspace.")
    parser.add_argument("path", help="Workspace directory to create")
    args = parser.parse_args()

    workspace = bootstrap_workspace(Path(args.path))
    print(workspace.manifest_path)


if __name__ == "__main__":
    main()
