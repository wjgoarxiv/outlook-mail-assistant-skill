from __future__ import annotations

import sys
from pathlib import Path

from outlook_mail_assistant.docx_export import convert_markdown_to_docx


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("usage: convert_md_to_docx.py <input.md> <output.docx>")
    output = convert_markdown_to_docx(Path(sys.argv[1]), Path(sys.argv[2]))
    print(output)


if __name__ == "__main__":
    main()
