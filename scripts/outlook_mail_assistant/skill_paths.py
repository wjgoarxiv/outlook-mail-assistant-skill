from __future__ import annotations

import os
from pathlib import Path


def _candidate_skill_roots() -> list[Path]:
    roots: list[Path] = []
    env_roots = [
        os.getenv("OMA_SKILLS_DIR"),
        os.getenv("CODEX_SKILLS_DIR"),
        os.getenv("CODEX_HOME"),
    ]
    for value in env_roots:
        if not value:
            continue
        root = Path(value).expanduser()
        roots.append(root / "skills" if root.name.lower() == ".codex" else root)

    home = Path.home()
    roots.extend(
        [
            home / "skills",
            home / ".vscode" / "skills",
            home / ".codex" / "skills",
            home / ".agents" / "skills",
        ]
    )

    unique: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        key = str(root).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(root)
    return unique


def find_skill_script(*relative_parts_groups: tuple[str, ...]) -> Path | None:
    for root in _candidate_skill_roots():
        for parts in relative_parts_groups:
            candidate = root.joinpath(*parts)
            if candidate.exists():
                return candidate
    return None
