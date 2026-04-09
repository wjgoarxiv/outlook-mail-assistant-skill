---
name: outlook-mail-assistant
description: "Ingests Outlook desktop mail from live Outlook profiles, .msg files, and .pst archives into a local workspace and SQLite index, then extracts summaries, tasks, deadlines, decisions, and follow-ups. Use this skill whenever the user wants Outlook mailbox analysis, weekly or monthly reporting from mail, local mail archiving, task extraction from Outlook messages, or Outlook-side actions such as moving mail or creating Outlook tasks/calendar items. Ask interactively for semantic ambiguity and high-impact execution."
---

# Outlook Mail Assistant

## Overview

Use this skill to turn local Outlook mail into structured work intelligence on Windows PCs without Graph or Exchange dependencies. The skill supports live Outlook desktop access, `.msg` imports, and `.pst` archive imports, then stores normalized records in a local workspace and SQLite index for reporting and follow-up workflows.

Deployment hygiene:
- never point `--workspace` at a path inside this skill directory
- keep workspaces under a user-private path such as `%USERPROFILE%\\Documents\\outlook-mail-assistant-runtime`
- treat exported JSONL, SQLite, CSV, XLSX, DOCX, and audit DB files as sensitive mail data

## Core Workflow

1. Select the ingestion surface:
   - live Outlook desktop profile
   - `.msg` file or directory
   - `.pst` archive
2. Normalize messages into the canonical schema described in `references/canonical-schema.md`.
3. Persist records to:
   - workspace artifacts for transparency
   - SQLite for search, reporting, and auditability
4. Extract:
   - weekly and monthly summaries
   - tasks, deadlines, decisions, and follow-ups
5. Ask one focused question at a time when:
   - semantic intent is ambiguous
   - an action would materially change mailbox, task, or calendar state
6. Produce outputs:
    - Markdown reports
    - DOCX reports via the bundled converter in this repo
    - XLSX or CSV outputs via the bundled exporters in this repo
   - Outlook tasks or calendar items

## Execution Boundaries

- Do not silently perform high-impact actions.
- Always ask before sending, deleting, bulk-moving, bulk-archiving, bulk-categorizing, or creating Outlook items from inferred intent.
- Record explicit facts, inferred facts, user-confirmed facts, and executed actions in the local audit trail.
- Prefer dry-run execution for automation scripts during testing and validation.

## Reference Files

- Read `references/canonical-schema.md` before changing normalization rules.
- Read `references/sqlite-schema.md` before changing database tables.

## Script Surface

Primary Python modules live under `scripts/outlook_mail_assistant/`.

- `canonical_schema.py`
  - canonical record model
  - stable dedupe hashing
- `storage.py`
  - workspace bootstrap
  - SQLite initialization

Keep adapters isolated. The rest of the skill should consume normalized records rather than source-specific Outlook data structures.
