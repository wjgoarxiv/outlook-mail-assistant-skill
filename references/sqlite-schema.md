# SQLite Schema

The initial local database must contain these tables:

- `messages`
- `tasks`
- `decisions`
- `audit_log`

## Purpose

- `messages`: normalized mail index
- `tasks`: extracted explicit or user-confirmed work items
- `decisions`: extracted decisions and commitments
- `audit_log`: execution trace for imports and later Outlook actions

## Initial design notes

- Use `CREATE TABLE IF NOT EXISTS`
- Keep the first migration idempotent
- Allow later ingestion and reporting slices to extend the schema without breaking the first bootstrap
