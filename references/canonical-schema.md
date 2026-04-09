# Canonical Schema

All ingestion adapters should normalize messages into a shared record shape before downstream reporting or action logic runs.

## Required fields

- `source_type`
- `source_path`
- `message_id`
- `subject`
- `sender_email`
- `received_at`
- `body_text`
- `dedupe_hash`

## Recommended optional fields

- `conversation_id`
- `sender_name`
- `to_recipients`
- `cc_recipients`
- `sent_at`
- `body_html`
- `attachment_manifest`
- `source_account`
- `source_folder`
- `imported_at`

## Normalization rules

- Store datetimes in UTC ISO-8601 strings ending with `Z`
- Preserve the original source path
- Compute `dedupe_hash` from stable identity fields
- Keep missing optional fields as `null` instead of inventing values
