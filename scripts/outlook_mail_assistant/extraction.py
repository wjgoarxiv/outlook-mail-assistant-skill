from __future__ import annotations

import re


STRONG_EXPLICIT_PATTERNS = [
    r"\bplease\b",
    r"\bkindly\b",
    r"\bcan you\b",
    r"\bcould you\b",
    r"\baction required\b",
    r"검토 요청",
    r"회신 요청",
    r"확인 요청",
    r"입력 요청",
    r"제출 요청",
    r"\brequest for input\b",
    r"\brequest\b",
    r"요청",
]

WEAK_EXPLICIT_PATTERNS = [
    r"확인",
    r"검토",
    r"참석",
    r"입력",
    r"제출",
]

INFERRED_PATTERNS = [
    r"remind",
    r"reminder",
    r"deadline",
    r"due",
    r"meeting",
    r"webinar",
    r"교육",
    r"마감",
    r"일정",
    r"회의",
    r"세미나",
]

IGNORE_PATTERNS = [
    r"newsletter",
    r"monthly news",
    r"멘션",
    r"訃告",
    r"부고",
    r"\bstatement\b",
    r"\bnotification\b",
    r"알림",
    r"\bfyi\b",
]
NOTICE_PATTERNS = [
    r"\bannouncement\b",
    r"공지",
    r"안내",
    r"notice",
]

DEADLINE_PATTERNS = [r"deadline", r"due", r"마감", r"기한", r"까지", r"by\s+\w+"]
MEETING_PATTERNS = [r"meeting", r"webinar", r"workshop", r"expo", r"회의", r"세미나", r"미팅", r"교육"]
DECISION_PATTERNS = [r"\bdecision\b", r"\bdecided\b", r"결정", r"최종 결정", r"승인 요청", r"승인 필요"]
EVENT_ACTION_PATTERNS = [
    r"register",
    r"registration",
    r"join",
    r"attend",
    r"reply",
    r"rsvp",
    r"review agenda",
    r"submit",
    r"prepare",
    r"참석 여부 회신",
    r"참석 요청",
    r"등록",
    r"신청",
    r"회신",
    r"사전 검토",
    r"회의자료",
]

SUBJECT_PREFIX_PATTERNS = [
    r"^(?:\s*(?:re|fw|fwd)\s*:\s*)+",
]

SUBJECT_TAG_PATTERNS = [
    r"^\[(?:external|remind|fyi)\]\s*",
]

LATEST_BODY_SPLIT_PATTERNS = [
    r"\nfrom:\s",
    r"\n보낸사람:\s",
    r"\n-----original message-----",
    r"\n________________________________",
]


def normalize_subject(subject: str | None) -> str:
    normalized = (subject or "").strip()
    changed = True
    while changed:
        changed = False
        for pattern in SUBJECT_PREFIX_PATTERNS + SUBJECT_TAG_PATTERNS:
            updated = re.sub(pattern, "", normalized, flags=re.I).strip()
            if updated != normalized:
                normalized = updated
                changed = True
    return normalized


def extract_latest_body_segment(body_text: str | None) -> str:
    text = body_text or ""
    lowered = text.lower()
    cut_positions = []
    for pattern in LATEST_BODY_SPLIT_PATTERNS:
        match = re.search(pattern, lowered, re.I)
        if match:
            cut_positions.append(match.start())
    if cut_positions:
        text = text[: min(cut_positions)]
    return text.strip()


def _has_any(patterns, text: str) -> bool:
    return any(re.search(pattern, text, re.I) for pattern in patterns)


def classify_work_item_signal(*, subject: str | None, body_text: str | None):
    normalized_subject = normalize_subject(subject)
    latest_body = extract_latest_body_segment(body_text)
    combined = "\n".join(filter(None, [normalized_subject, latest_body]))
    lowered = combined.lower()

    if _has_any(IGNORE_PATTERNS, lowered):
        return None

    has_strong_explicit = _has_any(STRONG_EXPLICIT_PATTERNS, lowered)
    has_weak_explicit = _has_any(WEAK_EXPLICIT_PATTERNS, lowered)
    has_deadline = _has_any(DEADLINE_PATTERNS, lowered)
    has_meeting = _has_any(MEETING_PATTERNS, lowered)
    has_event_action = _has_any(EVENT_ACTION_PATTERNS, lowered)
    has_decision = _has_any(DECISION_PATTERNS, lowered)
    has_notice = _has_any(NOTICE_PATTERNS, lowered)

    if has_notice and not (has_strong_explicit or has_event_action):
        return None

    if has_strong_explicit or (has_weak_explicit and has_deadline):
        if has_meeting and has_event_action:
            return {
                "kind": "meeting",
                "confidence": "inferred",
                "reason": "meeting/event wording with action cue detected",
                "needs_confirmation": True,
            }
        return {
            "kind": "task",
            "confidence": "explicit",
            "reason": "request-like wording detected",
            "needs_confirmation": False,
        }

    if has_meeting and has_event_action:
        return {
            "kind": "meeting",
            "confidence": "inferred",
            "reason": "meeting/event wording with action cue detected",
            "needs_confirmation": True,
        }

    if has_deadline and (has_strong_explicit or has_weak_explicit):
        return {
            "kind": "deadline",
            "confidence": "inferred",
            "reason": "deadline wording with action cue detected",
            "needs_confirmation": True,
        }

    if has_decision and (has_strong_explicit or has_deadline):
        return {
            "kind": "decision_signal",
            "confidence": "inferred",
            "reason": "decision wording detected",
            "needs_confirmation": True,
        }

    if _has_any(INFERRED_PATTERNS, lowered) and (has_strong_explicit or has_event_action):
        return {
            "kind": "task",
            "confidence": "inferred",
            "reason": "reminder/deadline wording with action cue detected",
            "needs_confirmation": True,
        }

    return None


def extract_task_candidates(records):
    items = []
    seen = {}
    for record in records:
        signal = classify_work_item_signal(
            subject=record.get("subject"),
            body_text=record.get("body_text"),
        )
        if not signal:
            continue

        normalized_subject = normalize_subject(record.get("subject"))
        dedupe_key = (
            record.get("store_name"),
            record.get("sender_email"),
            normalized_subject,
            signal["kind"],
        )
        item = {
            "message_id": record.get("message_id"),
            "store_name": record.get("store_name"),
            "folder_path": record.get("folder_path"),
            "subject": record.get("subject"),
            "sender_email": record.get("sender_email"),
            "received_at": record.get("received_at"),
            "kind": signal["kind"],
            "confidence": signal["confidence"],
            "reason": signal["reason"],
            "snippet": extract_latest_body_segment(record.get("body_text"))[:240],
            "needs_confirmation": signal["needs_confirmation"],
        }
        previous = seen.get(dedupe_key)
        if previous is None or (item.get("received_at") or "") > (previous.get("received_at") or ""):
            seen[dedupe_key] = item

    items = sorted(seen.values(), key=lambda item: item.get("received_at") or "", reverse=True)
    return items


def persist_task_candidates(connection, records) -> int:
    inserted = 0
    for record in records:
        before = connection.total_changes
        connection.execute(
            """
            INSERT INTO tasks (
                message_id, store_name, folder_path, subject, sender_email,
                received_at, kind, confidence, reason, snippet, needs_confirmation
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                record.get("message_id"),
                record.get("store_name"),
                record.get("folder_path"),
                record.get("subject"),
                record.get("sender_email"),
                record.get("received_at"),
                record.get("kind"),
                record.get("confidence"),
                record.get("reason"),
                record.get("snippet"),
                int(bool(record.get("needs_confirmation"))),
            ),
        )
        inserted += int(connection.total_changes > before)
    connection.commit()
    return inserted
