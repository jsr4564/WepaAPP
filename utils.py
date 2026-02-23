"""Utility helpers for IDs, timestamps, and text normalization."""

from __future__ import annotations

from datetime import datetime
import uuid


def new_event_id() -> str:
    return str(uuid.uuid4())


def now_iso_timestamp() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def parse_iso_timestamp(value: str) -> datetime | None:
    text = (value or "").strip()
    if not text:
        return None

    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None

    if dt.tzinfo is None:
        # Treat timezone-naive values as local time for compatibility.
        dt = dt.replace(tzinfo=datetime.now().astimezone().tzinfo)
    return dt


def display_timestamp(value: str) -> str:
    dt = parse_iso_timestamp(value)
    if dt is None:
        return value or "-"
    return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")


def normalize_key_id(value: str) -> str:
    return (value or "").strip().upper()


def clean_text(value: str | None) -> str:
    return (value or "").strip()

