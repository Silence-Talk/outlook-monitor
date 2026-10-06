"""Seen-message bookkeeping, so watch mode prints each message exactly once.

Stored as plain JSON under .var/ so it is inspectable and disposable: delete
the file and the next run re-reads everything.
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path

STATE_FILENAME = "state.json"
MAX_SEEN_PER_FOLDER = 5000


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")


def state_path(root: Path | None = None) -> Path:
    base = root or Path.cwd()
    return base / ".var" / STATE_FILENAME


def load(path: Path) -> dict:
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            return {"seen": {}}
        data.setdefault("seen", {})
        return data
    except FileNotFoundError:
        return {"seen": {}}
    except (json.JSONDecodeError, OSError):
        # A corrupt state file must not brick the monitor; we lose dedup only.
        return {"seen": {}}


def save(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=1, sort_keys=True)
    os.replace(tmp, path)


def seen_ids(data: dict, folder: str) -> set[str]:
    return set(data.get("seen", {}).get(folder, []))


def remember(data: dict, folder: str, ids: list[str]) -> None:
    bucket = data.setdefault("seen", {})
    merged = list(dict.fromkeys(bucket.get(folder, []) + ids))
    if len(merged) > MAX_SEEN_PER_FOLDER:
        merged = merged[-MAX_SEEN_PER_FOLDER:]
    bucket[folder] = merged


def note_run(data: dict, folder: str, *, criteria: dict,
             received: list[str]) -> None:
    """Record what was scanned and when, so a restart can resume.

    EntryID dedup answers "did I already print this message?". It cannot answer
    "where should the next scan start?" once the seen list is truncated at
    MAX_SEEN_PER_FOLDER, so the newest ReceivedTime and the criteria that
    produced it are stored separately and survive that truncation.
    """
    runs = data.setdefault("runs", {})
    entry = runs.setdefault(folder, {})
    entry["criteria"] = {k: v for k, v in criteria.items() if v is not None}
    entry["last_run_utc"] = _utc_now()
    newest = max((r for r in received if r), default=None)
    if newest:
        # Keep the max, not the last: a scan that found nothing newer must not
        # rewind the resume point.
        prev = entry.get("last_seen_utc")
        entry["last_seen_utc"] = max(prev, newest) if prev else newest


def resume_since(data: dict, folder: str) -> str | None:
    """ISO date of the newest message already scanned in `folder`, or None."""
    return data.get("runs", {}).get(folder, {}).get("last_seen_utc")


def criteria(data: dict, folder: str) -> dict:
    """The criteria used by the last scan of `folder` (empty if never run)."""
    return dict(data.get("runs", {}).get(folder, {}).get("criteria", {}))
