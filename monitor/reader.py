"""Read emails out of an Outlook folder and normalise them.

This is the only capability the project ships today. It reads; it does not
mutate. Anything that would change mailbox state is deliberately absent.
"""

from __future__ import annotations

import datetime as _dt

from .com import (connect, get_folder, list_folders, count, unread_count,
                  iter_items, newest_items, message)


def _since_ts(s: str) -> float:
    """Parse a --since bound to epoch seconds, the same unit as ReceivedTime."""
    d = _dt.datetime.fromisoformat(s)
    if d.tzinfo is None:
        d = d.replace(tzinfo=_dt.timezone.utc)
    return d.timestamp()


def _msg_ts(msg: dict) -> float | None:
    """Epoch seconds for a normalised message, via ReceivedTime.timestamp().

    Comparing timestamps rather than datetimes avoids the TypeError raised
    when a tz-aware pywintypes date meets a naive bound.
    """
    for key in ("received_time", "sent_on", "creation_time"):
        raw = msg.get(key)
        if not raw:
            continue
        try:
            d = _dt.datetime.fromisoformat(raw)
        except ValueError:
            continue
        if d.tzinfo is None:
            d = d.replace(tzinfo=_dt.timezone.utc)
        return d.timestamp()
    return None


def read_folder(app, folder_name: str, *, limit: int | None = None,
                unread_only: bool = False,
                since: str | None = None) -> list[dict]:
    """Return normalised messages from `folder_name`, newest first.

    newest_items asks Outlook to sort the collection, so the scan is already
    in date order and can stop early: `limit` stops after N messages, and
    `since` stops at the first message older than the bound. unread_only
    keeps scanning until N unread are found, which is why it is slower —
    MAPIFolder.UnReadItemCount is a count, not a selector.
    """
    folder = get_folder(app, folder_name)
    if folder is None:
        available = sorted(list_folders(app))
        raise LookupError(
            f"Folder {folder_name!r} not found. Available: {available}"
        )

    bound = _since_ts(since) if since else None
    out = []
    for item in newest_items(folder, None if unread_only else limit):
        msg = message(item)
        if unread_only and not msg.get("unread"):
            continue
        # newest_items yields date-descending, so once a message is older
        # than the bound, every later one is too.
        if bound is not None and (_msg_ts(msg) or 0.0) < bound:
            break
        out.append(msg)
        if limit is not None and len(out) >= limit:
            break
    return out


def summarize(msg: dict) -> str:
    """One-line human rendering of a message."""
    subj = (msg.get("subject") or "(no subject)").strip()
    # Identity order: SMTP address is the stable one; MailItem.From is absent
    # on this build and Sender is only a display name.
    frm = (msg.get("sender_email") or msg.get("sender_name")
           or msg.get("sender") or "?").strip()
    recv = (msg.get("received_time") or msg.get("sent_on")
            or msg.get("creation_time") or "-")
    flag = "UNREAD " if msg.get("unread") else ""
    return f"{recv} | {flag}{subj} | from {frm}"
