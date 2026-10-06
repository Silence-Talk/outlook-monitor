"""Outlook COM Automation Client bindings (pywin32 / win32com).

Scope: connection, folder discovery, and message enumeration ONLY.
No send / delete / move / mark-read calls live here.

The entry point is NOT Application.Inbox. That property does not exist in the
Outlook 16.0 type library. The mail object model hangs off the MAPI namespace:

    app  = Dispatch("Outlook.Application")
    ns   = app.GetNamespace("MAPI")            # <- the nickname, not a URI
    inbox = ns.GetDefaultFolder(6)             # 6 == olFolderInbox
    items = inbox.Items                        # collection
    item  = items.Item(i)                      # 1-BASED: Item(0) throws
                                               # "Array index out of bounds."

Folder ids are the OlDefaultFolders enum, read off the generated type library.
"""

from __future__ import annotations

import datetime as _dt
import time
import win32com.client

try:
    import pythoncom
except ImportError:  # pragma: no cover
    pythoncom = None

try:
    from pywintypes import com_error
except ImportError:  # pragma: no cover
    class com_error(Exception):
        pass

# Namespace nickname Outlook registers for the mail object model. A URI is
# rejected outright ("The parameter is incorrect."), which is the single
# most common dead end when first wiring this up.
MAPI_NAMESPACE = "MAPI"

# The COM bridge is not reliably ready on the first Dispatch: a thread that
# has never touched COM can fail outright, and Outlook may not have
# registered its servers yet while it finishes signing in. CoInitialize plus
# a short retry loop absorbs both. Verified `Name` afterwards so a ProgID
# answered by some other Office server cannot pass as Outlook.
DISPATCH_PROGID = "Outlook.Application"
DISPATCH_ATTEMPTS = 3
DISPATCH_RETRY_DELAY_S = 1.0
EXPECTED_APP_NAME = "Outlook"

# OlDefaultFolders, transcribed from the generated Outlook 16.0 type library.
FOLDER_IDS = {
    "inbox": 6,
    "sent": 5, "sentitems": 5, "sentmail": 5,
    "draft": 16, "drafts": 16,
    "outbox": 4,
    "trash": 3, "deleted": 3, "deleteditems": 3,
    "junk": 23,
    "tasks": 13,
    "contacts": 10,
    "calendar": 9,
    "notes": 12,
    "conflicts": 19,
    "localfailures": 21,
    "serverfailures": 22,
}

# Properties this build actually answers. Verified against a live MailItem:
# the classic names (From / Read / ReceivedDate / ID / Flagged) are NOT
# present on Outlook 16.0. The real date accessors are ReceivedTime and SentOn.
MESSAGE_FIELDS = {
    "subject": "Subject",
    "to": "To",
    "body": "Body",
    "html_body": "HTMLBody",
    "entry_id": "EntryID",
    "received_time": "ReceivedTime",
    "sent_on": "SentOn",
    "creation_time": "CreationTime",
    "flag_status": "FlagStatus",
    "unread": "UnRead",
    "conversation_topic": "ConversationTopic",
}
_DATETIME_FIELDS = {"received_time", "sent_on", "creation_time"}
_ABSENT_ON_16_0 = ("From", "Read", "ReceivedDate", "ID", "Flagged", "Labels")


def _dispatch_app():
    """Dispatch the Outlook COM server, retrying a cold COM bridge."""
    if pythoncom is not None:
        pythoncom.CoInitialize()
    last_exc = None
    for attempt in range(1, DISPATCH_ATTEMPTS + 1):
        try:
            return win32com.client.Dispatch(DISPATCH_PROGID)
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            if attempt < DISPATCH_ATTEMPTS:
                time.sleep(DISPATCH_RETRY_DELAY_S)
    raise LookupError(
        f"Dispatch({DISPATCH_PROGID!r}) failed after {DISPATCH_ATTEMPTS} "
        f"attempts ({DISPATCH_RETRY_DELAY_S}s apart): {last_exc}. "
        "Outlook must be running for its COM servers to be registered."
    )


def connect() -> object:
    """Return the MAPI namespace, or raise a readable error."""
    app = _dispatch_app()

    name = _read(app, "Name")
    if name is None or str(name).strip() != EXPECTED_APP_NAME:
        raise LookupError(
            f"ProgID {DISPATCH_PROGID!r} answered with Name={name!r}, not "
            f"{EXPECTED_APP_NAME!r}: this is not the Outlook COM server."
        )

    try:
        ns = app.GetNamespace(MAPI_NAMESPACE)
    except com_error as exc:
        raise LookupError(
            f"GetNamespace({MAPI_NAMESPACE!r}) failed: {exc}. "
            "One Outlook registers no MAPI namespace at all."
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise LookupError(f"GetNamespace({MAPI_NAMESPACE!r}) failed: {exc}") from exc
    if ns is None:
        raise LookupError("GetNamespace('MAPI') returned nothing.")
    return ns


def _read(obj, name, default=None):
    """Read one COM property, swallowing per-property COM failures."""
    try:
        return getattr(obj, name)
    except com_error:
        return default
    except AttributeError:
        return default
    except Exception:  # noqa: BLE001
        return default


def app_info(ns) -> dict:
    app = _read(ns, "Application") or ns
    out = {}
    for key in ("Name", "Version", "Build"):
        val = _read(app, key)
        if val is not None:
            out[key.lower()] = val
    return out


def get_folder(ns, name: str):
    """Resolve a folder by friendly name, then by walking the folder tree."""
    key = name.strip().lower()
    if key in FOLDER_IDS:
        return _read(ns, "GetDefaultFolder")(FOLDER_IDS[key])

    for folder, path in walk(ns):
        if path.lower() == key or path.rsplit("/", 1)[-1].lower() == key:
            return folder
    return None


def list_folders(ns, max_depth: int = 3) -> dict:
    return {path: folder for folder, path in walk(ns, max_depth=max_depth)}


def _collection(obj):
    """Yield members of a win32com collection (Count + 1-based Item)."""
    if obj is None:
        return []
    count = _read(obj, "Count") or 0
    item = _read(obj, "Item")
    if not callable(item):
        return []
    out = []
    for i in range(1, int(count) + 1):
        try:
            out.append(item(i))
        except com_error:
            continue
        except Exception:  # noqa: BLE001
            continue
    return out


def walk(ns, max_depth: int = 3):
    """Yield (folder, "Parent/Child") over top-level folders and subfolders."""
    out = []
    for top in _collection(_read(ns, "Folders")):
        name = _read(top, "Name") or "?"
        out.append((top, name))
        _descend(top, name, 1, max_depth, out)
    return out


def _descend(folder, path, depth, max_depth, out):
    if depth >= max_depth:
        return
    for sub in _collection(_read(folder, "Folders")):
        name = _read(sub, "Name") or "?"
        child_path = f"{path}/{name}"
        out.append((sub, child_path))
        _descend(sub, child_path, depth + 1, max_depth, out)


def count(folder) -> int:
    return int(_read(folder, "Items").Count or 0) if _read(folder, "Items") else 0


def unread_count(folder) -> int:
    # MAPIFolder exposes UnReadItemCount, not UnreadCount.
    return int(_read(folder, "UnReadItemCount") or 0)


def iter_items(folder, limit: int | None = None):
    """Yield MailItem objects. Item() is 1-based; Item(0) throws."""
    items = _read(folder, "Items")
    if items is None:
        return
    total = int(_read(items, "Count") or 0)
    stop = total if limit is None else min(total, limit)
    for i in range(1, stop + 1):
        try:
            yield items.Item(i)
        except com_error:
            continue


# `Items` is NOT date-ordered. Measured on a live Inbox of ~2000 items: the
# first three items were dated Oct 4, Oct 4, Oct 5 while the last three were
# Oct 1, Oct 2, Oct 5 — and the newest message of all sat at the last index. So
# taking the first N yields an arbitrary subset that can miss new arrivals.
#
# Outlook sorts the collection itself, and it is far cheaper than sorting in
# Python: items.Sort("ReceivedTime", True) reorders the live _Items object in
# place in 0.001-0.02 s for ~2000 items, versus 6.3 s to read ReceivedTime for
# every item and sort in Python. Newest-50 costs ~0.1 s total. Legal sort
# fields include ReceivedTime, SentOn, Subject, From; "MessageArrived" is
# rejected ('The property "MessageArrived" is unknown.'). Sort() with no args
# raises 'Type mismatch.' Sort mutates in place, and re-fetching folder.Items
# returns a FRESH, UNSORTED collection — so the sorted object must be the one
# iterated.
def newest_items(folder, limit: int | None = None, *, key: str = "ReceivedTime"):
    """Yield MailItem objects newest-first, using Outlook's own Sort."""
    items = _read(folder, "Items")
    if items is None:
        return
    total = int(_read(items, "Count") or 0)
    stop = total if limit is None else min(total, limit)
    try:
        items.Sort(key, True)
    except com_error:
        # Older/other builds may lack Sort; fall back to sorting in Python.
        yield from _newest_python_sort(folder, limit, key)
        return
    for i in range(1, stop + 1):
        try:
            yield items.Item(i)
        except com_error:
            continue


def _newest_python_sort(folder, limit, key):
    """Fallback: read `key` for every item and sort in Python."""
    items = _read(folder, "Items")
    if items is None:
        return
    total = int(_read(items, "Count") or 0)
    epoch = _dt.datetime(1970, 1, 1, tzinfo=_dt.timezone.utc)
    pairs = []
    for i in range(1, total + 1):
        try:
            m = items.Item(i)
        except com_error:
            continue
        except Exception:  # noqa: BLE001
            continue
        pairs.append((_read(m, key) or epoch, m))
    pairs.sort(key=lambda t: t[0], reverse=True)
    if limit is not None:
        pairs = pairs[:limit]
    for _, m in pairs:
        yield m


def message(item) -> dict:
    """Normalise one MailItem. Never raises."""
    out: dict = {}
    for key, com_name in MESSAGE_FIELDS.items():
        val = _read(item, com_name)
        if val is None:
            continue
        if key in _DATETIME_FIELDS:
            out[key] = _iso(val)
        elif key == "unread":
            out[key] = bool(val)
        elif key == "flag_status":
            out[key] = int(val or 0)
        else:
            out[key] = str(val)

    # `MailItem.From` does not exist on this build, and `Sender` is a display
    # name ("Instagram", "Google") rather than an identity. The reference
    # implementation reaches the SMTP address through the MAPI proptag
    # MAPI_PR_SMTP_ADDRESS (0x39FE001F); that route is unavailable here —
    # PropertyAccessor resolves but its generated stub exposes no
    # GetPropertyText/GetProperty, so an arbitrary proptag cannot be asked
    # for. The direct accessors return the same value, so they are the port.
    addr = _read(item, "SenderEmailAddress")
    if addr:
        out["sender_email"] = str(addr)
        etype = _read(item, "SenderEmailType")
        if etype:
            out["sender_email_type"] = str(etype)
    name = _read(item, "SenderName")
    if name:
        out["sender_name"] = str(name)
    if not out.get("sender_email") and not out.get("sender_name"):
        sender = _read(item, "Sender")
        if sender is not None:
            out["sender"] = str(_read(sender, "Name") or _read(sender, "Address") or "")

    atts = _read(item, "Attachments")
    if atts is not None:
        out["attachment_count"] = int(_read(atts, "Count") or 0)
    return out


def _iso(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.isoformat()
    fn = getattr(value, "isoformat", None)
    if callable(fn):
        try:
            return fn()
        except Exception:  # noqa: BLE001
            pass
    return str(value)
