"""Environment diagnosis: is the MAPI mail object model reachable?

Run this before anything else. It distinguishes the three states that look
identical from the outside:

  1. no COM server at all            -> One Outlook
  2. COM server, no MAPI namespace   -> build without the mail object model
  3. COM server + MAPI + folders     -> this project can read emails
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from .com import connect


def _local_app_data() -> Path:
    return Path(os.path.expandvars(r"%LOCALAPPDATA%"))


def _hkcu_office_outlook_flag() -> tuple[str, str]:
    """Read HKCU\\Software\\Microsoft\\Office\\Outlook -> EnableComAutomation."""
    try:
        import winreg
    except ImportError:
        return ("n/a", "winreg unavailable (not Windows)")
    key = None
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Office\Outlook")
        val, _vartype = winreg.QueryValueEx(key, "EnableComAutomation")
        return ("enabled" if val == 1 else "disabled",
                f"HKCU\\...\\Office\\Outlook\\EnableComAutomation = {val}")
    except OSError as exc:
        return ("absent", f"HKCU\\...\\Office\\Outlook\\EnableComAutomation: {exc.strerror}")
    finally:
        if key is not None:
            try:
                winreg.CloseKey(key)
            except OSError:
                pass


CHECKS = [
    ("platform", "Is this Windows? (pywin32 is Windows-only)"),
    ("pywin32", "Is pywin32 importable in this venv?"),
    ("one_outlook", "Is One Outlook installed? (%LOCALAPPDATA%\\Microsoft\\Olk)"),
    ("classic_profile", "Has classic Outlook been used? (%LOCALAPPDATA%\\Microsoft\\Outlook)"),
    ("classic_binary", "Does a classic OUTLOOK.EXE exist on disk?"),
    ("com_flag", "Is EnableComAutomation switched on?"),
    ("connect", "Does Dispatch('Outlook.Application') + GetNamespace('MAPI') work?"),
    ("mail_model", "Does the MAPI namespace yield a real Inbox with items?"),
]


def run(root: Path | None = None) -> list[dict]:
    lad = _local_app_data()
    results = []

    results.append({"name": "platform", "ok": sys.platform == "win32",
                    "detail": sys.platform})

    try:
        import win32com.client  # noqa: F401,F403
        import pythoncom  # noqa: F401,F403
        results.append({"name": "pywin32", "ok": True, "detail": "importable"})
    except ImportError as exc:
        results.append({"name": "pywin32", "ok": False, "detail": str(exc)})

    olk = lad / "Microsoft" / "Olk"
    results.append({"name": "one_outlook", "ok": olk.is_dir(),
                    "detail": f"{olk} {'present' if olk.is_dir() else 'absent'}"})

    classic = lad / "Microsoft" / "Outlook"
    results.append({"name": "classic_profile", "ok": classic.is_dir(),
                    "detail": f"{classic} {'present' if classic.is_dir() else 'absent'}"})

    exes = sorted(Path(os.path.expandvars(r"%ProgramFiles%")).glob(
        "Microsoft Office/root/*/OUTLOOK.EXE"))
    results.append({"name": "classic_binary", "ok": bool(exes),
                    "detail": f"{exes}" if exes else "no OUTLOOK.EXE under Program Files"})

    flag_state, flag_detail = _hkcu_office_outlook_flag()
    results.append({"name": "com_flag", "ok": flag_state == "enabled",
                    "detail": flag_detail})

    ns = None
    try:
        ns = connect()
        results.append({"name": "connect", "ok": True,
                        "detail": f"MAPI namespace: {ns!r} app={_app_detail(ns)}"})
    except LookupError as exc:
        results.append({"name": "connect", "ok": False, "detail": str(exc)})
    except Exception as exc:  # noqa: BLE001
        results.append({"name": "connect", "ok": False,
                        "detail": f"{type(exc).__name__}: {exc}"})

    if ns is not None:
        inbox = None
        try:
            inbox = ns.GetDefaultFolder(6)  # olFolderInbox
        except Exception as exc:  # noqa: BLE001
            results.append({"name": "mail_model", "ok": False,
                            "detail": f"GetDefaultFolder(6) failed: {exc}"})
        if inbox is not None:
            items = getattr(inbox, "Items", None)
            n = int(getattr(items, "Count", 0) or 0) if items else 0
            results.append({
                "name": "mail_model", "ok": n > 0,
                "detail": (f"Inbox={getattr(inbox, 'Name', '?')}, "
                           f"{n} items, {getattr(inbox, 'UnReadItemCount', 0)} unread"
                           if n else "Inbox resolved but contains 0 items"),
            })

    return results


def _app_detail(ns) -> str:
    from .com import app_info
    return str(app_info(ns))


def verdict(results: list[dict]) -> str:
    by = {r["name"]: r for r in results}
    if by.get("mail_model", {}).get("ok"):
        return "COM automation is live and the Inbox has items. `read` and `watch` should work."
    if by["connect"]["ok"]:
        return ("MAPI namespace resolved but no Inbox items: Outlook is open "
                "but has no readable messages.")
    if by["classic_binary"]["ok"]:
        return ("Classic Outlook is on disk but the MAPI namespace is not "
                "reachable. Launch Outlook before running this.")
    return "COM automation unreachable; see the individual check details."


def render(results: list[dict]) -> str:
    lines = []
    for r in results:
        mark = "OK  " if r["ok"] else "FAIL"
        lines.append(f"[{mark}] {r['name']}: {r['detail']}")
    lines.append("")
    lines.append(f"VERDICT: {verdict(results)}")
    return "\n".join(lines)
