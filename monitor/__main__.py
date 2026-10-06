"""CLI entry point.

    uv run python -m monitor doctor
    uv run python -m monitor folders
    uv run python -m monitor read --limit 20
    uv run python -m monitor watch --interval 30
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

from . import com, doctor, reader, state, decide, sources


def _stdio_utf8() -> None:
    """Force UTF-8 on the console.

    The Windows default code page is cp1252, and email bodies carry emoji,
    curly quotes and other unencodable characters, which crash printing
    outright (UnicodeEncodeError) rather than degrading.
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass


def _print(msg: dict, as_json: bool) -> None:
    if as_json:
        print(json.dumps(msg, indent=None, ensure_ascii=False))
    else:
        print(reader.summarize(msg))
        body = (msg.get("body") or "").strip()
        if body:
            print("   " + "\n   ".join(body[:600].splitlines()))
        d = (msg.get("decision") or {}).get("answers")
        if d:
            for qid in sorted(d):
                ans = d[qid]
                pick = (ans.get("choice") or ans.get("score")
                        or ans.get("probability"))
                print(f"   -> {qid}: {pick} "
                      f"(confidence {ans.get('confidence')})")
        print()


def _classify(msgs: list[dict], args) -> None:
    """Attach Jev-style decisions to each message, in place."""
    if not args.classify:
        return
    for m in msgs:
        m["decision"] = decide.decide(
            m, base_url=args.base_url, model=args.model)


def _classify_note(args) -> str:
    if not args.classify:
        return ""
    return f" [decisions via {args.model} @ {args.base_url}]"


def cmd_doctor(root: Path) -> int:
    results = doctor.run(root)
    print(doctor.render(results))
    return 0 if any(r["name"] == "connect" and r["ok"] for r in results) else 2


def _connect() -> object:
    """Connect, or exit with the reason the mail model is unreachable."""
    try:
        return com.connect()
    except LookupError as exc:
        raise SystemExit(f"COM host unavailable: {exc}")


def cmd_folders(args) -> int:
    app = _connect()
    folders = com.list_folders(app)
    for path in sorted(folders):
        print(f"{path}  ({com.count(folders[path])} msgs, "
              f"{com.unread_count(folders[path])} unread)")
    return 0


def cmd_read(args, root: Path) -> int:
    app = _connect()
    st = state.load(state.state_path(root))
    since = args.since or _resume_since(args, st)
    msgs = reader.read_folder(app, args.folder, limit=args.limit,
                              unread_only=args.unread_only, since=since)
    _classify(msgs, args)
    for m in msgs:
        _print(m, args.json)
    _note(st, args, msgs)
    state.save(state.state_path(root), st)
    if not args.json:
        print(f"{len(msgs)} message(s) from {args.folder}")
    return 0


def cmd_watch(args, root: Path) -> int:
    st = state.load(state.state_path(root))
    since = args.since or _resume_since(args, st)
    if since and not args.since:
        print(f"resuming after {since}")
    print(f"watching {args.folder!r} every {args.interval}s; "
          f"{len(state.seen_ids(st, args.folder))} already seen"
          f"{_classify_note(args)}")
    while True:
        try:
            # Reconnect each cycle: a cached COM object goes stale once
            # Outlook closes or the account set changes.
            app = _connect()
            msgs = reader.read_folder(app, args.folder, limit=args.limit,
                                      unread_only=args.unread_only,
                                      since=since)
            known = state.seen_ids(st, args.folder)
            fresh = [m for m in msgs if m.get("entry_id")
                     and m["entry_id"] not in known]
            _classify(fresh, args)
            for m in fresh:
                _print(m, args.json)
            # Record the scan even when nothing was new: last_run and the
            # resume point must survive a restart, and the seen list is
            # truncated at MAX_SEEN_PER_FOLDER so it cannot carry them.
            _note(st, args, msgs)
            if fresh:
                state.remember(st, args.folder,
                               [m["entry_id"] for m in fresh])
            state.save(state.state_path(root), st)
            if not args.json:
                print(f"-- {len(fresh)} new / {len(msgs)} total "
                      f"({time.strftime('%H:%M:%S')})")
        except KeyboardInterrupt:
            raise
        except Exception as exc:  # noqa: BLE001 - a monitor must not die
            print(f"cycle failed: {type(exc).__name__}: {exc}")
        time.sleep(args.interval)


def _resume_since(args, st: dict) -> str | None:
    """Where to start scanning when --since was not given."""
    if not args.resume:
        return None
    return state.resume_since(st, args.folder)


def _note(st: dict, args, msgs: list[dict]) -> None:
    """Record this scan into the caller's state object (caller saves it)."""
    state.note_run(
        st, args.folder,
        criteria={"limit": args.limit, "unread_only": args.unread_only,
                  "since": args.since, "interval": args.interval},
        received=[m.get("received_time") for m in msgs],
    )


def cmd_sources(args) -> int:
    """Show what each script advertises, which is what the model chooses from."""
    cfg = sources.load_config(args.sources)
    if not cfg:
        print(f"no sources configured (create {args.sources})")
        return 1
    for name in cfg:
        opts = sources.catalog(name, cfg)
        print(f"{name}  ({len(opts)} options)")
        for label, desc in opts.items():
            print(f"  {name}:{label}  {desc}")
    return 0


def cmd_pipeline(args, root: Path) -> int:
    """New mail -> model picks a lookup -> script runs -> items printed."""
    cfg = sources.load_config(args.sources)
    options = sources.catalogs(cfg) if cfg else {}
    questions = decide.build_questions(options)
    st = state.load(state.state_path(root))
    since = args.since or _resume_since(args, st)
    app = _connect()
    msgs = reader.read_folder(app, args.folder, limit=args.limit,
                              unread_only=args.unread_only, since=since)
    fresh = [m for m in msgs
             if m.get("entry_id") and m["entry_id"] not in
             state.seen_ids(st, args.folder)]
    for m in fresh:
        m["decision"] = decide.decide(m, questions=questions,
                                      base_url=args.base_url, model=args.model)
        _print(m, args.json)
        label, conf = decide.pick(m["decision"].get("answers", {}))
        if not label or label == "none":
            continue
        if conf < args.gate:
            if not args.json:
                print(f"   lookup candidate {label} at {conf:.2f} "
                      f"< {args.gate} - not run, needs review")
            continue
        try:
            items = sources.fetch(cfg, label)
        except (RuntimeError, OSError, ValueError, KeyError) as exc:
            print(f"   {label} fetch failed: {type(exc).__name__}: {exc}")
            continue
        if not args.json:
            print(f"   {label} ({conf:.2f}) -> {len(items)} item(s)")
            for it in items[:args.show]:
                print("   " + (it if isinstance(it, str)
                               else json.dumps(it, ensure_ascii=False)))
    _note(st, args, msgs)
    if fresh:
        state.remember(st, args.folder, [m["entry_id"] for m in fresh])
    state.save(state.state_path(root), st)
    if not args.json:
        print(f"{len(fresh)} new / {len(msgs)} total")
    return 0


def _detach(argv: list[str], root: Path, folder: str,
            log: str | None) -> int:
    """Re-run detached so the monitor survives the terminal closing.

    start_new_session=True is what keeps it alive: without it the child is
    killed when the parent shell exits.
    """
    inner = [a for a in argv if a != "--detach"]
    path = Path(log) if log else root / ".var" / f"watch-{folder}.log"
    path.parent.mkdir(parents=True, exist_ok=True)
    fh = path.open("a", encoding="utf-8", buffering=1)
    proc = subprocess.Popen(
        [sys.executable, "-u", "-m", "monitor", *inner],
        stdout=fh, stderr=subprocess.STDOUT, start_new_session=True,
    )
    print(f"detached pid={proc.pid}; log={path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    p = argparse.ArgumentParser(prog="monitor", description=__doc__)
    p.add_argument("command", choices=["doctor", "folders", "read", "watch",
                                       "sources", "pipeline"])
    p.add_argument("--folder", default="Inbox")
    p.add_argument("--limit", type=int, default=25)
    p.add_argument("--unread-only", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--since", default=None, metavar="ISO",
                   help="drop messages older than this ISO date, compared "
                        "against ReceivedTime.timestamp()")
    p.add_argument("--interval", type=int, default=30)
    p.add_argument("--resume", action="store_true",
                   help="start scanning after the newest message the last "
                        "run recorded for this folder (see .var/state.json); "
                        "overridden by an explicit --since")
    p.add_argument("--detach", action="store_true",
                   help="watch: re-run in the background and return, logging "
                        "to --log (default .var/watch-<folder>.log)")
    p.add_argument("--log", default=None, metavar="PATH",
                   help="watch --detach: where the background log goes")
    p.add_argument("--classify", action="store_true",
                   help="ask a System One model (Jev, or Von served locally) "
                        "for typed decisions about each message")
    p.add_argument("--base-url", default=decide.DEFAULT_BASE_URL,
                   metavar="URL",
                   help="System One endpoint; use http://127.0.0.1:8000 for "
                        "von serve, https://api.typesafe.ai for hosted Jev")
    p.add_argument("--model", default=decide.DEFAULT_MODEL,
                   help="jev-latest (hosted) or a local System One model id")
    p.add_argument("--sources", default=sources.CONFIG_PATH,
                   metavar="PATH", help="source scripts registry (json)")
    p.add_argument("--gate", type=float, default=decide.CONFIDENCE_GATE,
                   help="minimum confidence before a lookup is run")
    p.add_argument("--show", type=int, default=12,
                   help="how many fetched items to print")
    args = p.parse_args(argv)

    _stdio_utf8()
    root = Path.cwd()
    if args.command == "doctor":
        return cmd_doctor(root)
    if args.command == "folders":
        return cmd_folders(args)
    if args.command == "read":
        return cmd_read(args, root)
    if args.command == "watch":
        if args.detach:
            return _detach(argv, root, args.folder, args.log)
        return cmd_watch(args, root)
    if args.command == "sources":
        return cmd_sources(args)
    if args.command == "pipeline":
        return cmd_pipeline(args, root)
    return 1


if __name__ == "__main__":
    sys.exit(main())
