# outlook_access

Monitor the Outlook desktop and read emails. **Read-only** — nothing here sends,
drafts, moves, deletes, or marks messages as read.

## Setup

```bash
uv venv .venv --python 3.12      # CPython 3.12.13
uv pip install pywin32           # pywin32 312
```

## Usage

```bash
uv run python -m monitor doctor              # environment diagnosis, run first
uv run python -m monitor folders             # list folders + message counts
uv run python -m monitor read --limit 20     # one-shot read of Inbox
uv run python -m monitor read --unread-only --json
uv run python -m monitor read --since 2026-10-05 --limit 50
uv run python -m monitor watch --interval 30 # print each message once
uv run python -m monitor watch --interval 300 --detach  # background daemon
uv run python -m monitor watch --resume      # restart where the last run stopped
uv run python -m monitor read --limit 5 --classify       # Jev-style decisions
```

Flags: `--folder <name>` `--limit <n>` `--unread-only` `--json` `--since <ISO>`
`--interval <s>` `--resume` `--detach` `--log <path>` `--classify`
`--base-url <url>` `--model <id>` `--regen-stubs` (doctor).

State lives in `.var/state.json` (gitignored): an `EntryID` dedup list per
folder for `watch`, plus a `runs` block recording what each folder's last scan
looked like (see Background fetch).

## Picking "the recent messages"

`Items` is unordered, so "recent" has to be derived. What works, measured on
a live Inbox of ~2000 items:

| route | cost | verdict |
|---|---|---|
| **`items.Sort("ReceivedTime", True)` then `Item(1..N)`** | **0.57 s** | **default** — native, correct, fast |
| read every `ReceivedTime`, sort in Python | 6.3 s dates-only, 13.3 s full cycle | fallback, only if `Sort` raises |
| first N of `Items` | 0.5 s | **wrong** — unordered, so new arrivals are missed |
| last N of `Items` (`Count - N`) | 0.5 s | **wrong** — falsified, see below |

`Sort` is in-place on the `_Items` you already hold; re-fetching `folder.Items`
returns a fresh *unsorted* collection, so sort once per fetch. Legal property
names: `ReceivedTime`, `SentOn`, `Subject`, `From`. `Sort()` with no argument
raises `Type mismatch.`; `Sort("MessageArrived", True)` raises
`The property "MessageArrived" is unknown.`

`--since` stops the walk at the first message older than the bound, so a date
filter costs 0.22 s instead of a full scan.

### The tail window is falsified — do not reach for it

The `Count - N` shortcut (take the last N of `Items`) was tested and is
**unsound**: with `--window 200` on the live Inbox the top 5 were Oct 5, Oct 5,
Oct 2, Oct 1, Sep 30, while a full scan's top 5 were all Oct 5. Three Oct 5
messages sat **outside** the last 200 indices. Arrivals do not reliably append
to `Items`. It has been deleted from the code; `Sort` is both faster and
correct.

## How the mail object model is reached

The mail object model hangs off the **MAPI namespace**, not off `Application`:

```python
app   = win32com.client.Dispatch("Outlook.Application")
ns    = app.GetNamespace("MAPI")        # nickname, NOT a URI
inbox = ns.GetDefaultFolder(6)          # 6 == olFolderInbox
items = inbox.Items                     # _Items, 1-based: Item(1)
```

`GetNamespace` rejects URIs (`com_error(-2147024809, 'The parameter is
incorrect.')`); only `"MAPI"` works. `GetNewNickNames()` returns `None`.

Folder ids: inbox 6, sent 5, draft 16, outbox 4, trash 3, junk 23, tasks 13,
contacts 10, calendar 9, notes 12, conflicts 19, localfailures 21,
serverfailures 22.

## Findings on this machine (verified — do not re-derive)

| probe | result |
|---|---|
| COM gate | `HKCU\Software\Microsoft\Office\Outlook` → `EnableComAutomation` = `REG_DWORD 0x1` (set) |
| Outlook | classic mode; `OUTLOOK.EXE` under `Office16`; One Outlook dir `Olk` present but registers no MAPI namespace |
| COM host | `Outlook.Application` → `Name=Outlook`, `Version=16.0.0.20430` |
| Namespace handle | `GetNamespace("MAPI")` → `_NameSpace` — the only one that works |
| Inbox | `ns.GetDefaultFolder(6)` → `MAPIFolder`, ~2000 items, nearly all unread |
| Namespace folders | `GetFolderFromID(6, 1)` → `<account>/Inbox` |
| Item access | `Items.Item(1)` → `_MailItem` with a real subject; **1-based** (`Item(0)` → `Array index out of bounds.`) |
| Folder tree | every folder sits under one account; a second account contributes no folders to the COM tree |
| Folder unread count | **`MAPIFolder.UnReadItemCount`** (not `UnreadCount`) |
| MailItem absent | `MailItem.ID`, `Read`, `ReceivedDate`, `Flagged`, `Labels` → `AttributeError` |
| MailItem present | `EntryID` (read-only), `Subject`, `From`, `UnRead` (read/write), `SenderEmailAddress`, `SenderEmailType`, `SenderName`, `ReceivedTime`, `SentOn` |
| `_Items` collection API | `Sort`, `Restrict`, `Find`, `FindNext`, `GetFirst`, `GetLast`, `GetNext`, `GetPrevious`, `SetColumns`, `RawTable`, `Count`, `Item` (`SortEx` absent) |
| Absent APIs | `Items.SortOrder` (property), `MAPIFolder.FindItems` |
| DASL filter | `items.Restrict('@SQL="urn:schemas:httpmail:datereceived" >= \'2026-10-01\'')` → 10 items in 0.01 s, still sortable |
| `Find` | wants DASL (`[Subject] = 'x'` + `FindNext()`); rejects AQS (`subject: x` → `Condition is not valid.`) |
| `SetColumns` | needs a CSV string (`"Subject,ReceivedTime"`); a Python list is stringified and rejected. No measurable gain over `Item()` |
| proptag route | `PropertyAccessor` absent; `0x39FE001F` (`MAPI_PR_SMTP_ADDRESS`) unreachable — but the direct `SenderEmailAddress` accessor works |

## Robustness ports

### 1. Cold COM bridge — `pythoncom.CoInitialize()` + retry

`Dispatch("Outlook.Application")` intermittently raises before Outlook's
pythoncom bridge is warm. `com.py` calls `CoInitialize()` then retries
`Dispatch` 3 attempts / 1.0 s apart, and verifies `Name == "Outlook"`.

### 2. Real sender identity

`SenderEmailAddress` + `SenderEmailType` (`SMTP`/`EX`) + `SenderName` are the
real identity; `From` alone is not. Verified live on a service notification:
`<sender-service>@mail.<service>.com` / `SMTP` / `<SenderName>`.

### 3. Date filtering across timezones

`ReceivedTime.timestamp()` epoch seconds vs
`datetime.fromisoformat(--since).timestamp()` — sidesteps the tz-aware vs naive
`TypeError`.

## Message fields emitted

`entry_id`, `folder`, `unread`, `subject`, `from`, `sender_email`,
`sender_email_type`, `sender_name`, `received_time`, `sent_on`, `body`.

## Console encoding

The Windows default code page is cp1252; email bodies carry emoji and curly
quotes that crash `print` outright (`UnicodeEncodeError`). `_stdio_utf8()`
reconfigures stdout/stderr to UTF-8 with `errors="replace"`.

## Traps that produced a false "COM is dead" verdict

Probing `app.Inbox`, `app.Folders`, `app.SentItems`, `app.Drafts` yields
`AttributeError`, and `Dispatch(IID('{00061033-…-…}'))` (MailItem) yields
`Class not registered`. Both are expected on a healthy install. The mail model
hangs off `GetNamespace("MAPI")`.

## Verified end-to-end (test message)

An email sent from a second account to the monitored one (subject
`COM watch test <date>`) landed in the monitored account's Inbox:

- Inbox count moved N → N+1; the message sat at the last `Items` index
- `ReceivedTime` was tz-aware (`+00:00`), body matched what was sent
- `read --limit 3` surfaced it first; `watch` printed it once, then reported
  `0 new` on the next cycle
- after the Sort port: `read --limit 3` 0.57 s, `read --limit 50` 0.58 s,
  `read --unread-only --limit 3` 0.19 s, `read --since 2026-10-05` 0.22 s

## Background fetch and resume

`--detach` re-runs the same command detached and returns; `start_new_session=True`
is what keeps it alive after the terminal closes (without it the child is
killed when the parent shell exits). Output appends to `.var/watch-<folder>.log`:

```bash
uv run python -m monitor watch --limit 50 --interval 300 --detach
# detached pid=<pid>; log=.var/watch-Inbox.log
```

`state.json` then carries a `runs` block beside the `seen` list:

```json
"runs": {"Inbox": {
  "criteria": {"interval": 15, "limit": 3, "unread_only": false},
  "last_run_utc": "2026-01-01T00:00:00",
  "last_seen_utc": "2025-12-31T23:59:59"}}
```

`--resume` starts the next scan after `last_seen_utc`, so a restart does not
rescan the folder from the top. The `seen` list alone cannot carry that: it is
truncated at `MAX_SEEN_PER_FOLDER` (5000) to keep the file bounded, so the
newest date is stored separately and survives truncation. An explicit `--since`
still wins over `--resume`. Verified: the restarted watch printed
`resuming after <last_seen_utc>` and scanned 1 message, not 3.

## Jev-style decisions about each message

"Jev" is TypeSafe AI's **System One** model (released 2026-09-15): it does not
write prose. You send a `state` plus typed questions and it returns typed
answers with probabilities and a confidence — `choice` (pick an option),
`score` (a level on an ordered scale), `noul` (yes/no probability).
<https://docs.typesafe.ai/introduction>

`monitor/decide.py` speaks that wire protocol once (`POST /v1/systemone`) and
points at either side of it, which is why one client covers both:

| backend | where | cost | vetting |
|---|---|---|---|
| **Jev 1.13** (hosted) | `https://api.typesafe.ai` | $42 / 1B input tokens, **output free**, 64k tokens per request (32k for state + longest question) | official docs |
| **Von** (local, CPU pick) | `von serve` → `http://127.0.0.1:8000` | 395M ModernBERT encoder, one forward pass, Apache-2.0, **854 stars**; p50 0.34 s on CPU, 0.096 s measured under OpenVINO | <https://github.com/wfzyx/von> |
| Bekko System One | local | 17M / 68M / 400M family, MIT, Sentence-Transformers + ONNX exports, **31 stars** | <https://github.com/hotchpotch/bekko-system-one> |
| Verdict | local | 118M bi-encoder, ONNX int8 ~120 MB, no PyTorch needed, Apache-2.0 — but **3 stars**, and its latency/benchmark claims are self-reported | <https://github.com/Manavarya09/verdict> |

Heavier options exist (`jevos` 1B, `minojev` Qwen3-1.7B + 0.8M head, NeoHorse
4B); they are autoregressive backbones and slower on CPU than the encoders
above. The many "X is a Jev alternative" posts circulating right now are
AI-generated content farms — the table above is repos, licenses and star counts.

```bash
uv run python -m monitor read --limit 5 --classify \
  --base-url http://127.0.0.1:8000 --model von-latest
```

One request per message answers every question in `TRIAGE_QUESTIONS` at once
(`kind`, `urgency`, `attention`), so adding a question costs nothing. The
confidence gate (0.80, mirroring Von's documented default) splits answers into
`auto` — safe for software to act on — and `review`. Verified against a stub
endpoint: `kind` at 0.88 and `urgency` at 0.88 land in `auto`; `attention` at
0.55 lands in `review`.

Without a key or a local server the stage degrades instead of crashing:
`decision endpoint unreachable: HTTPError: HTTP Error 403: Forbidden`.

## What is not wired yet

- Sending, moving, labelling, or archiving mail (kept read-only on request),
  and acting on the decisions above: `decide.py` classifies but nothing
  consumes `auto`/`review` yet.
- A local System One model installed and running. `decide.py` speaks the
  protocol; `von serve` (or Jev with `TYPESAFE_API_KEY`) has to be installed
  by you before `--classify` returns answers.
- Multi-account monitoring: only the account whose folders appear in the COM
  tree is monitored; a second signed-in account can contribute nothing.

## Layout

```
outlook_access/
├── pyproject.toml
├── README.md
├── .var/           # state.json, watch-<folder>.log (gitignored)
└── monitor/
    ├── com.py      # COM bridge: connect, folders, newest_items (Sort), restrict_items
    ├── reader.py   # newest-first streaming read with limit/since early break
    ├── state.py    # EntryID dedup + runs block (criteria, last_run, last_seen)
    ├── decide.py   # System One client: /v1/systemone, TRIAGE_QUESTIONS, confidence gate
    ├── doctor.py   # diagnosis + --regen-stubs
    └── __main__.py # CLI, incl. --detach / --resume / --classify
```

## Shell notes for this machine

`ctx_batch_execute` routes through PowerShell (breaks bash heredocs and `for`
loops) — use the bash tool directly. `winreg`, `tasklist` and `reg query` need
`cmd //c` wrapping. Raw `curl` output is blocked; use `curl -s -o file URL`.
