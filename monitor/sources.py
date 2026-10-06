"""External lookups: the catalog the model chooses from, and the fetch it triggers.

The contract each source script must honour (two commands):

  catalog  prints the option names the script accepts — one per line,
           optionally "name<TAB>description", or JSON (list / dict)
  fetch    takes the chosen name and prints the items — one per line,
           or JSON

Why the catalog is read from the script instead of written here: the model
answers with a key of `criteria`, and those keys are passed straight to your
script as an argument. If the list were hand-written it could drift, and the
model would emit a name the script rejects. Sourced from the script, the
option set is the script's own vocabulary by construction.

The model never produces the name out of thin air, and it never produces an
arbitrary string — `choice` returns one of the keys you gave it. That is what
makes the call safe to automate: an unprompted hallucination is impossible,
only a wrong pick, and a wrong pick is visible as low confidence.
"""

from __future__ import annotations

import json
import subprocess

CONFIG_PATH = "sources.json"

# sources.json shape:
# {
#   "cisco": {
#     "description": "Cisco firewall: services, blocks, network events",
#     "catalog": ["python", "fw.py", "--list"],
#     "fetch":   ["python", "fw.py", "--get", "{label}"],
#     "timeout": 30
#   }
# }
# Commands are argument lists, never shell strings: no quoting bugs, and a
# label the model picked cannot become a shell command.


def load_config(path: str = CONFIG_PATH) -> dict:
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError as exc:
        raise SystemExit(f"sources.json is malformed: {exc}")


def run(args: list[str], timeout: float = 30.0) -> str:
    proc = subprocess.run(args, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=timeout)
    if proc.returncode != 0:
        raise RuntimeError(
            f"{args[0]} exited {proc.returncode}: "
            f"{(proc.stderr or proc.stdout or '').strip()[:300]}")
    return proc.stdout


def parse(text: str) -> list | dict:
    """Accept either JSON or one-item-per-line output."""
    text = text.strip()
    if text[:1] in ("[", "{"):
        return json.loads(text)
    return [line for line in text.splitlines() if line.strip()]


def catalog(name: str, cfg: dict) -> dict:
    """Option names this source accepts, as {label: description}."""
    spec = cfg[name]
    args = spec["catalog"]
    if isinstance(args, str):
        args = args.split()
    out = parse(run(args, spec.get("timeout", 30.0)))
    if isinstance(out, dict):
        return dict(out)
    table = {}
    for line in out:
        if not isinstance(line, str):
            continue
        label, _, desc = line.partition("\t")
        table[label.strip()] = (desc or f"{name}: {label.strip()}").strip()
    return table


def catalogs(cfg: dict) -> dict:
    """Every source's options, namespaced `source:label` so picks are unique."""
    merged = {}
    for name in cfg:
        try:
            for label, desc in catalog(name, cfg).items():
                merged[f"{name}:{label}"] = desc
        except (RuntimeError, OSError, ValueError) as exc:
            print(f"catalog {name} unavailable: {type(exc).__name__}: {exc}")
    return merged


def fetch(cfg: dict, pick: str) -> list:
    """Run the chosen source with the chosen label; return its items."""
    name, _, label = pick.partition(":")
    spec = cfg[name]
    args = spec["fetch"]
    if isinstance(args, str):
        args = args.split()
    args = [a.replace("{label}", label) for a in args]
    return parse(run(args, spec.get("timeout", 30.0)))
