"""
Permanent history storage for Center Matcher results.

Do backend hain:
  1. GitHub (separate data repo) -- Streamlit Cloud par PERMANENT, reboot ke
     baad bhi data rehta hai. Streamlit secrets me [github] section chahiye.
  2. Local folder `saved_runs/` -- fallback (sirf local PC / server par
     permanent; Streamlit Cloud par reboot pe mit jaata hai).

Har run ek .xlsx file hai, naam: <timestamp-IST>__<label>.xlsx
(koi index/database nahi, isliye kuch corrupt hone ka risk nahi).
"""
import base64
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import quote

import requests

IST = timezone(timedelta(hours=5, minutes=30))
LOCAL_DIR = Path(__file__).parent / "saved_runs"
API = "https://api.github.com"
TS_FORMAT = "%Y-%m-%d_%H-%M-%S"


def _safe_label(label):
    label = re.sub(r"[^A-Za-z0-9._+-]+", "_", str(label)).strip("_")
    return (label or "run")[:80]


def _make_name(label):
    return f"{datetime.now(IST).strftime(TS_FORMAT)}__{_safe_label(label)}.xlsx"


def _parse_name(name):
    stem = name[:-5] if name.lower().endswith(".xlsx") else name
    ts, _, label = stem.partition("__")
    try:
        saved_at = datetime.strptime(ts, TS_FORMAT).strftime("%d %b %Y, %I:%M %p")
    except ValueError:
        saved_at = ts
    return {"name": name, "label": label or stem, "saved_at": saved_at}


# ── GitHub helpers ─────────────────────────────────────────────
def _gh_headers(conf, raw=False):
    h = {
        "Authorization": f"Bearer {conf['token']}",
        "X-GitHub-Api-Version": "2022-11-28",
        "Accept": "application/vnd.github.raw+json" if raw else "application/vnd.github+json",
    }
    return h


def _gh_url(conf, name=""):
    path = conf.get("dir", "runs").strip("/")
    if name:
        path += "/" + quote(name)
    return f"{API}/repos/{conf['repo']}/contents/{path}"


# ── Public API ─────────────────────────────────────────────────
def backend_name(conf):
    return "github" if conf else "local"


def save_run(label, data, conf=None):
    """xlsx bytes ko permanently save karta hai. Saved file ka naam return karta hai."""
    name = _make_name(label)
    if conf:
        r = requests.put(
            _gh_url(conf, name),
            headers=_gh_headers(conf),
            json={
                "message": f"Save matcher run: {name}",
                "content": base64.b64encode(data).decode("ascii"),
                "branch": conf.get("branch", "main"),
            },
            timeout=60,
        )
        if r.status_code not in (200, 201):
            raise RuntimeError(f"GitHub save failed ({r.status_code}): {r.text[:200]}")
    else:
        LOCAL_DIR.mkdir(parents=True, exist_ok=True)
        (LOCAL_DIR / name).write_bytes(data)
    return name


def list_runs(conf=None):
    """Saved runs ki list, sabse naya pehle."""
    names = []
    if conf:
        r = requests.get(
            _gh_url(conf),
            headers=_gh_headers(conf),
            params={"ref": conf.get("branch", "main")},
            timeout=30,
        )
        if r.status_code == 404:          # folder abhi bana nahi = koi run nahi
            return []
        if r.status_code != 200:
            raise RuntimeError(f"GitHub list failed ({r.status_code}): {r.text[:200]}")
        names = [i["name"] for i in r.json()
                 if i.get("type") == "file" and i["name"].lower().endswith(".xlsx")]
    elif LOCAL_DIR.exists():
        names = [p.name for p in LOCAL_DIR.glob("*.xlsx")]
    return [_parse_name(n) for n in sorted(names, reverse=True)]


def load_run(name, conf=None):
    """Saved run ki xlsx bytes."""
    if conf:
        r = requests.get(
            _gh_url(conf, name),
            headers=_gh_headers(conf, raw=True),
            params={"ref": conf.get("branch", "main")},
            timeout=60,
        )
        if r.status_code != 200:
            raise RuntimeError(f"GitHub load failed ({r.status_code}): {r.text[:200]}")
        return r.content
    return (LOCAL_DIR / name).read_bytes()
