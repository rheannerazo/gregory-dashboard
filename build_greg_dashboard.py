#!/usr/bin/env python3
"""
Build a static, READ-ONLY, Greg-facing status dashboard.

Reads the same source data as tracker_dashboard.py (the xlsx master tracker +
tracker_state.json) but writes a single self-contained static HTML file to
gregory-dashboard/docs/index.html, safe to publish on GitHub Pages.

This script does NOT import tracker_dashboard.py (that module can start a
server as a side effect of being run). Instead the small set of data-loading
helpers it needs are copied below.

Visual design mirrors the private "Mission Control" page (build_mission_control.py)
exactly: light theme, page bg #f6f5f2, white cards, #e8e6e1 borders, #1f2430
text, #8a8f98 secondary, system-ui font, rounded 10-12px corners, chips/pills,
sticky white top bar, left sidebar with a logo row + nav + ROI-style progress
card, hash-routed views. Same layout patterns (sidebar + topbar + kanban +
calendar), but:

  - BLUE (#1A8CF0) / NAVY (#0A1730) brand accent, not orange.
  - Strictly READ-ONLY: no buttons, no dialogs, no fetch/POST, no /api
    anything. The only JS on the page is the hash-based view router (same
    pattern as the reference).
  - Department level only -- no team members, agent names, roster, squad, or
    internal ops detail. Never uses the phrase "Mission Control".
  - Every existing sanitization behavior is preserved byte-for-byte: safe(),
    hits_blocklist(), task_is_sensitive(), clean_client_text(),
    display_workstream(), and the waiting-on-greg-only rule for the
    needs-you strip.

Run:  python build_greg_dashboard.py
"""

import os
import re
import json
import shutil
import datetime
import calendar as calmod
import html as html_lib

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Startup_Architect_Master_Tracker.xlsx")
STATE = os.path.join(HERE, "tracker_state.json")
ASSETS = os.path.join(HERE, "assets")
MC_STATE_DIR = os.path.join(HERE, "mission-control", "state")
CALENDAR_JSON = os.path.join(MC_STATE_DIR, "calendar.json")
MC_TASKS_JSON = os.path.join(MC_STATE_DIR, "tasks.json")
MC_OUTBOX_JSON = os.path.join(MC_STATE_DIR, "outbox.json")
MC_SCOREBOARD_JSON = os.path.join(MC_STATE_DIR, "scoreboard.json")
MC_DELIVERABLES_JSON = os.path.join(MC_STATE_DIR, "deliverables.json")

OUT_ROOT = os.path.join(HERE, "gregory-dashboard", "docs")
OUT_ASSETS_MAPS = os.path.join(OUT_ROOT, "assets", "maps")
OUT_INDEX = os.path.join(OUT_ROOT, "index.html")

NAVY = "#0A1730"
BLUE = "#1A8CF0"
BG = "#f6f5f2"

STATUS_COLORS = {
    "to do": "#6b7280", "in progress": "#1A8CF0", "waiting on greg": "#d97706",
    "blocked": "#dc2626", "ongoing": "#0d9488", "done": "#16a34a",
}

SCOREBOARD = [
    {"key": "registrations", "label": "Lecture registrations", "target": 45},
    {"key": "attendance",    "label": "Lecture attendance",    "target": 30},
    {"key": "seats",         "label": "2-Day seats sold",      "target": 10},
    {"key": "sponsors",      "label": "Sponsors closed ($1k)", "target": 3},
]

MAPS = [
    ("image1.png", "Your Role with Greg",
     "One role, several components: LinkedIn, the content engine, the website, the funnel and sponsors are all part of this one Chief of Staff role."),
    ("image2.png", "The 8-Week Plan",
     "Phase 1 by workstream across 8 weeks, building toward the Week 6 lecture and the Week 7 intensive."),
    ("image3.png", "Week 1, Day by Day",
     "The daily content engine plus each day's build focus, Monday to Friday."),
    ("image4.png", "The Content Engine",
     "Create once, publish everywhere: one source piece repurposed to every owned audience and social channel."),
    ("image5.png", "Event Run-of-Show",
     "The single live-event flow: 30 min networking + merch, 45 min talk, 5 min break, 30 min book signing."),
]

# Sanitization blocklist -- never surface these in any rendered task text.
BLOCKLIST = [
    "password", "lastpass", "delphi", "login", "credential",
    "rotate", "api key", "handover",
]

# Department workstreams Gregory is allowed to see. Anything else (Ascend,
# Infrastructure, or any future internal-only workstream) is excluded
# entirely from mission-control-sourced content -- never rendered, never
# counted, never named.
GREG_WORKSTREAMS = {
    "events", "speaking", "outreach", "newsletter", "content", "website", "meetings",
}

# Internal agent names -- must never appear on this page. Attribute
# mission-control-sourced work to "the team" instead.
AGENT_NAMES = {
    "atlas", "forge", "vet", "scout", "scribe", "warden", "concierge", "jesamie",
    "planner", "builder", "checker",
}


def mc_workstream_allowed(ws):
    """True only for the Gregory-relevant department workstreams. This is the
    hard privacy-wall gate for every mission-control task/outbox/calendar item
    before it is ever rendered."""
    return (ws or "").strip().lower() in GREG_WORKSTREAMS


_AGENT_NAME_RE = re.compile(
    r"\b(" + "|".join(re.escape(n) for n in AGENT_NAMES) + r")\b", re.IGNORECASE)


def scrub_agent_names(text):
    """Replace any internal agent name with 'the team' and tidy up double
    spaces/articles left behind (e.g. 'with Jesamie' -> 'with the team').
    Applied to every mission-control-sourced string before render -- Jesamie
    is a real (human) name but is treated the same as the AI agent names for
    this page: department-level only, no individual attribution."""
    if not text:
        return ""
    t = _AGENT_NAME_RE.sub("the team", text)
    t = re.sub(r"\bthe team's\b", "the team's", t, flags=re.IGNORECASE)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def safe(text):
    """Return text unchanged unless it hits the sensitive-info blocklist, then ''."""
    if not text:
        return ""
    low = text.lower()
    for word in BLOCKLIST:
        if word in low:
            return ""
    return text


INTERNAL_PAREN_RE = re.compile(
    r"\s*\([^)]*(call|meeting note|Rheanne|assumption|note)\b[^)]*\)", re.IGNORECASE)


def clean_client_text(text):
    """Strip internal shorthand / raw notes out of any client-facing string.

    - Removes parentheticals that reference internal shorthand (call, meeting
      note, Rheanne, assumption, note).
    - Drops sentences that start with "Working assumption".
    - Collapses resulting double spaces and trailing punctuation/space.
    """
    if not text:
        return ""
    t = INTERNAL_PAREN_RE.sub("", text)
    sentences = t.split(". ")
    sentences = [s for s in sentences if not s.strip().lower().startswith("working assumption")]
    t = ". ".join(sentences)
    t = re.sub(r"\s{2,}", " ", t)
    t = re.sub(r"\s+([.,;:])", r"\1", t)
    t = t.strip()
    t = re.sub(r"[\s.]+$", "", t) if t.endswith(("..", ". .")) else t
    t = t.rstrip()
    return t


def hits_blocklist(text):
    low = (text or "").lower()
    return any(word in low for word in BLOCKLIST)


def task_is_sensitive(t):
    """True if the task's own name/workstream (not just note/description) touches
    something on the blocklist -- these are ops/security chores that shouldn't be
    named on a client-facing page at all, even though they still count toward totals."""
    return hits_blocklist(t.get("Task", "")) or hits_blocklist(t.get("Workstream", ""))


def display_workstream(ws):
    """Cosmetic relabel for a client-facing page (e.g. internal 'KPI' workstream name)."""
    if (ws or "").strip().lower() == "kpi":
        return "Performance tracking"
    return ws


# ---------------- copied from tracker_dashboard.py ----------------
def clean(v):
    if v is None:
        return ""
    return str(v).strip().replace("�", "-").replace("—", "-").replace("–", "-")


def slug(s):
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


def parse_due(s, year=2026):
    m = re.match(r"([A-Za-z]{3})\s+(\d{1,2})", (s or "").strip())
    if not m:
        return None
    months = {mo: i for i, mo in enumerate(
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}
    mon = months.get(m.group(1).title())
    if not mon:
        return None
    try:
        return datetime.date(year, mon, int(m.group(2)))
    except ValueError:
        return None


def _grab(wb, sheet, header_key, stop_prefix=None):
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    idx = None
    for i, r in enumerate(rows):
        if r and clean(r[0]) == header_key:
            idx = i
            break
    out = []
    if idx is None:
        return out
    hdr = [clean(c) for c in rows[idx]]
    for r in rows[idx + 1:]:
        if not r or not clean(r[0]):
            continue
        if stop_prefix and clean(r[0]).startswith(stop_prefix):
            break
        out.append({hdr[i]: clean(r[i]) for i in range(min(len(hdr), len(r)))})
    return out


def load_tracker():
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    tasks = _grab(wb, "Task List", "ID")
    daily = _grab(wb, "Week 1 Daily", "Day")
    return tasks, daily


def load_state():
    s = {}
    if os.path.exists(STATE):
        try:
            with open(STATE, "r", encoding="utf-8") as f:
                s = json.load(f)
        except Exception:
            s = {}
    s.setdefault("tasks", {})
    s.setdefault("funnel", {})
    s.setdefault("kpi", {})
    s.setdefault("content", [])
    s.setdefault("greg", {})
    s.setdefault("decisions", {})
    s.setdefault("drafts", {})
    s.setdefault("extra_tasks", [])
    return s


def merged_tasks():
    tasks, daily = load_tracker()
    s = load_state()
    tasks = list(tasks) + list(s.get("extra_tasks", []))
    for t in tasks:
        tid = t.get("ID", "")
        st = s["tasks"].get(tid, {})
        t["Status"] = st.get("status") or t.get("Status") or "To do"
        t["Note"] = st.get("note", "")
        t["Log"] = st.get("log", [])
    return tasks, daily, s


# ---------------- mission-control state (department live data) ----------------
# All loaders below are read-only and tolerate a missing/unreadable/malformed
# file by returning an empty result -- the page must degrade gracefully to
# current (tracker-only) behavior if mission-control/state/ isn't there.

def _load_json(path):
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError, ValueError):
        return None


def load_mc_tasks():
    """Mission-control tasks, filtered to Gregory-relevant workstreams only.

    Applies the full privacy wall: excludes Ascend/Infrastructure/any
    non-allowlisted workstream, excludes task_is_sensitive() hits, and
    sanitizes every rendered string with the same safe()/clean_client_text()
    used for tracker tasks. Never attributes anything to an agent name.
    """
    data = _load_json(MC_TASKS_JSON)
    if not data:
        return []
    raw_tasks = data.get("tasks") or []
    out = []
    for t in raw_tasks:
        ws_raw = t.get("workstream", "")
        if not mc_workstream_allowed(ws_raw):
            continue
        # Reuse the existing sensitive-task gate (checks Task/Workstream keys).
        if task_is_sensitive({"Task": t.get("title", ""), "Workstream": ws_raw}):
            continue
        title = scrub_agent_names(clean_client_text(safe(clean(t.get("title", "")))))
        if not title:
            continue
        desc = scrub_agent_names(clean_client_text(safe(clean(t.get("description", "")))))
        status = (t.get("status") or "").strip().lower()
        out.append({
            "id": clean(t.get("id", "")),
            "title": title,
            "description": desc,
            "workstream": display_workstream(ws_raw),
            "status": status,
            "updated": clean(t.get("updated", "")),
        })
    return out


def load_mc_outbox():
    """Mission-control outbox emails, sanitized: subjects only, never
    recipient addresses, never raw feed lines."""
    data = _load_json(MC_OUTBOX_JSON)
    if not data:
        return []
    raw = data.get("emails") or []
    out = []
    for e in raw:
        subject = scrub_agent_names(clean_client_text(safe(clean(e.get("subject", "")))))
        if not subject:
            continue
        out.append({
            "id": clean(e.get("id", "")),
            "subject": subject,
            "status": (e.get("status") or "").strip().lower(),
            "updated": clean(e.get("updated", "") or e.get("created", "")),
        })
    return out


def load_mc_scoreboard():
    """Mission-control scoreboard metrics. Returns [] if the file is missing
    or empty so callers can fall back to the old hardcoded scoreboard."""
    data = _load_json(MC_SCOREBOARD_JSON)
    if not data:
        return []
    metrics = data.get("metrics") or {}
    out = []
    for key, m in metrics.items():
        label = clean_client_text(safe(clean(m.get("label", ""))))
        if not label:
            continue
        try:
            value = float(m.get("value", 0) or 0)
        except (TypeError, ValueError):
            value = 0
        try:
            target = float(m.get("target", 0) or 0)
        except (TypeError, ValueError):
            target = 0
        out.append({"key": key, "label": label, "value": value, "target": target})
    return out


def _parse_dt(s):
    """Best-effort ISO8601 -> naive UTC-ish datetime.date, tolerant of a
    trailing Z/offset. Returns None on anything unparsable."""
    if not s:
        return None
    try:
        s2 = s.replace("Z", "+00:00")
        return datetime.datetime.fromisoformat(s2)
    except ValueError:
        return None


# ---------------- helpers ----------------
def esc(x):
    return html_lib.escape(str(x) if x is not None else "")


def friendly_due(due_str):
    d = parse_due(due_str)
    if d:
        return d.strftime("%A, %b ") + str(d.day)
    return due_str or "no fixed date"


# ---------------- Needs You Greg (existing logic, preserved byte-for-byte) ----------------
def is_greg_owner(owner):
    return "greg" in (owner or "").lower()


def _needs_you_items(tasks):
    items = []
    for t in tasks:
        if task_is_sensitive(t):
            continue
        st = (t.get("Status") or "").lower()
        if st == "waiting on greg":
            items.append(t)
    return items


def build_needs_you(tasks):
    items = _needs_you_items(tasks)
    if not items:
        return ""

    today = datetime.date(2026, 7, 4)
    cards = ""
    for t in items:
        title = clean_client_text(safe(clean(t.get("Task", ""))))
        desc = clean_client_text(safe(clean(t.get("Description", ""))))
        line = desc
        # If the real ask was redacted/empty, show a safe, plain, client-appropriate line
        # instead of leaving a bare title that reads as broken to Greg.
        if not line:
            if "access" in title.lower():
                line = "Share access to the remaining tools and accounts so we can finish setting them up."
            else:
                line = "A quick hand-off from you unblocks this one."
        due = t.get("Due", "")
        due_html = ""
        d = parse_due(due)
        if d and d >= today:
            due_html = f'<div class="ny-due">Due {esc(friendly_due(due))}</div>'
        cards += (
            f'<div class="needcard">'
            f'<div class="ny-title">{esc(title)}</div>'
            f'<div class="ny-line">{esc(line) if line else ""}</div>'
            f'{due_html}'
            f'</div>')
    return f"""
<section class="needs-section">
  <div class="needs-head">
    <span class="needs-flag">NEEDS YOU, GREG</span>
    <span class="needs-sub muted">The couple of things only you can move forward.</span>
  </div>
  <div class="needgrid">{cards}</div>
</section>
"""


# ---------------- Kanban board (In progress / Waiting on you / Recently completed) ----------------
BOARD_COLUMNS = [
    ("in_progress", "In progress", "#1A8CF0", ("in progress", "ongoing", "blocked")),
    ("waiting", "Waiting on you", "#d97706", ("waiting on greg",)),
    ("done", "Recently completed", "#16a34a", ("done",)),
]

COLUMN_CAP = 8

# mission-control status -> board bucket. "backlog" has no bucket here (not
# yet in progress, not waiting on Greg, not done) so it simply doesn't appear
# on this department-facing board.
MC_STATUS_BUCKET = {
    "in_progress": "in_progress",
    "assigned": "in_progress",
    "review": "in_progress",
    "blocked": "in_progress",
    "ongoing": "in_progress",
    "done": "done",
}


def _task_note_line(t):
    """A clean, client-safe one-line note if there's something worth showing."""
    note = clean_client_text(safe(clean(t.get("Note", ""))))
    if note:
        return note
    desc = clean_client_text(safe(clean(t.get("Description", ""))))
    return desc


def _mentions_greg_or_lisa(text):
    low = (text or "").lower()
    return "greg" in low or "gregory" in low or "lisa" in low


def _normalize_tracker_card(t):
    """Tracker (xlsx + tracker_state.json) task -> a common card shape."""
    title = clean_client_text(safe(clean(t.get("Task", ""))))
    if not title:
        return None
    ws = display_workstream(t.get("Workstream", "Other") or "Other")
    note = _task_note_line(t)
    status = (t.get("Status") or "").lower()
    if status == "waiting on greg":
        bucket = "waiting"
    elif status in ("in progress", "ongoing", "blocked"):
        bucket = "in_progress"
    elif status == "done":
        bucket = "done"
    else:
        return None
    return {
        "title": title, "workstream": ws, "note": note, "bucket": bucket,
        "updated": t.get("Updated", "") or t.get("Due", ""),
    }


def _normalize_mc_card(t):
    """Mission-control task (already workstream-filtered + sanitized by
    load_mc_tasks) -> a common card shape. needs_human only lands in the
    'waiting' bucket if the ask genuinely names Greg/Lisa -- otherwise it's
    department-internal follow-up, not something to put in front of him."""
    status = t.get("status", "")
    if status == "needs_human":
        haystack = f"{t.get('title', '')} {t.get('description', '')}"
        if not _mentions_greg_or_lisa(haystack):
            return None
        bucket = "waiting"
    else:
        bucket = MC_STATUS_BUCKET.get(status)
        if not bucket:
            return None
    return {
        "title": t.get("title", ""),
        "workstream": t.get("workstream", "Other"),
        "note": t.get("description", ""),
        "bucket": bucket,
        "updated": t.get("updated", ""),
    }


def build_board_card(card):
    title = esc(card["title"])
    ws = esc(card["workstream"])
    note = card.get("note", "")
    note_html = f'<div class="task-desc">{esc(note)}</div>' if note else ""
    return f"""
<div class="taskcard">
  <div class="task-title-row"><div class="task-title">{title}</div></div>
  <div class="tag-row"><span class="tag">{ws}</span></div>
  {note_html}
</div>
"""


def build_board(tasks, mc_tasks=None):
    """Merge tracker tasks + sanitized mission-control tasks into one board.
    mc_tasks must already be workstream-filtered/sanitized via load_mc_tasks()."""
    visible = [t for t in tasks if not task_is_sensitive(t)]
    cards = [c for c in (_normalize_tracker_card(t) for t in visible) if c]
    for t in (mc_tasks or []):
        c = _normalize_mc_card(t)
        if c:
            cards.append(c)

    counts = {}
    columns_html = ""
    for key, label, color, _statuses in BOARD_COLUMNS:
        col_cards = [c for c in cards if c["bucket"] == key]
        if key == "done":
            # Most recent 8, newest first.
            col_cards.sort(key=lambda c: c.get("updated", ""), reverse=True)
        total_n = len(col_cards)
        counts[key] = total_n
        shown = col_cards[:COLUMN_CAP]
        cards_html = "".join(build_board_card(c) for c in shown) or '<p class="col-empty">Nothing here right now.</p>'
        more_html = ""
        remaining = total_n - len(shown)
        if remaining > 0:
            more_html = f'<p class="col-more muted">+{remaining} more</p>'
        columns_html += f"""
    <div class="col">
      <div class="col-head">
        <span class="col-dot" style="background:{color}"></span>
        <span class="col-name">{esc(label).upper()}</span>
        <span class="col-count">{total_n}</span>
      </div>
      <div class="col-body">{cards_html}{more_html}</div>
    </div>
"""
    return f'<div class="board">{columns_html}</div>', counts


# ---------------- CALENDAR ----------------
KEY_DATES_CURATED = [
    {"date": None, "tbd": True, "title": "Know Your Phase masterclass",
     "where": "The Bornemann Theatre, San Marcos -- early August 2026"},
    {"date": None, "tbd": True, "title": "From Vision to Exit -- 2-day intensive",
     "where": "Following the masterclass"},
]


def load_calendar_events():
    """Read mission-control/state/calendar.json (public events only). Falls
    back to the curated key-dates list if the file is missing/unreadable."""
    if not os.path.exists(CALENDAR_JSON):
        return list(KEY_DATES_CURATED), False
    try:
        with open(CALENDAR_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return list(KEY_DATES_CURATED), False

    events = data.get("events") or []
    out = []
    for e in events:
        if not e.get("public"):
            continue
        title = clean_client_text(safe(clean(e.get("title", ""))))
        where = clean_client_text(safe(clean(e.get("where", ""))))
        if not title:
            continue
        date_str = clean(e.get("date", ""))
        d = None
        if date_str:
            try:
                d = datetime.date.fromisoformat(date_str)
            except ValueError:
                d = None
        end_str = clean(e.get("end", ""))
        end_d = None
        if end_str:
            try:
                end_d = datetime.date.fromisoformat(end_str)
            except ValueError:
                end_d = None
        out.append({
            "date": d,
            "end": end_d,
            "tbd": bool(e.get("tbd")),
            "title": title,
            "where": where,
        })
    if not out:
        return list(KEY_DATES_CURATED), False
    return out, True


def upcoming_calendar_events(today=None):
    """Sorted upcoming/TBD public events -- shared by the sidebar mini-list
    and the full Calendar view."""
    today = today or datetime.date.today()
    events, _ = load_calendar_events()

    def sort_key(e):
        d = e.get("date")
        if d is None:
            return (1, datetime.date.max)
        return (0, d)

    upcoming = [e for e in events if e.get("tbd") or (e.get("date") and e["date"] >= today)]
    upcoming.sort(key=sort_key)
    return upcoming


def build_sidebar_key_dates(today=None):
    """Compact mini-list of the next 3 upcoming public events, for the sidebar."""
    upcoming = upcoming_calendar_events(today)[:3]
    if not upcoming:
        return ""

    rows = ""
    for e in upcoming:
        d = e.get("date")
        if e.get("tbd") or not d:
            badge = "TBD"
        else:
            badge = d.strftime("%b %d")
        rows += (
            f'<div class="kd-row">'
            f'<span class="kd-badge">{esc(badge)}</span>'
            f'<span class="kd-title">{esc(e.get("title", ""))}</span>'
            f'</div>')

    return f"""
    <div class="kd-card">
      <div class="kd-head">KEY DATES</div>
      <div class="kd-list">{rows}</div>
    </div>
"""


# ---------------- THIS WEEK (department activity summary) ----------------
def _this_week_calendar_events():
    """Calendar events fit to show in a Greg-facing weekly summary: public
    events, plus non-public events on the Meetings workstream (meeting titles
    are fine per the privacy rules -- only Ascend/Infrastructure and sensitive
    items are excluded)."""
    data = _load_json(CALENDAR_JSON)
    if not data:
        return []
    raw = data.get("events") or []
    out = []
    for e in raw:
        ws = e.get("workstream", "")
        is_public = bool(e.get("public"))
        is_meeting = (ws or "").strip().lower() == "meetings"
        if not (is_public or is_meeting):
            continue
        if ws and not mc_workstream_allowed(ws) and not is_public:
            # Non-public + not an allowed workstream -> skip.
            continue
        title = scrub_agent_names(clean_client_text(safe(clean(e.get("title", "")))))
        if not title:
            continue
        out.append(e | {"_title": title})
    return out


def build_this_week(mc_tasks, today=None):
    """Compact 'Your team this week' summary, last 7 days, derived only from
    structured state (tasks.json / outbox.json / calendar.json) -- never from
    feed.jsonl. Returns '' if there's nothing to show (e.g. mission-control
    state is entirely missing) so the section simply doesn't render."""
    today = today or datetime.date(2026, 7, 6)
    window_start = today - datetime.timedelta(days=7)

    def in_window(dt_str):
        dt = _parse_dt(dt_str)
        if not dt:
            return False
        d = dt.date() if hasattr(dt, "date") else dt
        return window_start <= d <= today

    # 1. Tasks completed this week (mission-control, already workstream-filtered).
    completed = [t for t in mc_tasks if t.get("status") == "done" and in_window(t.get("updated"))]
    completed.sort(key=lambda t: t.get("updated", ""), reverse=True)
    completed_titles = [t["title"] for t in completed[:5]]

    # 2. Outbox emails sent this week (subjects only).
    outbox = load_mc_outbox()
    sent = [e for e in outbox if e.get("status") == "sent" and in_window(e.get("updated"))]

    # 3. Meetings/events booked this week: calendar entries whose *creation*
    # falls in the window. calendar.json doesn't track a separate "added" ts,
    # so we use the event's own date as a proxy for recency isn't right --
    # instead we treat entries dated within the trailing week as newly on the
    # calendar. This intentionally undercounts rather than guesses.
    events = _this_week_calendar_events()
    booked = [e for e in events if in_window(e.get("date", ""))]

    # 4. Next upcoming date across the same event set.
    upcoming_dates = []
    for e in events:
        d = clean(e.get("date", ""))
        try:
            dd = datetime.date.fromisoformat(d)
        except ValueError:
            continue
        if dd >= today:
            upcoming_dates.append((dd, e["_title"]))
    upcoming_dates.sort(key=lambda x: x[0])
    next_up = upcoming_dates[0] if upcoming_dates else None

    if not (completed or sent or booked or next_up):
        return ""

    completed_html = ""
    if completed_titles:
        items = "".join(f'<li>{esc(t)}</li>' for t in completed_titles)
        completed_html = f'<ul class="tw-list">{items}</ul>'

    next_html = ""
    if next_up:
        next_html = f'<div class="tw-next"><span class="tw-next-label">Next up:</span> {esc(next_up[1])} &mdash; {esc(next_up[0].strftime("%b %d"))}</div>'

    return f"""
<section class="tw-section">
  <div class="tw-head">
    <span class="tw-flag">YOUR TEAM THIS WEEK</span>
    <span class="tw-sub muted">The last 7 days, from your Chief of Staff department.</span>
  </div>
  <div class="tw-grid">
    <div class="tw-card">
      <div class="tw-num">{len(completed)}</div>
      <div class="tw-label">Tasks completed</div>
      {completed_html}
    </div>
    <div class="tw-card">
      <div class="tw-num">{len(sent)}</div>
      <div class="tw-label">Replies &amp; emails sent</div>
    </div>
    <div class="tw-card">
      <div class="tw-num">{len(booked)}</div>
      <div class="tw-label">Meetings booked</div>
    </div>
  </div>
  {next_html}
</section>
"""


def build_calendar(today=None):
    today = today or datetime.date.today()
    upcoming = upcoming_calendar_events(today)
    events, _ = load_calendar_events()

    # Mini-grids: current month + next 2.
    event_days = {}  # (year, month) -> set of day numbers
    for e in events:
        d = e.get("date")
        if not d:
            continue
        event_days.setdefault((d.year, d.month), set()).add(d.day)
        end_d = e.get("end")
        if end_d and end_d >= d:
            cur = d
            while cur <= end_d:
                event_days.setdefault((cur.year, cur.month), set()).add(cur.day)
                cur += datetime.timedelta(days=1)

    months_html = ""
    y, m = today.year, today.month
    for _ in range(3):
        cal = calmod.Calendar(firstweekday=6)  # Sunday first
        weeks = cal.monthdayscalendar(y, m)
        mdays = event_days.get((y, m), set())
        month_name = datetime.date(y, m, 1).strftime("%B %Y")
        dow = "".join(f'<div class="cal-dow">{d}</div>' for d in ("S", "M", "T", "W", "T", "F", "S"))
        weeks_html = ""
        for week in weeks:
            for day in week:
                if day == 0:
                    weeks_html += '<div class="cal-day empty"></div>'
                    continue
                cls = "cal-day"
                if day in mdays:
                    cls += " has-event"
                if (y, m, day) == (today.year, today.month, today.day):
                    cls += " is-today"
                weeks_html += f'<div class="{cls}"><span>{day}</span></div>'
        months_html += (
            f'<div class="cal-month">'
            f'<div class="cal-month-title">{esc(month_name)}</div>'
            f'<div class="cal-grid">{dow}{weeks_html}</div>'
            f'</div>')
        # advance to next month
        if m == 12:
            y, m = y + 1, 1
        else:
            m += 1

    rows = ""
    for e in upcoming:
        d = e.get("date")
        end_d = e.get("end")
        if e.get("tbd") or not d:
            badge = '<div class="cal-badge cal-badge-tbd">date TBD</div>'
        elif end_d and end_d != d:
            badge = f'<div class="cal-badge">{d.strftime("%b %d")}&ndash;{end_d.strftime("%d")}</div>'
        else:
            badge = f'<div class="cal-badge">{d.strftime("%b %d")}</div>'
        where_html = f'<div class="cal-where muted">{esc(e.get("where", ""))}</div>' if e.get("where") else ""
        rows += (
            f'<div class="cal-row">'
            f'{badge}'
            f'<div class="cal-info"><div class="cal-title">{esc(e.get("title", ""))}</div>{where_html}</div>'
            f'</div>')

    if not rows:
        rows = '<p class="col-empty">Nothing scheduled right now.</p>'

    return f"""
<div class="cal-body">
  <div class="cal-months">{months_html}</div>
  <div class="cal-upcoming">
    <div class="cal-upcoming-head">Upcoming</div>
    <div class="cal-upcoming-list">{rows}</div>
  </div>
</div>
"""


def build_scoreboard(s):
    """Prefer mission-control/state/scoreboard.json (department-maintained
    metrics); fall back to the old hardcoded funnel scoreboard if that file
    is missing, empty, or unreadable."""
    mc_metrics = load_mc_scoreboard()

    out = ""
    if mc_metrics:
        for m in mc_metrics:
            target = m["target"]
            pct = min(100, (m["value"] / target * 100)) if target else 0
            actual = m["value"]
            actual_disp = int(actual) if actual == int(actual) else actual
            target_disp = int(target) if target == int(target) else target
            out += (
                f'<div class="score">'
                f'<div class="score-label">{esc(m["label"])}</div>'
                f'<div class="score-num">{esc(actual_disp)} <span class="muted">/ {target_disp}</span></div>'
                f'<div class="progress"><div style="width:{pct:.0f}%"></div></div>'
                f'</div>')
        return out

    for m in SCOREBOARD:
        actual = s["funnel"].get(m["key"], 0) or 0
        try:
            actual = float(actual)
        except (TypeError, ValueError):
            actual = 0
        pct = min(100, (actual / m["target"] * 100)) if m["target"] else 0
        actual_disp = int(actual) if actual == int(actual) else actual
        out += (
            f'<div class="score">'
            f'<div class="score-label">{esc(m["label"])}</div>'
            f'<div class="score-num">{esc(actual_disp)} <span class="muted">/ {m["target"]}</span></div>'
            f'<div class="progress"><div style="width:{pct:.0f}%"></div></div>'
            f'</div>')
    return out


def build_maps():
    out = ""
    for fname, title, desc in MAPS:
        out += (
            f'<div class="mapcard">'
            f'<h3>{esc(title)}</h3>'
            f'<img src="assets/maps/{fname}" alt="{esc(title)}" loading="lazy">'
            f'<p class="muted">{esc(desc)}</p>'
            f'</div>')
    return out


# ---------------- shared nav (desktop sidebar list + mobile chip row) ----------------
def build_nav(extra_class=""):
    cls = ("nav " + extra_class).strip()
    return f"""
    <nav class="{cls}">
      <a class="navlink" data-view-link="overview" href="#overview"><span class="navicon">\U0001F4CB</span>Overview</a>
      <a class="navlink" data-view-link="calendar" href="#calendar"><span class="navicon">\U0001F4C5</span>Calendar</a>
      <a class="navlink" data-view-link="scoreboard" href="#scoreboard"><span class="navicon">\U0001F3AF</span>Scoreboard</a>
      <a class="navlink" data-view-link="roadmap" href="#roadmap"><span class="navicon">\U0001F5FA</span>Roadmap</a>
    </nav>
"""


def build_sidebar(pct, key_dates_html):
    nav = build_nav(extra_class="nav-desktop")
    progress = f"""
    <div class="roi-card" id="progress">
      <div class="roi-label">OVERALL PROGRESS</div>
      <div class="roi-amount">{pct:.0f}%</div>
      <div class="roi-bar"><div class="roi-bar-fill" style="width:{pct:.0f}%"></div></div>
      <div class="roi-sub">prepared by your Chief of Staff department</div>
    </div>
"""
    return f"""
<aside class="sidebar">
  <div class="logo-row">
    <div class="logo-mark">GS</div>
    <div class="logo-text">
      <div class="logo-title">The Startup Architect</div>
      <div class="logo-sub">status dashboard</div>
    </div>
  </div>
  {nav}
  {progress}
  {key_dates_html}
</aside>
"""


# ---------------- page template ----------------
PAGE = """<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<title>Gregory Shepard &mdash; The Startup Architect &mdash; Status</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<style>
:root{{
  --bg:#f6f5f2;
  --panel:#ffffff;
  --border:#e8e6e1;
  --text:#1f2430;
  --muted:#8a8f98;
  --navy:{NAVY};
  --blue:{BLUE};
  --blue-bg:#e8f2fe;
  --amber:#d97706;
  --amber-bg:#fff7e6;
  --amber-border:#f0c674;
  --green:#16a34a;
}}
*{{box-sizing:border-box;}}
body{{margin:0;font-family:system-ui,-apple-system,"Segoe UI",Inter,sans-serif;color:var(--text);background:var(--bg);-webkit-font-smoothing:antialiased;}}
a{{color:var(--blue);}}
.muted{{color:var(--muted);font-size:.92em;}}
.pill{{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;padding:5px 12px;border-radius:999px;}}
.pill-ontrack{{color:var(--blue);border:1px solid var(--blue);background:#fff;}}

/* Top bar */
.topbar{{background:#fff;border-bottom:1px solid var(--border);position:sticky;top:0;z-index:10;}}
.topbar-inner{{max-width:1440px;margin:0 auto;padding:14px 24px;display:flex;align-items:center;gap:24px;flex-wrap:wrap;}}
.tb-wordmark{{font-weight:800;font-size:14.5px;color:var(--navy);white-space:nowrap;}}
.tb-stats{{display:flex;gap:26px;}}
.tb-stat{{line-height:1.2;}}
.tb-num{{font-size:20px;font-weight:800;color:var(--text);}}
.tb-label{{font-size:9.5px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);}}
.tb-center{{flex:1;display:flex;justify-content:center;}}
.tb-right{{margin-left:auto;text-align:right;}}
.tb-updated{{font-size:12px;color:var(--muted);white-space:nowrap;}}

/* Layout -- sidebar persists; only the active [data-view] section shows. */
.layout{{
  max-width:1440px;margin:0 auto;padding:20px 24px 60px;
  display:grid;
  grid-template-columns:230px minmax(0,1fr);
  grid-template-areas:"sidebar center";
  gap:20px;align-items:start;
}}
.layout-sidebar{{grid-area:sidebar;min-width:0;}}
.layout-center{{grid-area:center;min-width:0;}}
html.js-ready [data-view]{{display:none;}}
html.js-ready [data-view].view-active{{display:block;}}

/* Sidebar */
.sidebar{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:18px;position:sticky;top:76px;}}
.logo-row{{display:flex;align-items:center;gap:10px;margin-bottom:18px;}}
.logo-mark{{width:34px;height:34px;border-radius:9px;background:var(--blue);color:#fff;font-weight:800;font-size:13px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}}
.logo-title{{font-weight:800;font-size:13.5px;line-height:1.2;}}
.logo-sub{{font-size:10.5px;color:var(--muted);}}

.nav{{display:flex;flex-direction:column;gap:2px;margin-bottom:16px;}}
.navlink{{display:flex;align-items:center;gap:9px;font-size:13px;font-weight:600;color:var(--text);text-decoration:none;padding:8px 10px;border-radius:8px;}}
.navlink:hover{{background:#f7f6f3;}}
.navlink.active{{background:var(--blue-bg);color:var(--blue);}}
.navicon{{font-size:14px;width:16px;text-align:center;}}
.nav-mobile{{display:none;}}

.roi-card{{background:var(--blue-bg);border-radius:10px;padding:14px 14px 12px;margin-bottom:18px;}}
.roi-label{{font-size:10px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--blue);margin-bottom:4px;}}
.roi-amount{{font-size:26px;font-weight:800;color:var(--text);line-height:1.15;margin-bottom:8px;}}
.roi-bar{{height:5px;border-radius:3px;background:#cfe6fb;overflow:hidden;margin-bottom:8px;}}
.roi-bar-fill{{height:100%;background:var(--blue);border-radius:3px;}}
.roi-sub{{font-size:10.5px;color:var(--muted);font-style:italic;line-height:1.4;}}

.kd-card{{background:#f7f6f3;border:1px solid var(--border);border-radius:10px;padding:12px 14px;}}
.kd-head{{font-size:10px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin-bottom:10px;}}
.kd-list{{display:flex;flex-direction:column;gap:9px;}}
.kd-row{{display:flex;align-items:flex-start;gap:8px;}}
.kd-badge{{flex-shrink:0;font-size:9.5px;font-weight:800;color:var(--blue);background:#fff;border-radius:5px;padding:3px 7px;white-space:nowrap;}}
.kd-title{{font-size:11.5px;color:#3a3f4a;line-height:1.35;}}

/* Center panel */
.center{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:18px;}}
h2.view-title{{font-size:12px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;margin:0 0 4px;color:var(--text);}}
.section-sub{{color:var(--muted);font-size:12.5px;margin:0 0 16px;}}

/* This week */
.tw-section{{background:#f7f6f3;border:1px solid var(--border);border-radius:12px;padding:16px 18px 18px;margin-bottom:20px;}}
.tw-head{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;}}
.tw-flag{{font-size:11px;font-weight:800;letter-spacing:.6px;color:var(--navy);}}
.tw-sub{{font-size:12px;}}
.tw-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;}}
.tw-card{{background:#fff;border:1px solid var(--border);border-radius:10px;padding:14px 16px;}}
.tw-num{{font-size:24px;font-weight:800;color:var(--blue);line-height:1.1;}}
.tw-label{{font-size:12px;color:var(--navy);font-weight:700;margin-top:4px;}}
.tw-list{{margin:8px 0 0;padding-left:16px;font-size:11.5px;color:#5b606b;line-height:1.5;}}
.tw-next{{margin-top:12px;font-size:12.5px;color:#3a3f4a;}}
.tw-next-label{{font-weight:700;color:var(--navy);}}

/* Needs you */
.needs-section{{background:var(--amber-bg);border:1px solid var(--amber-border);border-radius:12px;padding:16px 18px 18px;margin-bottom:20px;}}
.needs-head{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;}}
.needs-flag{{font-size:11px;font-weight:800;letter-spacing:.6px;color:var(--amber);}}
.needs-sub{{font-size:12px;}}
.needgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;}}
.needcard{{background:#fff;border:1px solid var(--amber-border);border-left:3px solid var(--amber);border-radius:10px;padding:14px 16px;}}
.ny-title{{font-weight:700;color:var(--navy);font-size:14px;margin-bottom:6px;}}
.ny-line{{font-size:13px;color:#3a3f4a;line-height:1.5;}}
.ny-due{{font-size:11.5px;color:var(--amber);font-weight:700;margin-top:8px;}}

/* Kanban board */
.board{{display:grid;grid-template-columns:repeat(3,minmax(200px,1fr));gap:14px;}}
.col{{background:#faf9f7;border-radius:10px;padding:10px;min-width:0;}}
.col-head{{display:flex;align-items:center;gap:6px;font-size:11px;font-weight:800;letter-spacing:.4px;padding:4px 4px 10px;}}
.col-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0;}}
.col-name{{flex:1;}}
.col-count{{color:var(--muted);font-weight:700;}}
.col-body{{display:flex;flex-direction:column;gap:10px;}}
.col-empty{{padding:6px 4px;font-size:11.5px;color:var(--muted);}}
.col-more{{padding:2px 4px;font-size:11px;}}

.taskcard{{background:#fff;border:1px solid var(--border);border-radius:10px;padding:12px 13px;}}
.task-title-row{{margin-bottom:6px;}}
.task-title{{font-weight:700;font-size:13.5px;line-height:1.3;color:var(--navy);}}
.task-desc{{font-size:12px;color:#5b606b;line-height:1.4;margin-top:6px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;}}
.tag-row{{display:flex;flex-wrap:wrap;gap:5px;}}
.tag{{background:#f3f2ee;color:#5b606b;font-size:10px;border-radius:5px;padding:2px 7px;}}

/* Scoreboard */
.scoregrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px;}}
.score{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}}
.score-label{{font-size:12.5px;color:var(--navy);font-weight:700;}}
.score-num{{font-size:25px;font-weight:800;color:var(--blue);margin:8px 0;}}
.progress{{background:#eeece6;border-radius:20px;height:8px;overflow:hidden;}}
.progress>div{{background:var(--blue);height:100%;}}

/* Calendar */
.cal-body{{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(0,1fr);gap:24px;align-items:start;}}
.cal-months{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}}
.cal-month{{background:#faf9f7;border-radius:10px;padding:10px;}}
.cal-month-title{{font-size:11.5px;font-weight:800;color:var(--navy);margin-bottom:8px;text-align:center;}}
.cal-grid{{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}}
.cal-dow{{font-size:9px;font-weight:700;color:var(--muted);text-align:center;padding-bottom:3px;}}
.cal-day{{position:relative;text-align:center;font-size:10.5px;color:var(--text);padding:4px 0;border-radius:5px;}}
.cal-day span{{position:relative;z-index:1;}}
.cal-day.empty{{visibility:hidden;}}
.cal-day.is-today{{background:var(--navy);color:#fff;font-weight:800;}}
.cal-day.has-event:not(.is-today){{background:var(--blue-bg);font-weight:700;color:var(--blue);}}
.cal-day.has-event::after{{content:"";position:absolute;bottom:1px;left:50%;transform:translateX(-50%);width:3px;height:3px;border-radius:50%;background:var(--blue);}}
.cal-day.is-today.has-event::after{{background:#fff;}}
.cal-upcoming-head{{font-size:10.5px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin-bottom:12px;}}
.cal-upcoming-list{{display:flex;flex-direction:column;gap:12px;}}
.cal-row{{display:flex;gap:12px;align-items:flex-start;}}
.cal-badge{{flex-shrink:0;background:var(--blue-bg);color:var(--blue);font-weight:800;font-size:11px;border-radius:7px;padding:5px 9px;white-space:nowrap;}}
.cal-badge-tbd{{background:#f3f2ee;color:var(--muted);}}
.cal-title{{font-weight:700;font-size:13px;color:var(--navy);}}
.cal-where{{font-size:12px;margin-top:2px;}}

/* Maps / roadmap */
.maps{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;}}
.mapcard{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}}
.mapcard h3{{margin:0 0 10px;color:var(--navy);font-size:14.5px;}}
.mapcard img{{width:100%;border-radius:8px;border:1px solid var(--border);display:block;margin-bottom:10px;}}
.mapcard p{{margin:0;line-height:1.5;}}

.foot{{margin-top:32px;color:var(--muted);font-size:11.5px;text-align:center;}}

@media (max-width:999px){{
  .layout{{
    grid-template-columns:1fr;
    grid-template-areas:
      "center"
      "sidebar";
  }}
  .sidebar{{position:static;}}
  .cal-body{{grid-template-columns:1fr;}}
  .board{{grid-template-columns:repeat(2,minmax(0,1fr));}}
}}
@media (max-width:767px){{
  .layout{{padding:12px 12px 60px;gap:16px;}}
  .topbar-inner{{padding:10px 12px;gap:10px;}}
  .tb-stats{{gap:16px;}}
  .tb-num{{font-size:17px;}}
  .tb-label{{font-size:8.5px;}}

  .nav-desktop{{display:none;}}
  .nav-mobile{{
    display:flex;flex-direction:row;flex-wrap:nowrap;gap:8px;
    overflow-x:auto;-webkit-overflow-scrolling:touch;
    margin:0;padding:10px 12px;background:var(--panel);
    border-bottom:1px solid var(--border);
  }}
  .nav-mobile .navlink{{flex-shrink:0;padding:8px 12px;background:#f7f6f3;white-space:nowrap;min-height:36px;}}
  .nav-mobile .navlink.active{{background:var(--blue-bg);color:var(--blue);}}

  .board{{
    display:flex;
    grid-template-columns:unset;
    gap:12px;
    overflow-x:auto;
    scroll-snap-type:x mandatory;
    -webkit-overflow-scrolling:touch;
    margin:0 -4px;
    padding:0 4px 8px;
  }}
  .col{{
    flex:0 0 85vw;
    width:85vw;
    max-height:70vh;
    overflow-y:auto;
    scroll-snap-align:start;
  }}
  .col-head{{position:sticky;top:0;background:#faf9f7;z-index:1;}}

  .cal-months{{grid-template-columns:1fr;}}
}}
</style>
</head>
<body>
<header class="topbar">
  <div class="topbar-inner">
    <div class="tb-wordmark">Gregory Shepard &mdash; The Startup Architect</div>
    <div class="tb-stats">
      <div class="tb-stat"><div class="tb-num">{PCT:.0f}%</div><div class="tb-label">Overall progress</div></div>
      <div class="tb-stat"><div class="tb-num">{DONE}/{TOTAL}</div><div class="tb-label">Tasks done</div></div>
      <div class="tb-stat"><div class="tb-num">{WORKSTREAMS}</div><div class="tb-label">Active workstreams</div></div>
    </div>
    <div class="tb-center"><span class="pill pill-ontrack">ON TRACK</span></div>
    <div class="tb-right">
      <div class="tb-updated">Updated {DATE}</div>
    </div>
  </div>
</header>
{NAV_MOBILE}

<div class="layout">
  <div class="layout-sidebar">{SIDEBAR}</div>

  <section class="center layout-center" id="overview" data-view="overview">
    <h2 class="view-title">Overview</h2>
    <p class="section-sub">Every workstream your Chief of Staff department is running, at a glance.</p>
    {THIS_WEEK}
    {NEEDS_YOU}
    {BOARD}
  </section>

  <section class="center layout-center" id="calendar" data-view="calendar">
    <h2 class="view-title">Calendar</h2>
    <p class="section-sub">What's on the calendar for the launch.</p>
    {CALENDAR}
  </section>

  <section class="center layout-center" id="scoreboard" data-view="scoreboard">
    <h2 class="view-title">Scoreboard</h2>
    <p class="section-sub">Event funnel numbers so far, next to the goal.</p>
    <div class="scoregrid">{SCOREBOARD}</div>
  </section>

  <section class="center layout-center" id="roadmap" data-view="roadmap">
    <h2 class="view-title">Roadmap</h2>
    <p class="section-sub">The visual plan.</p>
    <div class="maps">{MAPS}</div>
  </section>
</div>

<div class="foot">Prepared and kept current by your Chief of Staff department.</div>

<script>
// ---- View router: hash-based nav (sidebar + top bar persist; only the
// center region swaps). Default view is #overview. No-JS fallback: the CSS
// only hides non-active [data-view] sections once <html> has the .js-ready
// class, which this script adds -- if JS never runs, every view just
// renders stacked. This is the only script on the page: pure client-side
// hash routing, no fetch, no POST, no external requests.
(function() {{
  var VALID_VIEWS = ['overview', 'calendar', 'scoreboard', 'roadmap'];

  function currentView() {{
    var h = (location.hash || '').replace('#', '');
    return VALID_VIEWS.indexOf(h) !== -1 ? h : 'overview';
  }}

  function showView(view) {{
    document.querySelectorAll('[data-view]').forEach(function(el) {{
      el.classList.toggle('view-active', el.getAttribute('data-view') === view);
    }});
    document.querySelectorAll('.navlink[data-view-link]').forEach(function(link) {{
      link.classList.toggle('active', link.getAttribute('data-view-link') === view);
    }});
  }}

  document.documentElement.classList.add('js-ready');
  window.addEventListener('hashchange', function() {{ showView(currentView()); }});
  showView(currentView());
}})();
</script>
</body></html>
"""


def render():
    tasks, daily, s = merged_tasks()
    total = len(tasks)
    done_n = sum(1 for t in tasks if (t.get("Status") or "").lower() == "done")
    pct = (done_n / total * 100) if total else 0

    mc_tasks = load_mc_tasks()

    workstream_count = len({display_workstream(t.get("Workstream", "Other") or "Other") for t in tasks}
                            | {t["workstream"] for t in mc_tasks})

    needs_you_html = build_needs_you(tasks)
    this_week_html = build_this_week(mc_tasks)
    board_html, board_counts = build_board(tasks, mc_tasks)
    calendar_html = build_calendar()
    scoreboard_html = build_scoreboard(s)
    maps_html = build_maps()
    key_dates_html = build_sidebar_key_dates()
    sidebar_html = build_sidebar(pct, key_dates_html)
    nav_mobile_html = f'<div class="layout-nav-mobile">{build_nav(extra_class="nav-mobile")}</div>'

    counts = {
        "needs_you": len(_needs_you_items(tasks)),
        "in_progress": board_counts.get("in_progress", 0),
        "waiting": board_counts.get("waiting", 0),
        "done": done_n,
        "total": total,
        "workstreams": workstream_count,
    }

    html = PAGE.format(
        NAVY=NAVY,
        BLUE=BLUE,
        DATE=datetime.date.today().strftime("%A, %B %d, %Y"),
        DONE=done_n,
        TOTAL=total,
        PCT=pct,
        WORKSTREAMS=workstream_count,
        SIDEBAR=sidebar_html,
        NAV_MOBILE=nav_mobile_html,
        THIS_WEEK=this_week_html,
        NEEDS_YOU=needs_you_html,
        BOARD=board_html,
        CALENDAR=calendar_html,
        SCOREBOARD=scoreboard_html,
        MAPS=maps_html,
    )
    return html, counts


def main():
    os.makedirs(OUT_ASSETS_MAPS, exist_ok=True)

    html, counts = render()
    with open(OUT_INDEX, "w", encoding="utf-8") as f:
        f.write(html)

    for fname, _, _ in MAPS:
        src = os.path.join(ASSETS, "maps", fname)
        dst = os.path.join(OUT_ASSETS_MAPS, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)
        else:
            print(f"WARNING: missing map asset {src}")

    print("Build complete.")
    print(f"  Needs you, Greg: {counts['needs_you']}")
    print(f"  In progress:     {counts['in_progress']}")
    print(f"  Waiting on you:  {counts['waiting']}")
    print(f"  Done:            {counts['done']}")
    print(f"  Total tasks:     {counts['total']}")
    print(f"  Workstreams:     {counts['workstreams']}")
    print(f"Output: {OUT_INDEX}")


if __name__ == "__main__":
    main()
