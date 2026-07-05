#!/usr/bin/env python3
"""
Build a static, READ-ONLY, Greg-facing status dashboard.

Reads the same source data as tracker_dashboard.py (the xlsx master tracker +
tracker_state.json) but writes a single self-contained static HTML file to
gregory-dashboard/docs/index.html, safe to publish on GitHub Pages.

This script does NOT import tracker_dashboard.py (that module can start a
server as a side effect of being run). Instead the small set of data-loading
helpers it needs are copied below.

Visual system matches the private "Mission Control" page (light theme: page
bg #f6f5f2, white cards, #e8e6e1 borders, #1f2430 text, #8a8f98 secondary,
system-ui font, rounded 10-12px, chips/pills, sticky white top bar) but with
the BLUE/NAVY brand accent instead of orange, and at department level only --
no team members, agent names, roster, or internal ops detail.

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
CALENDAR_JSON = os.path.join(HERE, "mission-control", "state", "calendar.json")

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


def build_needs_you(tasks):
    items = []
    seen_ids = set()
    for t in tasks:
        if task_is_sensitive(t):
            continue
        tid = t.get("ID", "")
        st = (t.get("Status") or "").lower()
        if st == "waiting on greg":
            items.append(t)
            seen_ids.add(tid)

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


# ---------------- THE PLAN (workstream overview cards) ----------------
def build_plan(tasks):
    by_ws = {}
    for t in tasks:
        ws = display_workstream(t.get("Workstream", "Other") or "Other")
        by_ws.setdefault(ws, []).append(t)

    def most_recent_focus(ts):
        in_motion = [t for t in ts if (t.get("Status") or "").lower() in ("in progress", "ongoing")
                     and not task_is_sensitive(t)]
        if not in_motion:
            return ""
        t = in_motion[-1]
        return clean_client_text(safe(clean(t.get("Task", ""))))

    cards = ""
    for ws, ts in sorted(by_ws.items()):
        done_n = sum(1 for t in ts if (t.get("Status") or "").lower() == "done")
        total_n = len(ts)
        pct = (done_n / total_n * 100) if total_n else 0
        focus = most_recent_focus(ts)
        focus_html = f'<div class="plan-focus">{esc(focus)}</div>' if focus else '<div class="plan-focus muted">Steady state -- no active build right now.</div>'
        cards += (
            f'<div class="plancard">'
            f'<div class="plan-name">{esc(ws)}</div>'
            f'<div class="plan-bar"><div style="width:{pct:.0f}%"></div></div>'
            f'<div class="plan-count muted">{done_n}/{total_n} done</div>'
            f'{focus_html}'
            f'</div>')
    return f'<div class="plangrid">{cards}</div>'


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
    return out, True


def build_calendar(today=None):
    today = today or datetime.date.today()
    events, from_json = load_calendar_events()

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

    # Upcoming list: future or TBD events, sorted by date (TBD/no-date last).
    def sort_key(e):
        d = e.get("date")
        if d is None:
            return (1, datetime.date.max)
        return (0, d)

    upcoming = [e for e in events if e.get("tbd") or (e.get("date") and e["date"] >= today)]
    upcoming.sort(key=sort_key)

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
        rows = '<p class="muted">Nothing scheduled right now.</p>'

    return f"""
<div class="cal-wrap">
  <div class="cal-months">{months_html}</div>
  <div class="cal-upcoming">
    <div class="cal-upcoming-title">Upcoming</div>
    <div class="cal-list">{rows}</div>
  </div>
</div>
"""


def build_scoreboard(s):
    out = ""
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
  --amber:#d97706;
  --amber-bg:#fff7e6;
  --amber-border:#f0c674;
  --green:#16a34a;
}}
*{{box-sizing:border-box;}}
body{{margin:0;font-family:system-ui,-apple-system,"Segoe UI",Inter,sans-serif;color:var(--text);background:var(--bg);-webkit-font-smoothing:antialiased;}}
a{{color:var(--blue);}}
.muted{{color:var(--muted);font-size:.92em;}}

/* Top bar */
.topbar{{background:#fff;border-bottom:1px solid var(--border);position:sticky;top:0;z-index:10;}}
.topbar-inner{{max-width:1180px;margin:0 auto;padding:16px 24px;display:flex;align-items:center;gap:24px;flex-wrap:wrap;}}
.tb-wordmark{{font-weight:800;font-size:15px;line-height:1.25;color:var(--navy);}}
.tb-wordmark .tb-sub{{display:block;font-weight:600;font-size:11px;color:var(--muted);margin-top:2px;}}
.tb-stats{{display:flex;gap:26px;margin-left:8px;}}
.tb-stat{{line-height:1.2;}}
.tb-num{{font-size:20px;font-weight:800;color:var(--text);}}
.tb-label{{font-size:9.5px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);}}
.tb-right{{margin-left:auto;text-align:right;}}
.tb-updated{{font-size:12px;color:var(--muted);}}
.tb-dept{{font-size:11px;color:var(--muted);font-style:italic;margin-top:2px;}}

.wrap{{max-width:1180px;margin:0 auto;padding:28px 24px 70px;}}

h2{{font-size:12px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;margin:0 0 4px;color:var(--text);}}
.section{{margin-top:40px;}}
.section:first-child{{margin-top:0;}}
.section-sub{{color:var(--muted);font-size:12.5px;margin:0 0 16px;}}

/* Needs you */
.needs-section{{background:var(--amber-bg);border:1px solid var(--amber-border);border-radius:12px;padding:16px 18px 18px;margin-bottom:8px;}}
.needs-head{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;}}
.needs-flag{{font-size:11px;font-weight:800;letter-spacing:.6px;color:var(--amber);}}
.needs-sub{{font-size:12px;}}
.needgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;}}
.needcard{{background:#fff;border:1px solid var(--amber-border);border-left:3px solid var(--amber);border-radius:10px;padding:14px 16px;}}
.ny-title{{font-weight:700;color:var(--navy);font-size:14px;margin-bottom:6px;}}
.ny-line{{font-size:13px;color:#3a3f4a;line-height:1.5;}}
.ny-due{{font-size:11.5px;color:var(--amber);font-weight:700;margin-top:8px;}}

/* Plan cards */
.plangrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px;}}
.plancard{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}}
.plan-name{{font-weight:700;color:var(--navy);font-size:14px;margin-bottom:10px;}}
.plan-bar{{background:#eeece6;border-radius:20px;height:8px;overflow:hidden;margin-bottom:6px;}}
.plan-bar>div{{background:var(--blue);height:100%;border-radius:20px;}}
.plan-count{{font-size:11.5px;margin-bottom:10px;}}
.plan-focus{{font-size:12.5px;color:#3a3f4a;line-height:1.4;border-top:1px solid var(--border);padding-top:10px;}}

/* Calendar */
.cal-wrap{{display:grid;grid-template-columns:1.4fr 1fr;gap:20px;align-items:start;}}
.cal-months{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px;}}
.cal-month-title{{font-size:11.5px;font-weight:800;color:var(--navy);margin-bottom:8px;text-align:center;}}
.cal-grid{{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}}
.cal-dow{{font-size:9px;font-weight:700;color:var(--muted);text-align:center;padding-bottom:3px;}}
.cal-day{{position:relative;text-align:center;font-size:10.5px;color:var(--text);padding:4px 0;border-radius:5px;}}
.cal-day span{{position:relative;z-index:1;}}
.cal-day.empty{{visibility:hidden;}}
.cal-day.is-today{{background:var(--navy);color:#fff;font-weight:800;}}
.cal-day.has-event:not(.is-today){{background:#e8f2fe;font-weight:700;color:var(--blue);}}
.cal-day.has-event::after{{content:"";position:absolute;bottom:1px;left:50%;transform:translateX(-50%);width:3px;height:3px;border-radius:50%;background:var(--blue);}}
.cal-day.is-today.has-event::after{{background:#fff;}}
.cal-upcoming{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}}
.cal-upcoming-title{{font-size:11px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin-bottom:12px;}}
.cal-list{{display:flex;flex-direction:column;gap:12px;}}
.cal-row{{display:flex;gap:12px;align-items:flex-start;}}
.cal-badge{{flex-shrink:0;background:#e8f2fe;color:var(--blue);font-weight:800;font-size:11px;border-radius:7px;padding:5px 9px;white-space:nowrap;}}
.cal-badge-tbd{{background:#f3f2ee;color:var(--muted);}}
.cal-title{{font-weight:700;font-size:13px;color:var(--navy);}}
.cal-where{{font-size:12px;margin-top:2px;}}

/* Scoreboard */
.scoregrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px;}}
.score{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}}
.score-label{{font-size:12.5px;color:var(--navy);font-weight:700;}}
.score-num{{font-size:25px;font-weight:800;color:var(--blue);margin:8px 0;}}
.progress{{background:#eeece6;border-radius:20px;height:8px;overflow:hidden;}}
.progress>div{{background:var(--blue);height:100%;}}

/* Maps */
.maps{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;}}
.mapcard{{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}}
.mapcard h3{{margin:0 0 10px;color:var(--navy);font-size:14.5px;}}
.mapcard img{{width:100%;border-radius:8px;border:1px solid var(--border);display:block;margin-bottom:10px;}}
.mapcard p{{margin:0;line-height:1.5;}}

.foot{{margin-top:56px;color:var(--muted);font-size:11.5px;text-align:center;}}

@media (max-width:900px){{
  .cal-wrap{{grid-template-columns:1fr;}}
  .cal-months{{grid-template-columns:1fr;}}
}}
@media (max-width:640px){{
  .wrap{{padding:22px 16px 56px;}}
  .topbar-inner{{padding:14px 16px;}}
  .tb-right{{margin-left:0;text-align:left;}}
}}
</style>
</head>
<body>
<header class="topbar">
  <div class="topbar-inner">
    <div class="tb-wordmark">Gregory Shepard &mdash; The Startup Architect<span class="tb-sub">Prepared by your Chief of Staff department</span></div>
    <div class="tb-stats">
      <div class="tb-stat"><div class="tb-num">{PCT:.0f}%</div><div class="tb-label">Overall progress</div></div>
      <div class="tb-stat"><div class="tb-num">{DONE}/{TOTAL}</div><div class="tb-label">Tasks done</div></div>
      <div class="tb-stat"><div class="tb-num">{WORKSTREAMS}</div><div class="tb-label">Active workstreams</div></div>
    </div>
    <div class="tb-right">
      <div class="tb-updated">Updated {DATE}</div>
    </div>
  </div>
</header>

<div class="wrap">

  {NEEDS_YOU}

  <div class="section">
    <h2>The Plan</h2>
    <p class="section-sub">Every workstream your Chief of Staff department is running, with current progress and focus.</p>
    {PLAN}
  </div>

  <div class="section">
    <h2>Calendar</h2>
    <p class="section-sub">What's on the calendar for the launch.</p>
    {CALENDAR}
  </div>

  <div class="section">
    <h2>Scoreboard</h2>
    <p class="section-sub">Event funnel numbers so far, next to the goal.</p>
    <div class="scoregrid">{SCOREBOARD}</div>
  </div>

  <div class="section">
    <h2>Maps &amp; Big Picture</h2>
    <p class="section-sub">The visual plan.</p>
    <div class="maps">{MAPS}</div>
  </div>

  <div class="foot">The Startup Architect &middot; status snapshot, updated periodically</div>
</div>
</body></html>
"""


def render():
    tasks, daily, s = merged_tasks()
    total = len(tasks)
    done_n = sum(1 for t in tasks if (t.get("Status") or "").lower() == "done")
    pct = (done_n / total * 100) if total else 0

    workstream_count = len({display_workstream(t.get("Workstream", "Other") or "Other") for t in tasks})

    needs_you_html = build_needs_you(tasks)
    plan_html = build_plan(tasks)
    calendar_html = build_calendar()
    scoreboard_html = build_scoreboard(s)
    maps_html = build_maps()

    counts = {
        "needs_you": len([t for t in tasks if (t.get("Status") or "").lower() == "waiting on greg"
                          and not task_is_sensitive(t)]),
        "in_progress": len([t for t in tasks if (t.get("Status") or "").lower() in ("in progress", "ongoing")]),
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
        NEEDS_YOU=needs_you_html,
        PLAN=plan_html,
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
    print(f"  Done:            {counts['done']}")
    print(f"  Total tasks:     {counts['total']}")
    print(f"  Workstreams:     {counts['workstreams']}")
    print(f"Output: {OUT_INDEX}")


if __name__ == "__main__":
    main()
