#!/usr/bin/env python3
"""
Build a static, READ-ONLY, Greg-facing status dashboard.

Reads the same source data as tracker_dashboard.py (the xlsx master tracker +
tracker_state.json) but writes a single self-contained static HTML file to
gregory-dashboard/docs/index.html, safe to publish on GitHub Pages.

This script does NOT import tracker_dashboard.py (that module can start a
server as a side effect of being run). Instead the small set of data-loading
helpers it needs are copied below.

Run:  python build_greg_dashboard.py
"""

import os
import re
import json
import shutil
import datetime
import html as html_lib

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Startup_Architect_Master_Tracker.xlsx")
STATE = os.path.join(HERE, "tracker_state.json")
ASSETS = os.path.join(HERE, "assets")

OUT_ROOT = os.path.join(HERE, "gregory-dashboard", "docs")
OUT_ASSETS_MAPS = os.path.join(OUT_ROOT, "assets", "maps")
OUT_INDEX = os.path.join(OUT_ROOT, "index.html")

NAVY = "#0A1730"
BLUE = "#1A8CF0"
BG = "#f5f7fa"

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


def badge(status):
    c = STATUS_COLORS.get((status or "").lower().strip(), "#6b7280")
    return f'<span class="badge" style="background:{c}">{esc(status or "-")}</span>'


def friendly_due(due_str):
    d = parse_due(due_str)
    if d:
        return d.strftime("%A, %b ") + str(d.day)
    return due_str or "no fixed date"


# ---------------- Greg buckets ----------------
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
    for t in tasks:
        if task_is_sensitive(t):
            continue
        tid = t.get("ID", "")
        if tid in seen_ids:
            continue
        st = (t.get("Status") or "").lower()
        if is_greg_owner(t.get("Owner", "")) and st != "done":
            items.append(t)
            seen_ids.add(tid)

    if not items:
        return '<div class="card"><p class="muted">Nothing is waiting on you right now.</p></div>'

    cards = ""
    for t in items:
        desc = safe(clean(t.get("Description", "")))
        note = safe(clean(t.get("Note", "")))
        line = desc or note
        extra = ""
        if note and desc and note != desc:
            extra = f'<div class="ny-note">{esc(note)}</div>'
        due = t.get("Due", "")
        due_html = f'<div class="ny-due">Due {esc(friendly_due(due))}</div>' if due else ""
        cards += (
            f'<div class="needcard">'
            f'<div class="ny-title">{esc(t.get("Task",""))}</div>'
            f'<div class="ny-line">{esc(line) if line else ""}</div>'
            f'{extra}{due_html}'
            f'</div>')
    return f'<div class="needgrid">{cards}</div>'


def build_in_progress(tasks):
    items = [t for t in tasks if (t.get("Status") or "").lower() in ("in progress", "ongoing")
             and not task_is_sensitive(t)]
    if not items:
        return '<div class="card"><p class="muted">Nothing actively in motion right now.</p></div>'
    rows = ""
    for t in items:
        desc = safe(clean(t.get("Description", "")))
        rows += (
            f'<div class="progitem">'
            f'<div class="pi-head"><span class="pi-title">{esc(t.get("Task",""))}</span>{badge(t.get("Status"))}</div>'
            f'<div class="pi-desc">{esc(desc)}</div>'
            f'<div class="pi-meta muted">{esc(display_workstream(t.get("Workstream","")))}</div>'
            f'</div>')
    return f'<div class="proggrid">{rows}</div>'


def build_done(tasks):
    items = [t for t in tasks if (t.get("Status") or "").lower() == "done" and not task_is_sensitive(t)]
    if not items:
        return '<div class="card"><p class="muted">No completed items yet.</p></div>'
    by_ws = {}
    for t in items:
        ws = display_workstream(t.get("Workstream", "Other") or "Other")
        by_ws.setdefault(ws, []).append(t)
    out = ""
    for ws, ts in sorted(by_ws.items()):
        chips = "".join(f'<li>{esc(t.get("Task",""))}</li>' for t in ts)
        out += f'<div class="donegroup"><h4>{esc(ws)} <span class="muted">({len(ts)})</span></h4><ul class="donelist">{chips}</ul></div>'
    return f'<div class="donegrid">{out}</div>'


def build_where_things_stand(tasks):
    by_ws = {}
    for t in tasks:
        ws = display_workstream(t.get("Workstream", "Other") or "Other")
        by_ws.setdefault(ws, []).append(t)

    def rollup(ts):
        statuses = [(t.get("Status") or "").lower() for t in ts]
        if any(s == "waiting on greg" for s in statuses):
            return "Waiting on you", "#d97706"
        if any(s in ("in progress", "ongoing") for s in statuses):
            return "Building", BLUE
        if all(s == "done" for s in statuses):
            return "Done", "#16a34a"
        return "Up next", "#6b7280"

    cards = ""
    for ws, ts in sorted(by_ws.items()):
        label, color = rollup(ts)
        done_n = sum(1 for t in ts if (t.get("Status") or "").lower() == "done")
        total_n = len(ts)
        cards += (
            f'<div class="wscard">'
            f'<div class="ws-name">{esc(ws)}</div>'
            f'<div class="ws-state" style="color:{color}">{label}</div>'
            f'<div class="ws-count muted">{done_n}/{total_n} done</div>'
            f'</div>')
    return f'<div class="wsgrid">{cards}</div>'


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


def build_key_dates(tasks):
    td = datetime.date.today()
    upcoming = []
    for t in tasks:
        if (t.get("Status") or "").lower() == "done":
            continue
        if task_is_sensitive(t):
            continue
        d = parse_due(t.get("Due", ""), year=2026)
        if d and d >= td:
            upcoming.append((d, t))
    upcoming.sort(key=lambda x: x[0])

    rows = ""
    for d, t in upcoming[:8]:
        rows += (
            f'<div class="daterow">'
            f'<div class="date-day">{d.strftime("%b %d")}</div>'
            f'<div class="date-task">{esc(t.get("Task",""))}</div>'
            f'</div>')
    if not rows:
        rows = '<p class="muted">No upcoming dated tasks right now.</p>'

    curated = (
        '<div class="daterow curated">'
        '<div class="date-day">TBD</div>'
        '<div class="date-task"><b>Know Your Phase masterclass</b> &middot; The Bornemann Theatre, San Marcos &middot; early August 2026 (date TBD)</div>'
        '</div>')
    return f'<div class="datelist">{rows}{curated}</div>'


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
<title>The Startup Architect &mdash; Status</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Anton&family=Montserrat:wght@500;600;700;800&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{{
  --navy:#0A1730;
  --blue:#1A8CF0;
  --bg:#f5f7fa;
  --text:#1f2937;
  --muted:#6b7280;
}}
*{{box-sizing:border-box;}}
body{{margin:0;font-family:'Inter',system-ui,sans-serif;color:var(--text);background:var(--bg);-webkit-font-smoothing:antialiased;}}
h1,h2,h3,h4{{font-family:'Montserrat',system-ui,sans-serif;}}

header{{background:var(--navy);color:#fff;padding:36px 24px 30px;}}
.header-inner{{max-width:1100px;margin:0 auto;}}
header .eyebrow{{color:var(--blue);font-weight:700;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin-bottom:8px;}}
header h1{{font-family:'Anton',Montserrat,sans-serif;font-weight:400;margin:0 0 8px;font-size:clamp(28px,5vw,42px);letter-spacing:.5px;}}
header .subtitle{{color:#cdd7e8;font-size:15px;margin:0 0 18px;max-width:640px;line-height:1.5;}}
header .asof{{color:#9fb0cc;font-size:12px;margin-bottom:14px;}}
.overall-wrap{{max-width:420px;}}
.overall-label{{display:flex;justify-content:space-between;font-size:12px;color:#cdd7e8;margin-bottom:6px;}}
.overall-bar{{background:rgba(255,255,255,.15);border-radius:20px;height:12px;overflow:hidden;}}
.overall-bar>div{{background:var(--blue);height:100%;border-radius:20px;}}

.wrap{{max-width:1100px;margin:0 auto;padding:32px 24px 80px;}}

h2{{color:var(--navy);font-size:22px;font-weight:800;margin:44px 0 6px;padding-bottom:10px;border-bottom:3px solid var(--blue);}}
h2:first-of-type{{margin-top:0;}}
.section-sub{{color:var(--muted);font-size:13px;margin:0 0 18px;}}

.card{{background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 2px 10px rgba(10,23,48,.06);}}
.muted{{color:var(--muted);font-size:.92em;}}

/* Needs you */
.needgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;}}
.needcard{{background:#fff;border-left:4px solid var(--blue);border-radius:12px;padding:16px 18px;box-shadow:0 2px 10px rgba(10,23,48,.06);}}
.ny-title{{font-weight:700;color:var(--navy);font-size:14.5px;margin-bottom:6px;}}
.ny-line{{font-size:13.5px;color:#374151;line-height:1.5;}}
.ny-note{{font-size:12.5px;color:var(--muted);margin-top:6px;font-style:italic;}}
.ny-due{{font-size:12px;color:var(--blue);font-weight:600;margin-top:8px;}}

/* In progress */
.proggrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;}}
.progitem{{background:#fff;border-radius:12px;padding:14px 18px;box-shadow:0 2px 10px rgba(10,23,48,.06);border-top:3px solid var(--blue);}}
.pi-head{{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px;}}
.pi-title{{font-weight:700;color:var(--navy);font-size:14px;}}
.pi-desc{{font-size:13px;color:#374151;line-height:1.45;}}
.pi-meta{{margin-top:8px;font-size:11.5px;text-transform:uppercase;letter-spacing:.5px;}}

/* Done */
.donegrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px;}}
.donegroup{{background:#fff;border-radius:12px;padding:14px 18px;box-shadow:0 2px 10px rgba(10,23,48,.06);}}
.donegroup h4{{margin:0 0 8px;color:var(--navy);font-size:13.5px;}}
.donelist{{list-style:none;margin:0;padding:0;}}
.donelist li{{font-size:13px;color:#374151;padding:5px 0 5px 22px;position:relative;border-top:1px solid #eef1f5;}}
.donelist li:first-child{{border-top:none;}}
.donelist li::before{{content:"\\2713";position:absolute;left:0;color:#16a34a;font-weight:700;}}

/* Where things stand */
.wsgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;}}
.wscard{{background:#fff;border-radius:12px;padding:16px;text-align:center;box-shadow:0 2px 10px rgba(10,23,48,.06);}}
.ws-name{{font-weight:700;color:var(--navy);font-size:13px;margin-bottom:8px;}}
.ws-state{{font-weight:800;font-size:15px;margin-bottom:4px;}}
.ws-count{{font-size:12px;}}

/* Scoreboard */
.scoregrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px;}}
.score{{background:#fff;border-radius:12px;padding:16px;box-shadow:0 2px 10px rgba(10,23,48,.06);border-top:4px solid var(--navy);}}
.score-label{{font-size:13px;color:var(--navy);font-weight:700;}}
.score-num{{font-size:26px;font-weight:800;color:var(--blue);margin:8px 0;}}
.progress{{background:#e5e7eb;border-radius:20px;height:10px;overflow:hidden;}}
.progress>div{{background:var(--blue);height:100%;}}

/* Key dates */
.datelist{{background:#fff;border-radius:12px;box-shadow:0 2px 10px rgba(10,23,48,.06);overflow:hidden;}}
.daterow{{display:flex;gap:16px;align-items:baseline;padding:12px 18px;border-top:1px solid #eef1f5;}}
.daterow:first-child{{border-top:none;}}
.date-day{{flex-shrink:0;width:70px;font-weight:800;color:var(--blue);font-size:13px;}}
.date-task{{font-size:13.5px;color:#374151;}}
.daterow.curated{{background:#f7fafe;}}

/* Maps */
.maps{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:18px;}}
.mapcard{{background:#fff;border-radius:14px;padding:16px 18px;box-shadow:0 2px 10px rgba(10,23,48,.06);}}
.mapcard h3{{margin:0 0 10px;color:var(--navy);font-size:15px;}}
.mapcard img{{width:100%;border-radius:8px;border:1px solid #eef1f5;display:block;margin-bottom:10px;}}
.mapcard p{{margin:0;line-height:1.5;}}

.foot{{margin-top:56px;color:#9ca3af;font-size:12px;text-align:center;}}

@media (max-width:640px){{
  .wrap{{padding:24px 16px 60px;}}
  header{{padding:28px 16px 24px;}}
  h2{{font-size:19px;}}
}}
</style>
</head>
<body>
<header>
  <div class="header-inner">
    <div class="eyebrow">The Startup Architect</div>
    <h1>Status Dashboard</h1>
    <p class="subtitle">A plain-language snapshot of where the launch stands: what's done, what's moving, and the couple of things only you can unblock.</p>
    <div class="asof">As of {DATE}</div>
    <div class="overall-wrap">
      <div class="overall-label"><span>Overall progress</span><span>{DONE}/{TOTAL} tasks done</span></div>
      <div class="overall-bar"><div style="width:{PCT:.0f}%"></div></div>
    </div>
  </div>
</header>

<div class="wrap">

  <h2 id="needs-you">Needs You, Greg</h2>
  <p class="section-sub">The couple of things only you can move forward. Everything else is being handled.</p>
  {NEEDS_YOU}

  <h2 id="in-progress">In Progress</h2>
  <p class="section-sub">Actively being worked on right now.</p>
  {IN_PROGRESS}

  <h2 id="where-things-stand">Where Things Stand</h2>
  <p class="section-sub">A one-word status for each area of the launch.</p>
  {WHERE_STAND}

  <h2 id="scoreboard">Scoreboard</h2>
  <p class="section-sub">Event funnel numbers so far, next to the goal.</p>
  <div class="scoregrid">{SCOREBOARD}</div>

  <h2 id="key-dates">Key Dates</h2>
  <p class="section-sub">What's coming up.</p>
  {KEY_DATES}

  <h2 id="done">Done</h2>
  <p class="section-sub">Wins so far, grouped by area.</p>
  {DONE_SECTION}

  <h2 id="maps">Maps &amp; Big Picture</h2>
  <p class="section-sub">The visual plan.</p>
  <div class="maps">{MAPS}</div>

  <div class="foot">The Startup Architect &middot; status snapshot, updated periodically</div>
</div>
</body></html>
"""


def render():
    tasks, daily, s = merged_tasks()
    total = len(tasks)
    done_n = sum(1 for t in tasks if (t.get("Status") or "").lower() == "done")
    pct = (done_n / total * 100) if total else 0

    needs_you_html = build_needs_you(tasks)
    in_progress_html = build_in_progress(tasks)
    done_html = build_done(tasks)
    where_html = build_where_things_stand(tasks)
    scoreboard_html = build_scoreboard(s)
    key_dates_html = build_key_dates(tasks)
    maps_html = build_maps()

    counts = {
        "needs_you": len([t for t in tasks if (t.get("Status") or "").lower() == "waiting on greg"
                          or (is_greg_owner(t.get("Owner", "")) and (t.get("Status") or "").lower() != "done")]),
        "in_progress": len([t for t in tasks if (t.get("Status") or "").lower() in ("in progress", "ongoing")]),
        "done": done_n,
        "total": total,
    }

    html = PAGE.format(
        DATE=datetime.date.today().strftime("%A, %B %d, %Y"),
        DONE=done_n,
        TOTAL=total,
        PCT=pct,
        NEEDS_YOU=needs_you_html,
        IN_PROGRESS=in_progress_html,
        WHERE_STAND=where_html,
        SCOREBOARD=scoreboard_html,
        KEY_DATES=key_dates_html,
        DONE_SECTION=done_html,
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
    print(f"Output: {OUT_INDEX}")


if __name__ == "__main__":
    main()
