#!/usr/bin/env python3
"""
Build a static, READ-ONLY "Mission Control" page for monitoring autonomous AI
agents: a squad sidebar, a mission-queue kanban board, and a live feed.

Reads state/agents.json, state/tasks.json, state/feed.jsonl (see agentctl.py /
whatever writes them) and renders a single self-contained HTML file to
gregory-dashboard/docs/agents.html, safe to publish alongside the existing
Greg-facing dashboard on GitHub Pages.

This script does not modify build_greg_dashboard.py or docs/index.html, and
does not link agents.html from anywhere -- it is a standalone render.

Run:  python build_mission_control.py [--state-dir PATH]
"""

import os
import sys
import json
import argparse
import datetime
import html as html_lib

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_STATE = os.path.join(HERE, "state")
OUT_ROOT = os.path.join(HERE, "docs")
OUT_FILE = os.path.join(OUT_ROOT, "agents.html")
REPO_URL = "https://github.com/rheannerazo/gregory-dashboard"

NAVY = "#0A1730"
BLUE = "#1A8CF0"
BG = "#f5f7fa"

# Same palette family as build_greg_dashboard.py's STATUS_COLORS, extended
# with the agent/task statuses this page needs.
STATUS_COLORS = {
    "working": "#16a34a",
    "idle": "#6b7280",
    "blocked": "#dc2626",
    "backlog": "#6b7280",
    "assigned": "#6b7280",
    "in_progress": "#1A8CF0",
    "review": "#7c3aed",
    "needs_human": "#d97706",
    "done": "#16a34a",
}

FEED_TYPE_COLORS = {
    "approved": "#16a34a",
    "done": "#16a34a",
    "rejected": "#dc2626",
    "blocked": "#dc2626",
    "needs_human": "#d97706",
    "review_requested": "#1A8CF0",
}
FEED_TYPE_DEFAULT = "#6b7280"

# Sanitization blocklist -- never surface these in any rendered task text.
# Copied verbatim from build_greg_dashboard.py so both public pages agree.
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


# ---------------- helpers ----------------
def esc(x):
    return html_lib.escape(str(x) if x is not None else "")


def clean(v):
    if v is None:
        return ""
    return str(v).strip()


def load_json(path, required_top_key=None):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if required_top_key and required_top_key not in data:
        data[required_top_key] = []
    return data


def load_jsonl(path):
    out = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return out


def friendly_time(ts):
    """Render an ISO timestamp as something human-scannable; fall back to the
    raw string (sanitized) if it doesn't parse."""
    if not ts:
        return ""
    s = str(ts).strip()
    try:
        s2 = s[:-1] + "+00:00" if s.endswith("Z") else s
        d = datetime.datetime.fromisoformat(s2)
        return d.strftime("%b %d, %H:%M")
    except ValueError:
        return esc(s)


def task_is_sensitive(t):
    """True if the task's own title/workstream (not just description/goal)
    touches something on the blocklist -- these shouldn't be named on the
    public page at all, even though they still count toward totals."""
    return hits_blocklist(t.get("title", "")) or hits_blocklist(t.get("workstream", ""))


# ---------------- agents / squad sidebar ----------------
def build_squad(agents):
    if not agents:
        return '<div class="card"><p class="muted">No agents reporting.</p></div>'
    cards = ""
    for a in agents:
        status = (a.get("status") or "idle").lower()
        color = STATUS_COLORS.get(status, "#6b7280")
        pulse = ' pulse' if status == "working" else ""
        name = esc(safe(clean(a.get("name", "?"))))
        role = esc(safe(clean(a.get("role", ""))))
        model = esc(clean(a.get("model", "")))
        task_id = clean(a.get("current_task", "") or "")
        task_html = ""
        if task_id and not hits_blocklist(task_id):
            task_html = f'<a class="agent-task" href="#{esc(task_id)}">{esc(task_id)}</a>'
        else:
            task_html = '<span class="agent-task muted">&mdash;</span>'
        cards += (
            f'<div class="agentcard">'
            f'<div class="agent-row">'
            f'<span class="dot{pulse}" style="background:{color}"></span>'
            f'<span class="agent-name">{name}</span>'
            f'</div>'
            f'<div class="agent-role">{role}</div>'
            f'<div class="agent-meta">'
            f'<span class="chip-model">{model}</span>'
            f'{task_html}'
            f'</div>'
            f'</div>')
    return f'<div class="squadlist">{cards}</div>'


# ---------------- needs you strip ----------------
def build_needs_you(tasks):
    items = [t for t in tasks if (t.get("status") or "").lower() == "needs_human"
             and not task_is_sensitive(t)]
    if not items:
        return ""
    cards = ""
    for t in items:
        title = esc(clean(t.get("title", "")))
        desc = esc(safe(clean(t.get("description", ""))))
        goal = esc(safe(clean(t.get("goal", ""))))
        tid = esc(clean(t.get("id", "")))
        cards += (
            f'<div class="needcard">'
            f'<div class="ny-title">{title}</div>'
            f'<div class="ny-line">{desc}</div>'
            f'{f"<div class=" + chr(34) + "ny-goal" + chr(34) + ">Done when: " + goal + "</div>" if goal else ""}'
            f'<a class="ny-link" href="#{tid}">Jump to task {tid}</a>'
            f'</div>')
    return (
        '<section class="needs-you">'
        '<div class="wrap">'
        '<h2 class="ny-heading">Needs You &mdash; Human Circuit Breaker</h2>'
        f'<div class="needgrid">{cards}</div>'
        '</div>'
        '</section>')


# ---------------- mission queue / kanban ----------------
QUEUE_COLUMNS = [
    ("assigned", "Assigned", ("backlog", "assigned")),
    ("in_progress", "In Progress", ("in_progress",)),
    ("review", "Review", ("review",)),
    ("blocked", "Blocked", ("blocked", "needs_human")),
    ("done", "Done", ("done",)),
]


def task_card(t):
    tid = esc(clean(t.get("id", "")))
    urgent = (t.get("urgency") or "").lower() == "urgent"
    urgent_cls = " urgent" if urgent else ""
    title = esc(clean(t.get("title", "")))
    ws = esc(safe(clean(t.get("workstream", ""))))
    tags = t.get("tags") or []
    tags_html = "".join(f'<span class="tag">{esc(safe(clean(tag)))}</span>' for tag in tags if safe(clean(tag)))
    assignee = esc(safe(clean(t.get("assignee", ""))))
    reviewer = esc(safe(clean(t.get("reviewer", ""))))
    route = ""
    if assignee or reviewer:
        route = f'<div class="task-route">{assignee or "&mdash;"} &rarr; {reviewer or "&mdash;"}</div>'
    goal = esc(safe(clean(t.get("goal", ""))))
    goal_html = f'<div class="task-goal">Done when: {goal}</div>' if goal else ""
    updated = friendly_time(t.get("updated", ""))
    return (
        f'<div class="taskcard{urgent_cls}" id="{tid}">'
        f'{f"<div class=" + chr(34) + "urgent-flag" + chr(34) + ">Urgent</div>" if urgent else ""}'
        f'<div class="task-id">{tid}</div>'
        f'<div class="task-title">{title}</div>'
        f'{f"<div class=" + chr(34) + "ws-chip" + chr(34) + ">" + ws + "</div>" if ws else ""}'
        f'{f"<div class=" + chr(34) + "tag-row" + chr(34) + ">" + tags_html + "</div>" if tags_html else ""}'
        f'{route}'
        f'{goal_html}'
        f'<div class="task-updated muted">Updated {esc(updated)}</div>'
        f'</div>')


def build_queue(tasks):
    visible = [t for t in tasks if not task_is_sensitive(t)]
    columns_html = ""
    for key, label, statuses in QUEUE_COLUMNS:
        col_tasks = [t for t in visible if (t.get("status") or "").lower() in statuses]
        col_tasks.sort(key=lambda t: clean(t.get("updated", "")), reverse=True)
        if key == "done":
            col_tasks = col_tasks[:5]
        cards = "".join(task_card(t) for t in col_tasks) or '<p class="muted col-empty">Nothing here.</p>'
        columns_html += (
            f'<div class="col">'
            f'<div class="col-head"><span>{esc(label)}</span><span class="col-count">{len(col_tasks)}</span></div>'
            f'<div class="col-body">{cards}</div>'
            f'</div>')
    return f'<div class="board">{columns_html}</div>'


# ---------------- live feed ----------------
def build_feed(events):
    if not events:
        return '<div class="card"><p class="muted">No activity yet.</p></div>'

    def event_ts(e):
        return clean(e.get("ts", ""))

    ordered = sorted(events, key=event_ts, reverse=True)[:50]
    rows = ""
    for e in ordered:
        etype = clean(e.get("type", "")).lower()
        color = FEED_TYPE_COLORS.get(etype, FEED_TYPE_DEFAULT)
        agent = esc(safe(clean(e.get("agent", ""))))
        msg = esc(safe(clean(e.get("message", ""))))
        ts = esc(friendly_time(e.get("ts", "")))
        task_id = clean(e.get("task_id", "") or "")
        task_link = ""
        if task_id and not hits_blocklist(task_id):
            task_link = f' <a class="feed-task" href="#{esc(task_id)}">{esc(task_id)}</a>'
        rows += (
            f'<div class="feeditem">'
            f'<div class="feed-top">'
            f'<span class="feed-time">{ts}</span>'
            f'<span class="feed-badge" style="background:{color}">{esc(etype or "note")}</span>'
            f'</div>'
            f'<div class="feed-agent">{agent}{task_link}</div>'
            f'<div class="feed-msg">{msg}</div>'
            f'</div>')
    return f'<div class="feedlist">{rows}</div>'


# ---------------- page template ----------------
PAGE = """<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<title>Mission Control &mdash; Agent Operations</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Anton&family=Montserrat:wght@500;600;700;800&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@500&display=swap" rel="stylesheet">
<style>
:root{{
  --navy:#0A1730;
  --blue:#1A8CF0;
  --bg:#f5f7fa;
  --panel:#0f1f3d;
  --panel-border:#1c2f52;
  --text:#e6ebf5;
  --muted:#8b9bbd;
}}
*{{box-sizing:border-box;}}
body{{margin:0;font-family:'Inter',system-ui,sans-serif;color:#1f2937;background:var(--bg);-webkit-font-smoothing:antialiased;}}
h1,h2,h3,h4{{font-family:'Montserrat',system-ui,sans-serif;}}
.mono{{font-family:'JetBrains Mono',ui-monospace,monospace;}}
.muted{{color:var(--muted);font-size:.92em;}}

/* Header */
header{{background:var(--navy);color:#fff;padding:28px 24px;}}
.header-inner{{max-width:1400px;margin:0 auto;display:flex;flex-wrap:wrap;gap:18px;justify-content:space-between;align-items:flex-end;}}
header .eyebrow{{color:var(--blue);font-weight:700;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;}}
header h1{{font-family:'Anton',Montserrat,sans-serif;font-weight:400;margin:0 0 6px;font-size:clamp(26px,4vw,38px);letter-spacing:.5px;}}
header .subtitle{{color:#cdd7e8;font-size:14px;margin:0;}}
.stat-chips{{display:flex;flex-wrap:wrap;gap:10px;}}
.chip{{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);border-radius:10px;padding:8px 14px;text-align:center;min-width:110px;}}
.chip .chip-num{{font-size:20px;font-weight:800;color:#fff;}}
.chip .chip-label{{font-size:10.5px;color:#9fb0cc;text-transform:uppercase;letter-spacing:.5px;margin-top:2px;}}
.chip.warn .chip-num{{color:#f59e0b;}}
.chip.built-at{{background:transparent;border:none;text-align:right;min-width:0;}}
.chip.built-at .chip-label{{margin-top:4px;}}

/* Needs you strip */
.needs-you{{background:#3a2205;border-bottom:3px solid #d97706;}}
.needs-you .wrap{{max-width:1400px;margin:0 auto;padding:20px 24px;}}
.ny-heading{{color:#fbbf24;font-size:14px;text-transform:uppercase;letter-spacing:1px;margin:0 0 12px;border:none;padding:0;}}
.needgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;}}
.needcard{{background:#241505;border:1px solid #d97706;border-radius:10px;padding:14px 16px;}}
.ny-title{{font-weight:700;color:#fde68a;font-size:14.5px;margin-bottom:6px;}}
.ny-line{{font-size:13px;color:#f3e2c4;line-height:1.5;}}
.ny-goal{{font-size:12px;color:#d1a86a;margin-top:8px;font-style:italic;}}
.ny-link{{display:inline-block;margin-top:10px;font-size:12px;color:#fbbf24;text-decoration:none;font-weight:700;}}
.ny-link:hover{{text-decoration:underline;}}

/* Layout grid */
.layout{{max-width:1400px;margin:0 auto;padding:24px;display:grid;grid-template-columns:220px minmax(0,1fr) 320px;gap:20px;align-items:start;}}
.panel{{background:var(--panel);border:1px solid var(--panel-border);border-radius:14px;padding:16px;color:var(--text);}}
.panel h2{{font-size:13px;text-transform:uppercase;letter-spacing:1.5px;color:#9fb0cc;margin:0 0 14px;padding-bottom:10px;border-bottom:1px solid var(--panel-border);}}
.sidebar{{position:sticky;top:20px;}}
.feedpanel{{position:sticky;top:20px;max-height:calc(100vh - 40px);overflow-y:auto;}}

/* Squad */
.squadlist{{display:flex;flex-direction:column;gap:10px;}}
.agentcard{{background:#132447;border:1px solid var(--panel-border);border-radius:10px;padding:10px 12px;}}
.agent-row{{display:flex;align-items:center;gap:8px;}}
.dot{{width:9px;height:9px;border-radius:50%;flex-shrink:0;}}
.dot.pulse{{box-shadow:0 0 0 rgba(22,163,74,.6);animation:pulse 1.8s infinite;}}
@keyframes pulse{{
  0%{{box-shadow:0 0 0 0 rgba(22,163,74,.6);}}
  70%{{box-shadow:0 0 0 7px rgba(22,163,74,0);}}
  100%{{box-shadow:0 0 0 0 rgba(22,163,74,0);}}
}}
.agent-name{{font-weight:700;font-size:13.5px;color:#fff;}}
.agent-role{{font-size:11.5px;color:var(--muted);margin:3px 0 8px 17px;}}
.agent-meta{{display:flex;align-items:center;justify-content:space-between;margin-left:17px;}}
.chip-model{{font-family:'JetBrains Mono',ui-monospace,monospace;font-size:10px;background:#0a1730;border:1px solid var(--panel-border);border-radius:5px;padding:2px 6px;color:#8bb4f0;}}
.agent-task{{font-size:11px;color:var(--blue);text-decoration:none;font-weight:600;}}
.agent-task:hover{{text-decoration:underline;}}

/* Board */
.board{{display:grid;grid-template-columns:repeat(5,minmax(220px,1fr));gap:14px;overflow-x:auto;}}
.col{{background:#f0f2f6;border-radius:12px;padding:10px;min-width:0;}}
.col-head{{display:flex;justify-content:space-between;align-items:center;font-size:12px;font-weight:800;color:var(--navy);text-transform:uppercase;letter-spacing:.5px;padding:4px 6px 10px;}}
.col-count{{background:#dbe4f3;color:var(--navy);border-radius:10px;padding:1px 8px;font-size:11px;}}
.col-body{{display:flex;flex-direction:column;gap:10px;}}
.col-empty{{padding:6px;font-size:12px;}}

.taskcard{{background:#fff;border-radius:10px;padding:12px 14px;box-shadow:0 2px 8px rgba(10,23,48,.08);border-left:3px solid transparent;position:relative;scroll-margin-top:20px;}}
.taskcard.urgent{{border-left-color:#dc2626;}}
.urgent-flag{{position:absolute;top:8px;right:10px;background:#dc2626;color:#fff;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;border-radius:4px;padding:2px 6px;}}
.task-id{{font-family:'JetBrains Mono',ui-monospace,monospace;font-size:10px;color:#9ca3af;margin-bottom:4px;}}
.task-title{{font-weight:700;color:var(--navy);font-size:13.5px;line-height:1.35;margin-bottom:6px;padding-right:50px;}}
.ws-chip{{display:inline-block;background:#eef2ff;color:#3730a3;font-size:10.5px;font-weight:700;border-radius:6px;padding:2px 8px;margin-bottom:6px;}}
.tag-row{{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:6px;}}
.tag{{background:#f3f4f6;color:#4b5563;font-size:10px;border-radius:5px;padding:2px 6px;}}
.task-route{{font-size:11.5px;color:#374151;font-weight:600;margin-bottom:6px;}}
.task-goal{{font-size:11.5px;color:#6b7280;font-style:italic;line-height:1.4;margin-bottom:6px;}}
.task-updated{{font-size:10.5px;}}

/* Feed */
.feedlist{{display:flex;flex-direction:column;gap:10px;}}
.feeditem{{background:#132447;border:1px solid var(--panel-border);border-radius:10px;padding:10px 12px;}}
.feed-top{{display:flex;justify-content:space-between;align-items:center;margin-bottom:5px;}}
.feed-time{{font-size:10.5px;color:var(--muted);font-family:'JetBrains Mono',ui-monospace,monospace;}}
.feed-badge{{font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;border-radius:5px;padding:2px 7px;color:#fff;}}
.feed-agent{{font-size:12px;font-weight:700;color:#fff;margin-bottom:3px;}}
.feed-task{{font-size:11px;color:var(--blue);text-decoration:none;font-weight:600;margin-left:4px;}}
.feed-task:hover{{text-decoration:underline;}}
.feed-msg{{font-size:12px;color:#c3cee3;line-height:1.4;}}

.card{{background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 2px 10px rgba(10,23,48,.06);}}

.foot{{max-width:1400px;margin:0 auto;padding:10px 24px 60px;color:#9ca3af;font-size:12px;text-align:center;}}
.foot a{{color:var(--blue);text-decoration:none;}}
.foot a:hover{{text-decoration:underline;}}

@media (max-width:1100px){{
  .layout{{grid-template-columns:1fr;}}
  .sidebar, .feedpanel{{position:static;max-height:none;}}
  .board{{grid-template-columns:repeat(5,minmax(200px,1fr));}}
}}
@media (max-width:640px){{
  .layout{{padding:16px;}}
  header{{padding:22px 16px;}}
  .board{{grid-template-columns:1fr;}}
  .header-inner{{align-items:flex-start;}}
}}
</style>
</head>
<body>
<header>
  <div class="header-inner">
    <div>
      <div class="eyebrow">The Startup Architect</div>
      <h1>Mission Control</h1>
      <p class="subtitle">Agent operations &mdash; The Startup Architect</p>
    </div>
    <div class="stat-chips">
      <div class="chip"><div class="chip-num">{AGENTS_ACTIVE}</div><div class="chip-label">Agents active</div></div>
      <div class="chip"><div class="chip-num">{TASKS_QUEUED}</div><div class="chip-label">Tasks in queue</div></div>
      <div class="chip warn"><div class="chip-num">{NEEDS_ATTENTION}</div><div class="chip-label">Need attention</div></div>
      <div class="chip built-at"><div class="chip-label">Built {BUILT_AT}</div></div>
    </div>
  </div>
</header>

{NEEDS_YOU}

<div class="layout">
  <aside class="panel sidebar">
    <h2>Squad</h2>
    {SQUAD}
  </aside>

  <main class="panel">
    <h2>Mission Queue</h2>
    {QUEUE}
  </main>

  <aside class="panel feedpanel">
    <h2>Live Feed</h2>
    {FEED}
  </aside>
</div>

<div class="foot">Read-only render of state/ &mdash; agents update via agentctl.py &middot; <a href="{REPO_URL}">{REPO_URL}</a></div>
</body></html>
"""


def render(state_dir):
    agents_path = os.path.join(state_dir, "agents.json")
    tasks_path = os.path.join(state_dir, "tasks.json")
    feed_path = os.path.join(state_dir, "feed.jsonl")

    missing = [p for p in (agents_path, tasks_path, feed_path) if not os.path.exists(p)]
    if missing:
        names = ", ".join(os.path.basename(p) for p in missing)
        raise SystemExit(
            f"ERROR: missing state file(s) in {state_dir}: {names}\n"
            "Refusing to invent state -- another process owns state/. "
            "Pass --state-dir to point at a directory that has agents.json, "
            "tasks.json, and feed.jsonl, or wait until they exist."
        )

    agents_data = load_json(agents_path, required_top_key="agents")
    tasks_data = load_json(tasks_path, required_top_key="tasks")
    events = load_jsonl(feed_path)

    agents = agents_data.get("agents") or []
    tasks = tasks_data.get("tasks") or []

    agents_active = sum(1 for a in agents if (a.get("status") or "").lower() == "working")
    tasks_queued = sum(1 for t in tasks if (t.get("status") or "").lower() != "done")
    needs_attention = sum(
        1 for t in tasks
        if (t.get("status") or "").lower() in ("needs_human", "blocked")
        or (t.get("urgency") or "").lower() == "urgent"
    )

    squad_html = build_squad(agents)
    needs_you_html = build_needs_you(tasks)
    queue_html = build_queue(tasks)
    feed_html = build_feed(events)

    built_at = datetime.datetime.now().strftime("%b %d, %Y %H:%M")

    html = PAGE.format(
        AGENTS_ACTIVE=agents_active,
        TASKS_QUEUED=tasks_queued,
        NEEDS_ATTENTION=needs_attention,
        BUILT_AT=esc(built_at),
        NEEDS_YOU=needs_you_html,
        SQUAD=squad_html,
        QUEUE=queue_html,
        FEED=feed_html,
        REPO_URL=esc(REPO_URL),
    )
    counts = {
        "agents_active": agents_active,
        "tasks_queued": tasks_queued,
        "needs_attention": needs_attention,
        "agents_total": len(agents),
        "tasks_total": len(tasks),
        "feed_events": len(events),
    }
    return html, counts


def main():
    parser = argparse.ArgumentParser(description="Build the Mission Control agent-ops page.")
    parser.add_argument("--state-dir", default=DEFAULT_STATE,
                         help="Directory containing agents.json, tasks.json, feed.jsonl (default: ./state)")
    parser.add_argument("--out", default=OUT_FILE,
                         help="Output HTML path (default: docs/agents.html)")
    args = parser.parse_args()

    state_dir = os.path.abspath(args.state_dir)
    out_path = os.path.abspath(args.out)

    html, counts = render(state_dir)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print("Build complete.")
    print(f"  State dir:       {state_dir}")
    print(f"  Agents active:   {counts['agents_active']} / {counts['agents_total']}")
    print(f"  Tasks in queue:  {counts['tasks_queued']} / {counts['tasks_total']}")
    print(f"  Need attention:  {counts['needs_attention']}")
    print(f"  Feed events:     {counts['feed_events']}")
    print(f"Output: {out_path}")


if __name__ == "__main__":
    main()
