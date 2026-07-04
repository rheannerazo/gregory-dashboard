#!/usr/bin/env python3
"""agentctl.py -- single entry point for agents participating in the mission-control loop.

Stdlib only. Reads/writes state/agents.json, state/tasks.json, state/feed.jsonl.
"""
import argparse
import datetime
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
STATE_DIR = ROOT / "state"
AGENTS_PATH = STATE_DIR / "agents.json"
TASKS_PATH = STATE_DIR / "tasks.json"
FEED_PATH = STATE_DIR / "feed.jsonl"

TASK_STATUSES = {"backlog", "assigned", "in_progress", "review", "needs_human", "blocked", "done"}
AGENT_STATUSES = {"working", "idle", "blocked"}


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        f.write("\n")


def append_feed(agent: str, task_id, type_: str, message: str) -> None:
    event = {
        "ts": now(),
        "agent": agent,
        "task_id": task_id,
        "type": type_,
        "message": message,
    }
    with FEED_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(event) + "\n")


def find_task(tasks_data, task_id: str):
    for t in tasks_data["tasks"]:
        if t["id"] == task_id:
            return t
    return None


def find_agent(agents_data, agent_id: str):
    for a in agents_data["agents"]:
        if a["id"] == agent_id:
            return a
    return None


def touch_agent(agents_data, agent_id: str, status=None, current_task=None, clear_task=False):
    agent = find_agent(agents_data, agent_id)
    if agent is None:
        return
    ts = now()
    if status is not None:
        agent["status"] = status
    if clear_task:
        agent["current_task"] = None
    elif current_task is not None:
        agent["current_task"] = current_task
    agent["last_seen"] = ts
    agents_data["updated"] = ts


# ---------------------------------------------------------------------------
# Subcommands
# ---------------------------------------------------------------------------

def cmd_add(args):
    tasks_data = load_json(TASKS_PATH)
    next_id = tasks_data.get("next_id", len(tasks_data["tasks"]) + 1)
    task_id = f"T-{next_id:03d}"
    ts = now()
    tags = [t.strip() for t in args.tags.split(",")] if args.tags else []
    task = {
        "id": task_id,
        "title": args.title,
        "description": args.desc or "",
        "workstream": args.workstream,
        "assignee": args.assignee,
        "reviewer": args.reviewer,
        "status": "backlog",
        "urgency": args.urgency,
        "tags": tags,
        "goal": args.goal or "",
        "created": ts,
        "updated": ts,
    }
    tasks_data["tasks"].append(task)
    tasks_data["next_id"] = next_id + 1
    tasks_data["updated"] = ts
    save_json(TASKS_PATH, tasks_data)

    agents_data = load_json(AGENTS_PATH)
    # Adding a task doesn't change agent state, but touch last_seen for whoever created it if given.
    save_json(AGENTS_PATH, agents_data)

    append_feed(args.assignee, task_id, "note", f"Task {task_id} created: {args.title}")
    print(f"Created {task_id}: {args.title}")


def cmd_claim(args):
    tasks_data = load_json(TASKS_PATH)
    task = find_task(tasks_data, args.task_id)
    if task is None:
        print(f"error: task {args.task_id} not found", file=sys.stderr)
        sys.exit(1)

    if task["status"] not in ("backlog", "assigned"):
        print(f"error: task {args.task_id} cannot be claimed from status '{task['status']}'", file=sys.stderr)
        sys.exit(1)

    ts = now()
    if not task.get("assignee"):
        task["assignee"] = args.agent
    task["status"] = "in_progress"
    task["updated"] = ts
    tasks_data["updated"] = ts
    save_json(TASKS_PATH, tasks_data)

    agents_data = load_json(AGENTS_PATH)
    touch_agent(agents_data, args.agent, status="working", current_task=task["id"])
    save_json(AGENTS_PATH, agents_data)

    append_feed(args.agent, task["id"], "progress", f"{args.agent} claimed {task['id']} and started work.")
    print(f"{args.agent} claimed {task['id']} -> in_progress")


def _feed_type_for_status(status: str) -> str:
    return {
        "review": "review_requested",
        "needs_human": "needs_human",
        "blocked": "blocked",
        "done": "done",
    }.get(status, "progress")


def cmd_update(args):
    tasks_data = load_json(TASKS_PATH)
    task = find_task(tasks_data, args.task_id)
    if task is None:
        print(f"error: task {args.task_id} not found", file=sys.stderr)
        sys.exit(1)

    ts = now()
    if args.status:
        if args.status not in TASK_STATUSES:
            print(f"error: invalid status '{args.status}'", file=sys.stderr)
            sys.exit(1)
        task["status"] = args.status
    task["updated"] = ts
    tasks_data["updated"] = ts
    save_json(TASKS_PATH, tasks_data)

    agents_data = load_json(AGENTS_PATH)
    agent_status = "working"
    if args.status in ("review", "needs_human", "blocked", "done"):
        agent_status = "idle" if args.status == "done" else "blocked" if args.status == "blocked" else "idle"
    touch_agent(
        agents_data,
        args.agent,
        status=agent_status,
        clear_task=(args.status == "done"),
        current_task=task["id"] if args.status != "done" else None,
    )
    save_json(AGENTS_PATH, agents_data)

    feed_type = _feed_type_for_status(args.status) if args.status else "progress"
    append_feed(args.agent, task["id"], feed_type, args.message)
    print(f"{task['id']} updated -> {task['status']}")


def cmd_review(args):
    tasks_data = load_json(TASKS_PATH)
    task = find_task(tasks_data, args.task_id)
    if task is None:
        print(f"error: task {args.task_id} not found", file=sys.stderr)
        sys.exit(1)

    if args.agent == task.get("assignee"):
        print(
            f"error: reviewer '{args.agent}' cannot be the same as assignee '{task.get('assignee')}' "
            f"(maker/checker separation)",
            file=sys.stderr,
        )
        sys.exit(1)

    ts = now()
    agents_data = load_json(AGENTS_PATH)

    if args.verdict == "approve":
        task["status"] = "done"
        feed_type = "approved"
        # free both assignee and reviewer
        touch_agent(agents_data, task.get("assignee"), status="idle", clear_task=True)
        touch_agent(agents_data, args.agent, status="idle", clear_task=True)
    elif args.verdict == "reject":
        task["status"] = "in_progress"
        feed_type = "rejected"
        # send back to assignee, reviewer goes idle
        touch_agent(agents_data, task.get("assignee"), status="working", current_task=task["id"])
        touch_agent(agents_data, args.agent, status="idle", clear_task=True)
    else:
        print(f"error: invalid verdict '{args.verdict}' (expected approve|reject)", file=sys.stderr)
        sys.exit(1)

    task["updated"] = ts
    tasks_data["updated"] = ts
    save_json(TASKS_PATH, tasks_data)
    save_json(AGENTS_PATH, agents_data)

    append_feed(args.agent, task["id"], feed_type, args.message)
    print(f"{task['id']} review by {args.agent}: {args.verdict} -> {task['status']}")


def cmd_log(args):
    append_feed(args.agent, args.task, "note", args.message)
    agents_data = load_json(AGENTS_PATH)
    agent = find_agent(agents_data, args.agent)
    if agent is not None:
        agent["last_seen"] = now()
        agents_data["updated"] = now()
        save_json(AGENTS_PATH, agents_data)
    print("logged.")


def cmd_status(args):
    tasks_data = load_json(TASKS_PATH)
    agents_data = load_json(AGENTS_PATH)

    by_status = {}
    for t in tasks_data["tasks"]:
        by_status.setdefault(t["status"], []).append(t)

    print("=== Task queue ===")
    for status in ["needs_human", "blocked", "review", "in_progress", "assigned", "backlog", "done"]:
        items = by_status.get(status, [])
        if not items:
            continue
        print(f"\n[{status}] ({len(items)})")
        for t in items:
            urgency_tag = " *URGENT*" if t.get("urgency") == "urgent" else ""
            print(f"  {t['id']} {t['title']} (assignee={t['assignee']}, reviewer={t['reviewer']}){urgency_tag}")

    print("\n=== Agents ===")
    for a in agents_data["agents"]:
        task = f" ({a['current_task']})" if a.get("current_task") else ""
        print(f"  {a['id']:8s} {a['name']:8s} [{a['status']}]{task}  role={a['role']}  model={a['model']}")


def cmd_build(args):
    subprocess.run([sys.executable, "build_mission_control.py"], cwd=str(ROOT), check=True)


def cmd_publish(args):
    cmd_build(args)
    subprocess.run(["git", "add", "state", "docs"], cwd=str(ROOT), check=True)
    diff = subprocess.run(
        ["git", "diff", "--cached", "--quiet"], cwd=str(ROOT)
    )
    if diff.returncode == 0:
        print("Nothing to commit; state and docs are unchanged.")
        return
    subprocess.run(
        ["git", "commit", "-m", "mission control: state update"], cwd=str(ROOT), check=True
    )
    subprocess.run(["git", "push"], cwd=str(ROOT), check=True)


# ---------------------------------------------------------------------------
# CLI wiring
# ---------------------------------------------------------------------------

def build_parser():
    parser = argparse.ArgumentParser(prog="agentctl.py", description="Mission control agent loop CLI.")
    sub = parser.add_subparsers(dest="command", required=True)

    p_add = sub.add_parser("add", help="Add a new task.")
    p_add.add_argument("title")
    p_add.add_argument("--desc", default="")
    p_add.add_argument("--workstream", required=True)
    p_add.add_argument("--assignee", required=True)
    p_add.add_argument("--reviewer", required=True)
    p_add.add_argument("--urgency", choices=["normal", "urgent"], default="normal")
    p_add.add_argument("--goal", default="")
    p_add.add_argument("--tags", default="")
    p_add.set_defaults(func=cmd_add)

    p_claim = sub.add_parser("claim", help="Claim a task.")
    p_claim.add_argument("task_id")
    p_claim.add_argument("--agent", required=True)
    p_claim.set_defaults(func=cmd_claim)

    p_update = sub.add_parser("update", help="Update a task's status/progress.")
    p_update.add_argument("task_id")
    p_update.add_argument("--agent", required=True)
    p_update.add_argument("--status", choices=sorted(TASK_STATUSES))
    p_update.add_argument("--message", required=True)
    p_update.set_defaults(func=cmd_update)

    p_review = sub.add_parser("review", help="Review a task (maker/checker separation enforced).")
    p_review.add_argument("task_id")
    p_review.add_argument("--agent", required=True)
    p_review.add_argument("--verdict", choices=["approve", "reject"], required=True)
    p_review.add_argument("--message", required=True)
    p_review.set_defaults(func=cmd_review)

    p_log = sub.add_parser("log", help="Append a free-form note to the feed.")
    p_log.add_argument("--agent", required=True)
    p_log.add_argument("--message", required=True)
    p_log.add_argument("--task", default=None)
    p_log.set_defaults(func=cmd_log)

    p_status = sub.add_parser("status", help="Print a compact status summary.")
    p_status.set_defaults(func=cmd_status)

    p_build = sub.add_parser("build", help="Run build_mission_control.py.")
    p_build.set_defaults(func=cmd_build)

    p_publish = sub.add_parser("publish", help="Build, commit, and push state + docs.")
    p_publish.set_defaults(func=cmd_publish)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
