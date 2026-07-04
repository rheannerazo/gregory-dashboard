# Agent Loop Protocol

## How this works

The repo is the shared memory—agents forget, the repo doesn't. All persistent state lives in `state/` as structured JSON: `agents.json` (roster of active agents), `tasks.json` (work queue with status), and `feed.jsonl` (event log). The read-only dashboard (`docs/agents.html`) renders this state and is rebuilt by `build_mission_control.py`. All mutations go through `agentctl.py` (the command interface); never hand-edit state files directly.

## Joining a session

1. Read `state/tasks.json` to see the queue.
2. Pick a task matching your role and skill.
3. Run `python agentctl.py claim T-XXX --agent <your-id>` to lock the task.
4. Do the work, consulting the task's `goal` (the stopping condition).
5. Run `python agentctl.py update T-XXX --agent <your-id> --status review --message "..."` when done.
6. A checker (different agent) will review before the task closes.

## Command reference

| Command | Purpose |
|---------|---------|
| `add` | Create a new task (add to backlog) |
| `claim` | Lock a task to your agent ID (blocks reassignment) |
| `update` | Report status change or progress message |
| `review` | Inspect a task in review; approve or bounce back to assignee |
| `log` | Append a timestamped event to feed.jsonl |
| `status` | Show task queue state (summary) |
| `build` | Rebuild `docs/agents.html` from state/ |
| `publish` | Build + commit + push (end-of-session) |

**Task statuses:** backlog, assigned, in_progress, review, needs_human, blocked, done.

## Rules

- **Maker ≠ checker:** you may not review a task you implemented. `agentctl.py review` enforces this.
- **Every task has a `goal`:** a verifiable stopping condition. Done only when the goal is met and the reviewer approves.
- **Judgment calls:** anything requiring Rheanne's or Greg's decision → `--status needs_human`; never guess.
- **Public repo:** no credentials, client-sensitive details, or internal-only notes in task text or feed messages (a sanitizer also strips them at render).
- **Scheduled work:** recurring discovery tasks belong in the Cowork cloud roster, not local schedulers.
- **End of session:** run `python agentctl.py publish` (build + commit + push) before signing off.

## The roster

Five seed agents, each routed by capability and cost:

| Agent | Model | Role | Tasks |
|-------|-------|------|-------|
| **Atlas** (planner) | Opus | Planning & judgment | Break down features, design flows, high-stakes decisions |
| **Forge** (builder) | Sonnet | Implementation | Code, content, builds, PRs |
| **Vet** (checker) | Sonnet | Independent review | QA, code review, fact-check (never your own work) |
| **Scout** (discovery) | Haiku | Triage & search | Find files, scan logs, quick lookups, monitor feeds |
| **Scribe** (docs) | Haiku | Documentation & publishing | Render dashboards, write guides, publish outputs |

Route work to the cheapest model that can do it. Don't over-spec.
