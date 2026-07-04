# Gregory Shepard - Status Dashboard

A read-only status dashboard for Gregory Shepard's Startup Architect launch, generated as a
static site and published via GitHub Pages. It shows what's done, what's in progress, what
needs Greg's input, the event scoreboard, and key upcoming dates. There is nothing editable
on the page (no inputs, no forms, no JavaScript that writes data) and no sensitive operational
detail (credentials, logins, KPI/pay figures) is ever rendered here.

## Rebuild

The site is generated from the live tracker data, which lives outside this repo. To rebuild:

1. From the project root (`Gregory Project/`), run:
   ```
   python build_greg_dashboard.py
   ```
   This regenerates `gregory-dashboard/docs/index.html` and copies the map images into
   `gregory-dashboard/docs/assets/maps/`.
2. From `gregory-dashboard/`, commit and push:
   ```
   git add -A
   git commit -m "update dashboard"
   git push
   ```

See `PUBLISH.md` for one-time GitHub setup and the live Pages URL.
