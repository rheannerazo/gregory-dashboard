# Publishing this dashboard

`gh` is not authenticated in the environment that built this repo, so none of the steps
below have been run yet. Run them yourself, in order, from inside this `gregory-dashboard/`
folder.

```
# One-time, interactive:
gh auth login

# Create a PRIVATE repo and push (client data -> keep private):
gh repo create gregory-dashboard --private --source . --remote origin --push

# Enable GitHub Pages from the /docs folder on main:
gh api -X POST "repos/{owner}/gregory-dashboard/pages" -f "source[branch]=main" -f "source[path]=/docs"
# (replace {owner} with the GitHub username, or set it in the URL)

# Your live URL will be: https://<username>.github.io/gregory-dashboard/
```

## Note on private repos + GitHub Pages

GitHub Pages on a **private** repo requires a paid GitHub plan (Pro/Team). On a free account
you have two options:

- **(a) Make the repo public** (only if the content is confirmed OK to be public - this
  dashboard is Greg's business status with no secrets, so it likely is):
  ```
  gh repo edit --visibility public --accept-visibility-change-consequences
  ```
- **(b) Keep it private** and share the built `docs/index.html` file directly instead of a
  live Pages URL.

## Updating later

Rerun the build from the project root, then push from `gregory-dashboard/`:

```
python build_greg_dashboard.py
git add -A
git commit -m "update"
git push
```
