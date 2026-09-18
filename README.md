# NeoChrono — GitHub-only edition

This rebuild removes Google Apps Script and Google Sheets from the runtime architecture.

## Architecture

```text
GitHub Pages (HTML/CSS/JavaScript)
            |
            | HTTPS / GitHub REST API
            v
Private GitHub data repository
  data/admins.json
  data/users.json
  data/skills.json
  data/shifts.json
  data/staffing.json
  data/requests.json
  data/weekly-availability.json
  data/preceptor-calendar.json
  data/schedule.json
  data/settings.json
  data/audit.json
  data/swaps.json
```

The browser reads and writes JSON with GitHub's Repository Contents API. Every write becomes a Git commit, so you automatically get file history and rollback capability.

## Recommended repository layout

Use **two repositories**:

1. `neochrono-app` — this code, published with GitHub Pages.
2. `neochrono-data` — a **private** repository containing the JSON database.

Do not put live pharmacy scheduling data in a public repository.

## Authentication model

A GitHub-only static site has no private server where a repository credential can be hidden. Therefore the user supplies a fine-grained GitHub token at sign-in. The token is stored only in `sessionStorage` and disappears when that browser session is cleared. Never commit a token to this repository.

The app also keeps NeoChrono administrator usernames/password hashes in `admins.json`. In a pure static app, the real write-access boundary is still the GitHub repository permission attached to the token. The NeoChrono password is an additional application-level gate, not a replacement for GitHub permissions.

## First setup

1. Create an empty private repository for the data, for example `neochrono-data`.
2. Create a fine-grained GitHub personal access token scoped only to that repository with **Contents: Read and write**.
3. Publish this application repository with GitHub Pages.
4. Open the Pages site and select **First-time database setup**.
5. Enter the data repository owner, name, branch, folder, and token.
6. Create the first NeoChrono administrator username/password.
7. The app creates all JSON database files automatically.
8. Sign in and complete the pharmacist skills, 7-on/7-off anchors, preferred shifts/home units, weekend groups, preceptor settings, and staffing requirements before generating a schedule.

## Current scheduling rules implemented

The GitHub scheduler includes the core rules from the current NeoChrono project:

- Sunday–Saturday scheduling weeks.
- Regular pharmacists: maximum/exact target of 5 workdays per complete week.
- Maximum 5 consecutive workdays for regular pharmacists.
- Configurable weekly-hour cap.
- 56-day / 8-week periods and a configurable 320 credited-hour target.
- One assignment per employee per calendar day.
- Skills are required for automatic scheduling.
- Approved PTO and recurring weekly availability block assignments.
- A/B/C weekend rotation.
- PRN employees are weekend-only and require weekend availability records.
- 7-on/7-off work pattern with a required anchor date and dedicated preferred shift.
- Evening → next-day Day/Morning transition blocked.
- Resident E2 requirement per week.
- Resident weekend restrictions.
- EDD/EDE paired coverage attempts to use two different pharmacists.
- Preceptor calendar: ON weekdays stay on the home/preferred unit; OFF weekdays may use home unit, E1, or E2.
- Off-day coverage skills are stored separately from normal skills. They are intentionally not granted to the automatic scheduler in this first GitHub build.
- PTO: first two requests for each requested calendar date are automatically approved first-come, first-served; later requests remain pending unless manually reviewed.

The scheduler validates the completed period and reports exact five-day, hour, resident-E2, PTO, skill, duplicate-day, 7-on/7-off, consecutive-day, and evening-to-morning violations.

## Important GitHub limitation

GitHub is not a transactional database. It is workable for this project, but each save creates a commit and two administrators editing the same JSON file at the same moment can conflict. The storage layer uses the current GitHub file SHA and retries conflicts instead of silently overwriting newer data.

For a small scheduler with a few administrators, this is workable. If the app later has many simultaneous users, a real database will scale better.

## Files

- `index.html` — static application shell.
- `assets/css/app.css` — UI styles.
- `assets/js/github-store.js` — GitHub REST API read/write layer.
- `assets/js/data-service.js` — JSON database and administrator login operations.
- `assets/js/rules.js` — rules, PTO reconciliation, validation, and reporting helpers.
- `assets/js/scheduler.js` — in-browser schedule generation algorithm.
- `assets/js/app.js` — screens and user interactions.
- `assets/js/seed.js` — initial settings, shifts, staffing, and pharmacist seed records.
- `docs/SETUP.md` — detailed deployment instructions.
- `docs/DATA_MODEL.md` — database file descriptions.
