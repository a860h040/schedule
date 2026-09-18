# NeoChrono — GitHub-Only Rebuild

This package rebuilds the pharmacy scheduling project so the **runtime no longer uses Google Apps Script or Google Sheets**.

## Runtime architecture

```text
GitHub Pages (HTML/CSS/JavaScript)
        ↓
GitHub REST API
        ↓
Private GitHub data repository
        ↓
data/database.json
```

Normal users see the same basic login pattern as the existing project:

```text
Username
Password
[ Sign In ]
```

They do **not** enter a GitHub token on the login page.

## Important GitHub-only security limitation

GitHub Pages is static hosting. It cannot keep a private server secret. To read/write a **private** GitHub data repository without an external backend, an administrator must provision each authorized browser once with a fine-grained GitHub token using `provision.html`.

The token is stored in that browser's local storage and is not embedded in the source code. Restrict the token to only the NeoChrono data repository and only the permissions required for repository contents.

Because any secret used by a browser can ultimately be inspected by a sufficiently privileged user of that browser, this pure GitHub-only design should **not** be used for PHI or other data that requires a protected server-side secret. If the application later needs stronger security, keep the GitHub database but move repository access behind a small server-side API.

## Main functionality included

The GitHub version includes the same functional areas from the current scheduling project:

- Username/password login
- Administrator and pharmacist permissions
- Password change/reset
- Monthly schedule calendar
- My Schedule
- Up to 6-pharmacist filtering
- Show Unfilled filter
- Schedule generation
- Locked assignments preserved during regeneration
- Manual assignment editing
- Manual-rule warnings and administrator override reason
- Assignment lock / unlock / remove
- Employee/pharmacist management
- Employee Skills
- Off-Day Coverage Skills
- Shift Definitions
- Staffing Requirements
- PTO requests
- First two PTO requests **for each individual date** auto-approved first-come
- Request #3+ for the same date remains Pending for administrator review
- One-day availability records
- Recurring weekly availability
- Preceptor calendar
- Preceptor ON = home/preferred unit
- Preceptor OFF = home/preferred unit or E1/E2
- Preceptor evening target when not precepting the whole month
- 7-on / 7-off scheduling
- A/B/C weekend rotation and per-employee weekend anchor
- Residents on assigned weekends
- Resident E2 weekly requirement
- EDD/EDE paired ED coverage using different pharmacists
- Exactly 5 regular workdays per complete Sunday-Saturday week validation
- Maximum 5 consecutive regular workdays
- Evening → next-day Morning/Day restriction
- One shift per day
- Weekly hour maximum
- Exact schedule-period hour target (default 320) validation and assignment scoring
- Monthly evening maximum
- Swap Market
- Counter-offers only from shifts where the other pharmacist is not already working
- Direct acceptance without manager authorization
- Swap expiration after the posted shift date
- Schedule health dashboard
- Fairness report
- Audit log
- CSV export
- Finalize / publish schedule

The Google receiver workbook from the Apps Script version is replaced by:

```text
data/published-schedule.json
```

This is the GitHub-only equivalent of the finalized pharmacist schedule destination.

## Files

```text
index.html                      Main NeoChrono application
provision.html                  One-time browser/repository provisioning
import.html                     One-time legacy database import
css/app.css                     Application styling
js/app.js                       UI and application workflows
js/auth.js                      Username/password authentication
js/github-store.js              GitHub JSON database read/write layer
js/scheduler.js                 Scheduling engine and validation
js/schema.js                    Database schema/defaults/helpers
js/provision.js                 Provisioning workflow
js/import.js                    Legacy JSON import workflow
migration/ExportLegacyNeoChrono.gs
                                Optional one-time exporter from the old Google Sheet project
```

## Recommended repository setup

Use two repositories:

### 1. App repository

Example:

```text
neochrono-app
```

Contains this package and is served by GitHub Pages.

### 2. Data repository

Example:

```text
neochrono-data
```

Keep this repository private. NeoChrono creates:

```text
data/database.json
data/published-schedule.json
```

The main database is intentionally one JSON file. GitHub's file SHA acts as optimistic concurrency control, so NeoChrono detects competing saves instead of blindly overwriting another user's newer version.

## First-time setup — new blank database

1. Create the `neochrono-app` repository and upload this package.
2. Enable GitHub Pages for the app repository.
3. Create a private `neochrono-data` repository.
4. Create a fine-grained GitHub token restricted to the data repository with repository Contents read/write access.
5. Open:

```text
https://YOUR-GITHUB-PAGES-URL/provision.html
```

6. Enter:
   - GitHub owner
   - Data repository name
   - Branch, normally `main`
   - Fine-grained token
   - First administrator name
   - First administrator username
   - First administrator password
7. Click **Provision This Device**.
8. Open `index.html` and sign in using only username/password.

## Migrating the existing Google Sheets / Apps Script project

A one-time migration helper is included. The new runtime remains GitHub-only after migration.

### In the old project

1. Open the current NeoChrono Google Sheet.
2. Open Extensions → Apps Script.
3. Add a temporary script file.
4. Paste the contents of:

```text
migration/ExportLegacyNeoChrono.gs
```

5. Run:

```javascript
exportNeoChronoForGitHub()
```

6. Authorize it once.
7. It creates:

```text
NeoChrono_GitHub_Migration.json
```

in Google Drive.
8. Download that JSON file.

### In the new GitHub version

1. Provision the browser/repository using `provision.html`.
2. Open:

```text
import.html
```

3. Select `NeoChrono_GitHub_Migration.json`.
4. Confirm the import.
5. Return to `index.html` and log in normally.

### Existing usernames/passwords

The importer preserves legacy Apps Script SHA-256 password hashes. `auth.js` understands the legacy format. When an imported account logs in successfully for the first time, NeoChrono upgrades that account to PBKDF2-SHA256 automatically.

If a pharmacist did not have login credentials in the current project, an administrator can edit that pharmacist in **Pharmacists** and set a password.

## Scheduling configuration

The app does not invent missing shift times, credited hours, skills, or staffing numbers. Before generating a schedule, **Config Check** validates the configuration and blocks generation when important required information is missing.

Important defaults in this build follow the later work-pattern version of the existing project:

```text
Default Schedule Days: 56
Week Start: Sunday
Regular Workdays Per Week: 5
Maximum Consecutive Workdays: 5
Weekly Hours Limit: 40
Required Hours Per Schedule Period: 320
Weekend Rotation: A,B,C
Resident E2 Shifts Per Week: 1
PTO Auto-Approval Limit Per Date: 2
Paired ED Coverage: Yes
ED Day Shift: EDD
ED Evening Shift: EDE
```

All of these can be edited in the Settings page.

## Deployment updates

When you change HTML/CSS/JavaScript, commit the changes to the app repository. GitHub Pages will serve the updated static site.

Data changes made inside NeoChrono are committed directly to the private data repository through the GitHub API.

## Backups

GitHub commit history provides a history of every database update. Before a large migration or manual data edit, also create a branch/tag or download `data/database.json`.

## Testing included

`tests/test-scheduler.mjs` performs a basic scheduler and PTO auto-approval smoke test.

Run locally with a recent Node version:

```bash
node tests/test-scheduler.mjs
```
