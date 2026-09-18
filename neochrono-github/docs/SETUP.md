# Setup: GitHub Pages + private GitHub JSON database

## 1. Create the application repository

Create a repository such as `neochrono-app` and upload the contents of this project.

For GitHub Free, Pages is normally published from a public repository. If you do not want the app source public, use a GitHub plan that supports Pages from private repositories. The app source itself contains no live scheduling data or GitHub token.

## 2. Create the data repository

Create a separate repository such as `neochrono-data` and make it **private**.

It can be empty. The NeoChrono setup screen creates `data/*.json` automatically.

## 3. Create a fine-grained GitHub token

Create a fine-grained token that has access to **only the data repository**.

Required repository permission:

- Contents: Read and write

No GitHub token is placed in the code. The administrator enters it in the browser and it is stored in `sessionStorage` for the current browser session.

## 4. Enable GitHub Pages

In the app repository:

1. Open **Settings**.
2. Open **Pages**.
3. Choose **Deploy from a branch**.
4. Choose `main` and `/ (root)`.
5. Save.

Open the GitHub Pages URL after deployment completes.

## 5. Initialize NeoChrono

On the NeoChrono login screen:

1. Click **First-time database setup**.
2. Enter the GitHub owner of the private data repository.
3. Enter the repository name, normally `neochrono-data`.
4. Enter branch `main`.
5. Enter data folder `data`.
6. Paste the fine-grained token.
7. Create the first NeoChrono administrator account.
8. Click **Create GitHub database**.

The app creates all database JSON files as Git commits.

## 6. Complete scheduler configuration

Before generating a schedule, review:

- Pharmacists: Active, Schedule Type, Weekend Group, weekly maximum, target hours, preceptor/resident flags, preferred shift/home unit, OFF-day coverage skills, eligibility flags, and 7-on/7-off anchor.
- Skills: normal scheduling skills for each pharmacist.
- Shifts: start/end time, actual hours, credited hours, type, required skill, active state, and priority.
- Staffing Requirements: required slot count by shift and weekday.
- Weekly Availability.
- Preceptor Calendar.
- Settings.

NeoChrono blocks generation when required configuration is missing, such as an active 7-on/7-off pharmacist with no anchor or dedicated preferred shift.

## 7. Generate the first schedule

The default generation window is 56 days. For regular pharmacists, it must start on a Sunday and end on a Saturday.

Use **Generate new period** for dates with no existing schedule. Use **Regenerate existing period** when the period already exists; locked assignments are preserved.

After generation, review validation issues in the dashboard/calendar before final use.

## Security notes

- Never commit a personal access token.
- Keep the data repository private.
- Give the token the minimum possible repository access.
- Use separate GitHub accounts/tokens if you want GitHub's audit trail to identify different administrators.
- The browser-based NeoChrono password layer cannot provide stronger permissions than the GitHub token itself because there is no private server in a GitHub-only architecture.
