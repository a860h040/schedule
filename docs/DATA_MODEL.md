# NeoChrono GitHub JSON data model

All runtime data is stored under the configured data folder, normally `data/`.

| File | Purpose |
|---|---|
| `meta.json` | Database/schema metadata |
| `admins.json` | NeoChrono administrator accounts and password hashes |
| `users.json` | Pharmacist scheduling profiles |
| `skills.json` | Normal scheduling skills |
| `shifts.json` | Shift definitions and credited hours |
| `staffing.json` | Required slot counts by weekday |
| `requests.json` | PTO, date availability, and PRN weekend availability |
| `weekly-availability.json` | Recurring weekly availability rules |
| `preceptor-calendar.json` | Per-date preceptor ON/OFF mode |
| `schedule.json` | Generated and manually edited assignments |
| `settings.json` | Scheduler configuration |
| `audit.json` | NeoChrono audit events |
| `swaps.json` | Swap-market posts and offers |

## Why the database is split

GitHub updates a repository file by replacing that file at a known SHA. Splitting the data reduces write conflicts: approving PTO changes `requests.json` without rewriting `schedule.json`, while editing a shift definition changes `shifts.json` without touching the audit history except for its separate append.

## Schedule row identity

Each schedule row has an `Assignment ID` UUID plus a `Generation ID`. Date + Shift + Slot identifies the required staffing slot. The calendar edits the existing `Assignment ID` instead of creating a duplicate slot.
