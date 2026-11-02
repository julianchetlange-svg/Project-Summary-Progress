# McCoy's Project Summary Report — automated

`index.html` is the McCoy's Building Supply Store Development Project Summary Report, published with GitHub Pages. Every Thursday it rebuilds itself from two Fexa report emails sent to **mccoysfleet@gmail.com**.

## Weekly timeline (Thursday, Central time)

| Time | What happens |
|---|---|
| **5:00** | Fexa emails **2026 Project Report - Module** (`.xls`). |
| **5:10–6:40** | *Thursday 5am – module export* workflow reads it, cleans it, compares it with the published report and emails **chet.lange@mccoys.com** the projects that **closed** since last week, with budget, final cost, over/under and ready-to-paste contingency-log wording. It also lists other changes and anything that needs a look. If the email hasn't arrived by 6:40 you get an alert instead. |
| **5:00–11:00** | Chet updates the contingency log in Fexa. |
| **11:00** | Fexa emails **Current Budget Year Contingency** (`.xls`). |
| **11:05–11:25** | *Thursday 11am – contingency + publish* workflow merges it with the 5am data, rebuilds `index.html`, commits it to `main` (Pages republishes within a minute or two) and emails a short confirmation with the contingency changes. If the contingency email is late, it publishes at 11:25 with last week's log and says so. |

GitHub's scheduler can start runs a few minutes late at busy times. The schedule leaves room for that, but a delayed run can push publication past 11:30.

## Cleaning rules applied every week

- Projects with status **New** are removed. Project **4541** (508 Retreat fencing) is always removed.
- **Blank store**: use the store from last week's report for that project. If it's a new project, use the store whose 3-digit number appears in the project name or description. Each one is listed in the 5am email.
- **Blank assignee**: keep last week's PM and flag it. PM names are normalized to the seven standard names.
- **Springtown 7412 / Canyon Lake 3551**: restored from last week if Fexa drops them, labeled "2026‑27 carry". Springtown keeps its "(fka Reno)" name.
- **PENDING-071** (Nacogdoches asphalt) stays as a placeholder until a matching Fexa project appears. When one does, its notes move over and it's flagged.
- **Contingency**: the 11am log is authoritative. Project 4410's net drawn is set to the log's total, and the confirmation email shows any gap against the 5am Fexa figure. $0 lines and blank-name total rows are left out.
- **Carryover**: projects whose budget year spans two years, or that were already marked carryover, keep the yellow pill.
- **Flags, not silent fixes**: an Active project starting in a future year, a Closed project starting after the report date, an unknown PM, a duplicate ID, or an unrecognized status.

Column headers are detected by name pattern, and the header row is found automatically, so small layout changes in the Fexa export won't break it.

## Editing content by hand (on github.com)

- **Add budget meeting notes:** create `meeting-notes/YYYY-MM-DD.md` with the notes as plain text. The *Re-render* workflow rebuilds the page automatically, newest first.
- **Add or edit Budget Request / project notes:** edit the `notes` field for that project in `data/projects.json`, appending with `  ·  M/D/YY meeting: …`. It's carried forward every week and the page re-renders on save.
- **Change who gets emails, PM names or standing rules:** edit `config.json`.
- **KPI targets, owners, design:** these live in `index.html` itself and are never overwritten. The script only replaces the data blocks and the dates.

## Running it manually

Go to **Actions**, pick a workflow, then **Run workflow**. Enter the report date (`YYYY-MM-DD`) when you're catching up on an earlier Thursday. Use **force** to reprocess an email or republish a date.

## One-time setup

1. Turn on 2-Step Verification for **mccoysfleet@gmail.com**, then create an app password at <https://myaccount.google.com/apppasswords>.
2. In Gmail settings, under **Forwarding and POP/IMAP**, make sure **IMAP is enabled**.
3. In this repo go to **Settings → Secrets and variables → Actions → New repository secret** and add:
   - `GMAIL_USER` = `mccoysfleet@gmail.com`
   - `GMAIL_APP_PASSWORD` = the 16-character app password
4. Under **Settings → Actions → General → Workflow permissions**, choose **Read and write permissions**.

## Files

| Path | Purpose |
|---|---|
| `index.html` | The published report (self-contained; Chart.js from cdnjs). |
| `scripts/fexa_report.py` | All logic: Gmail fetch, cleaning, diff, render, emails. |
| `config.json` | Recipients, PM roster, standing removals/overrides, time windows. |
| `data/projects.json`, `data/contingency_log.json` | Data behind the last published report, including notes. |
| `data/staging/` | Thursday 5am results waiting for the 11am publish. |
| `data/snapshots/` | Per-week status/budget/cost per project, kept for future duration and on-time metrics. |
| `data/runs.json`, `data/meta.json` | Which emails were processed and when the report was last published. |
| `meeting-notes/` | One file per budget meeting. |

Raw Fexa exports are not committed (this repo is public); each run keeps them as a downloadable artifact for 90 days.

## Local test

```bash
pip install -r scripts/requirements.txt
python scripts/fexa_report.py module --file "2026 Project Report - Module.xls" --as-of 2026-11-05 --no-email
python scripts/fexa_report.py contingency --file "Current Budget Year Contingency.xls" --as-of 2026-11-05 --no-email
```
`--no-email` writes the emails to `out/` instead of sending them.
