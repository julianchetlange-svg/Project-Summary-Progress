# McCoy's Project Summary Report — automated

`index.html` is the McCoy's Building Supply Store Development Project Summary Report, published with GitHub Pages. Every Thursday it rebuilds itself from two Fexa report emails sent to **mccoysfleet@gmail.com**.

## How the pieces fit

```
Fexa ──email──▶ mccoysfleet@gmail.com ──Apps Script (every 5 min)──▶ private repo Project-Summary-Inbox
                         ▲                                                   │ starts workflow
                         └──── sends emails left in outbox/ ◀── GitHub Actions (this repo) ──▶ index.html → GitHub Pages
```

No Gmail password is stored anywhere. The Apps Script runs inside the Gmail account, and one GitHub token (stored in the script and as a repo secret) lets the two sides talk. Raw Fexa exports stay in the **private** inbox repo; only the report goes to this public repo.

## Weekly timeline (Thursday, Central time)

| Time | What happens |
|---|---|
| **5:00** | Fexa emails **2026 Project Report - Module**. Within ~5 min the Apps Script uploads it and starts *Thursday 5am – module export*. |
| **~5:10** | The workflow cleans the export, compares it with the published report and queues an email to **chet.lange@mccoys.com**. That email lists projects that **closed** since last week, with budget, final cost, over/under and ready-to-paste contingency-log wording, plus other changes and anything that needs a look. Gmail sends it on the next 5-minute pass. If nothing arrives by 6:40 you get an alert instead. |
| **5:00–11:00** | Chet updates the contingency log in Fexa. |
| **11:00** | Fexa emails **Current Budget Year Contingency**; it's uploaded and *Thursday 11am – contingency + publish* starts. |
| **~11:10** | The report is rebuilt and committed to `main`, Pages republishes within a minute or two, and a short confirmation email (contingency changes) follows. If the contingency export hasn't arrived by 11:25, it publishes with last week's log and says so. |

Fexa exports that arrive on other days are uploaded but not processed automatically. To use one, run the workflow by hand with that date.

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

Go to **Actions**, pick a workflow, then **Run workflow**. Enter the report date (`YYYY-MM-DD`, the date the Fexa email arrived) when catching up or using an off-day export. Use **force** to reprocess an export or republish a date.

## One-time setup

1. **Create the private inbox repo.** On GitHub choose **New repository**, name it `Project-Summary-Inbox`, set it to **Private** and tick **Add a README file**.
2. **Create one access token.** Go to GitHub **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**:
   - Repository access: **Only select repositories**, then pick `Project-Summary-Progress` and `Project-Summary-Inbox`.
   - Permissions: **Contents → Read and write** (nothing else).
   - Expiration: the longest offered. Put a reminder on your calendar to renew it, since everything stops when it expires.
3. **Give the token to this repo.** In `Project-Summary-Progress` go to **Settings → Secrets and variables → Actions → New repository secret**, name it `INBOX_TOKEN` and paste the token. Then under **Settings → Actions → General → Workflow permissions**, choose **Read and write permissions**.
4. **Install the Gmail script.** Signed in as **mccoysfleet@gmail.com**, open <https://script.google.com> and choose **New project**:
   - Replace the code with the contents of `apps-script/Code.gs`.
   - Under **Project Settings → Script properties**, add `GITHUB_TOKEN` and paste the same token.
   - In the editor, pick `setup` and click **Run**, then approve the Gmail and external-request permissions. It checks the token, then starts the 5-minute schedule.

   Only Fexa emails received after `setup` runs are handed off.

## Files

| Path | Purpose |
|---|---|
| `index.html` | The published report (self-contained; Chart.js from cdnjs). |
| `scripts/fexa_report.py` | All logic: reading exports, cleaning, diff, render, composing emails. |
| `apps-script/Code.gs` | Runs in the mccoysfleet Gmail account: uploads Fexa attachments, sends queued emails. |
| `config.json` | Recipients, PM roster, standing removals/overrides, time windows. |
| `data/projects.json`, `data/contingency_log.json` | Data behind the last published report, including notes. |
| `data/staging/` | Thursday 5am results waiting for the 11am publish. |
| `data/snapshots/` | Per-week status/budget/cost per project, kept for future duration and on-time metrics. |
| `data/runs.json`, `data/meta.json` | Which exports were processed and when the report was last published. |
| `meeting-notes/` | One file per budget meeting. |

Raw Fexa exports and outgoing emails live in the private `Project-Summary-Inbox` repo, not here.

## Local test

```bash
pip install -r scripts/requirements.txt
python scripts/fexa_report.py module --file "2026 Project Report - Module.xls" --as-of 2026-11-05 --no-email
python scripts/fexa_report.py contingency --file "Current Budget Year Contingency.xls" --as-of 2026-11-05 --no-email
```
`--no-email` writes the emails to `out/` instead of queuing them in the inbox.
