#!/usr/bin/env python3
"""
McCoy's Project Summary Report — weekly automation.

Inputs arrive through a private "inbox" repo (see apps-script/Code.gs); emails
go out the same way.

Commands
  module       Thursday ~5am CT. Read "2026 Project Report - Module",
               clean it, diff against the published report, stage the result and
               email Chet the closed projects with their over/under.
  contingency  Thursday ~11am CT. Read "Current Budget Year Contingency", merge
               with the staged module data, render index.html and record a
               snapshot. The workflow then commits and pushes.
  render       Re-render index.html from data/ (after hand edits to notes etc.).

Local/testing flags
  --file PATH      use a local .xls/.xlsx instead of the inbox
  --as-of DATE     override the report date (YYYY-MM-DD)
  --no-email       write emails to out/ instead of sending
  --scheduled      only act inside the configured Central-time window
  --force          reprocess an export that was already handled
  --thursday-only  (upload-triggered runs) ignore exports received on other days
"""
import argparse, datetime as dt, html, json, os, re, sys
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
CFG = json.loads((ROOT / "config.json").read_text())
TZ = ZoneInfo(CFG["timezone"])
OUT = ROOT / "out"

STATUSES = ["Active", "Pre-Construction", "Budget Plan", "Sourcing",
            "Post-Construction", "Closed", "Budget Request"]
BASE_KEYS = ["id", "status", "store", "region", "assignee", "budget", "cost", "type",
             "name", "start", "effortDays", "pct", "carryover"]
EXTRA_KEYS = ["carryLabel", "requestDate", "notes"]
REGION_RE = re.compile(r"Region\s*0?(\d)\b", re.I)
FIXED_REGIONS = ["Headquarters", "Distribution", "Millwork"]


# ───────────────────────── small helpers ─────────────────────────
def now_ct():
    return dt.datetime.now(TZ)

def load(p, default=None):
    p = DATA / p
    return json.loads(p.read_text()) if p.exists() else default

def save(p, obj):
    p = DATA / p
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, indent=1, ensure_ascii=False) + "\n")

def clean(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def clean_name(s):
    s = clean(s)
    s = re.sub(r"(?<=\w)- (?=\S)", " - ", s)     # "Pharr- Sitework" -> "Pharr - Sitework"
    s = re.sub(r"(?<=\S) -(?=\w)", " - ", s)     # "Baytown -Split"  -> "Baytown - Split"
    return s

def num(v):
    try:
        f = float(str(v).replace("$", "").replace(",", ""))
        if pd.isna(f):
            return 0.0
        return int(f) if f.is_integer() else round(f, 2)
    except (TypeError, ValueError):
        return 0.0

def money(v, signed=False):
    s = f"${abs(v):,.2f}"
    if v < 0:
        return "-" + s
    return ("+" + s) if signed and v > 0 else s

def pid_key(i):
    return str(i)

def fexa_url(i):
    return CFG["fexa_project_url"].format(id=i)

def esc(s):
    return html.escape(str(s))

def in_window(win_start, win_end, t=None):
    t = t or now_ct()
    if t.weekday() != 3:                      # Thursday
        return False
    hm = t.strftime("%H:%M")
    return win_start <= hm <= win_end


# ───────────────────────── reading Fexa exports ─────────────────────────
def read_sheet(path, header_words):
    """Read an export whose header row is not necessarily row 1."""
    raw = pd.read_excel(path, header=None, dtype=object)
    for i in range(min(15, len(raw))):
        row = " | ".join(clean(x).lower() for x in raw.iloc[i].tolist())
        if any(w in row for w in header_words):
            df = raw.iloc[i + 1:].copy()
            df.columns = [clean(c) or f"Unnamed: {j}" for j, c in enumerate(raw.iloc[i].tolist())]
            return df.reset_index(drop=True)
    raise ValueError(f"No header row found in {path} (looked for {header_words})")

def find_col(df, include, exclude=None, required=True):
    for c in df.columns:
        lc = c.lower()
        if re.search(include, lc) and not (exclude and re.search(exclude, lc)):
            return c
    if required:
        raise KeyError(f"Column matching /{include}/ not found. Columns: {list(df.columns)}")
    return None


class PMs:
    def __init__(self, names):
        self.full = {n.lower(): n for n in names}
        self.last = {n.split()[-1].lower(): n for n in names}
        self.first = {}
        for n in names:
            self.first.setdefault(n.split()[0].lower(), []).append(n)

    def norm(self, s):
        s = clean(s)
        if not s or s.lower() in ("nan", "nan nan", "unassigned"):
            return ""
        if s.lower() in self.full:
            return self.full[s.lower()]
        parts = s.replace(",", " ").split()
        for p in reversed(parts):
            if p.lower() in self.last:
                return self.last[p.lower()]
        if len(parts) == 1 and len(self.first.get(parts[0].lower(), [])) == 1:
            return self.first[parts[0].lower()][0]
        return s

    def score(self, series):
        vals = [clean(v) for v in series if clean(v)]
        return sum(self.norm(v) in self.full.values() for v in vals) / max(len(vals), 1)


def detect_module_columns(df, pms):
    c = {
        "id": find_col(df, r"project id|^id$"),
        "status": find_col(df, r"status"),
        "region": find_col(df, r"district|region", required=False),
        "store": find_col(df, r"store", r"district|region"),
        "budget": find_col(df, r"budget", r"variance|year|percent"),
        "cost": find_col(df, r"total cost|actual cost|^cost", r"projected"),
        "type": find_col(df, r"type"),
        "budget_year": find_col(df, r"budget year", required=False),
        "name": find_col(df, r"project name|^name$"),
        "desc": find_col(df, r"description", required=False),
        "start": find_col(df, r"start"),
        "effort": find_col(df, r"effort|days", required=False),
        "pct": find_col(df, r"completion|percent complete|% complete", r"date", required=False),
        "first": find_col(df, r"first name", required=False),
        "last": find_col(df, r"last name", required=False),
        "assignee": find_col(df, r"assi[gn]+ee|assigned", r"first|last", required=False),
    }
    if not (c["first"] and c["last"]) and not c["assignee"]:
        # header drift: pick whichever column best matches the PM roster
        best = max(df.columns, key=lambda col: pms.score(df[col]))
        if pms.score(df[best]) > 0.5:
            c["assignee"] = best
        else:
            raise KeyError("Could not find an assignee column")
    return c


def store_directory(*project_lists):
    """3-digit store code -> (store label, region) from everything we know."""
    d = {}
    for plist in project_lists:
        for p in plist:
            m = re.match(r"(\d{3})\b", p.get("store") or "")
            if m and p.get("store"):
                d.setdefault(m.group(1), (p["store"], p.get("region") or "Unknown"))
    return d


def norm_region(raw, fallback):
    r = clean(raw)
    m = REGION_RE.search(r)
    if m:
        return f"Region 0{m.group(1)}"
    for f in FIXED_REGIONS:
        if f.lower() in r.lower():
            return f
    return fallback or "Unknown"


def parse_module(path, prior, as_of):
    pms = PMs(CFG["project_managers"])
    df = read_sheet(path, ["project id"])
    col = detect_module_columns(df, pms)
    prior_by = {pid_key(p["id"]): p for p in prior}
    flags, removed, inferred = [], [], []
    rows = []

    # first pass: raw records (to build a store directory from this export too)
    for _, r in df.iterrows():
        rid = clean(r[col["id"]])
        if not rid:
            continue
        try:
            rid = int(float(rid))
        except ValueError:
            pass
        rows.append((rid, r))
    export_stores = [{"store": clean(r[col["store"]]),
                      "region": norm_region(r[col["region"]] if col["region"] else "", "")}
                     for _, r in rows]
    stores = store_directory(prior, export_stores)

    out, seen = [], set()
    for rid, r in rows:
        status = clean(r[col["status"]])
        status = next((s for s in STATUSES if s.lower() == status.lower()), status)
        name = clean_name(r[col["name"]])
        if status in CFG["remove_statuses"]:
            removed.append(f"{rid} {name} (status {status})")
            continue
        if rid in CFG["remove_project_ids"]:
            removed.append(f"{rid} {name} (standing removal)")
            continue
        if rid in seen:
            flags.append(f"Duplicate project ID {rid} in export; kept the first row.")
            continue
        seen.add(rid)
        pr = prior_by.get(pid_key(rid), {})

        # assignee
        if col["first"] and col["last"]:
            a_raw = f"{clean(r[col['first']])} {clean(r[col['last']])}"
        else:
            a_raw = clean(r[col["assignee"]])
        assignee = pms.norm(a_raw)
        if not assignee:
            assignee = pr.get("assignee") or "Unassigned"
            flags.append(f"{rid} {name}: no assignee in Fexa — "
                         + (f"kept prior PM {assignee}." if pr.get("assignee") else "shown as Unassigned."))
        elif assignee not in CFG["project_managers"]:
            flags.append(f"{rid} {name}: assignee '{assignee}' is not one of the 7 PMs "
                         "(add them to config.json and the avC/shA maps in index.html).")

        # store + region
        store = clean(r[col["store"]])
        region_raw = r[col["region"]] if col["region"] else ""
        region = norm_region(region_raw, pr.get("region") or stores.get(store[:3], ("", ""))[1])
        if not store:
            desc = clean(r[col["desc"]]) if col["desc"] else ""
            if pr.get("store"):
                store, region = pr["store"], pr.get("region", region)
                how = "kept the store from the prior report"
            else:
                code = next((m for m in re.findall(r"\b(\d{3})\b", f"{name} {desc}") if m in stores), None)
                if code:
                    store, region = stores[code]
                    how = f"matched store number {code} in the project name/description"
                else:
                    lead = re.match(r"(\d{3}\s+[^-]+?)\s+-", name)
                    store = lead.group(1) if lead else "Unknown"
                    how = "used the project name prefix" if lead else "could not determine — please set in Fexa"
            inferred.append(f"{rid} {name} → store '{store}', {region} ({how}).")

        start = pd.to_datetime(r[col["start"]], errors="coerce")
        start = start.strftime("%Y-%m-%d") if not pd.isna(start) else pr.get("start", "")
        by = clean(r[col["budget_year"]]) if col["budget_year"] else ""
        multi_year = bool(by) and by != str(CFG["report_year"])

        p = {
            "id": rid, "status": status, "store": store, "region": region,
            "assignee": assignee, "budget": num(r[col["budget"]]), "cost": num(r[col["cost"]]),
            "type": clean(r[col["type"]]) or pr.get("type", "Other"), "name": name, "start": start,
            "effortDays": int(num(r[col["effort"]])) if col["effort"] else pr.get("effortDays", 0),
            "pct": int(num(r[col["pct"]])) if col["pct"] else pr.get("pct", 0),
            "carryover": multi_year or bool(pr.get("carryover")),
        }
        if by and re.match(rf"{CFG['report_year']}\s*-\s*{CFG['report_year'] + 1}", by):
            p["carryLabel"] = f"{CFG['report_year']}‑{str(CFG['report_year'] + 1)[2:]} carry"
        for k in EXTRA_KEYS:                          # carry forward notes, request dates, labels
            if pr.get(k) and k not in p:
                p[k] = pr[k]
        if status == "Budget Request" and not p.get("requestDate"):
            p["requestDate"] = as_of
        p.update({k: v for k, v in CFG["overrides"].get(str(rid), {}).items() if not k.startswith("_")})
        out.append(p)

        # data-quality checks (after overrides, so corrected values aren't re-flagged)
        start = p["start"]
        if status == "Active" and start and int(start[:4]) > CFG["report_year"]:
            flags.append(f"{rid} {name}: Active but planned start is {start}.")
        if status == "Closed" and start and start > as_of:
            flags.append(f"{rid} {name}: Closed but planned start {start} is after the report date.")
        if not re.match(r"\d{3}", name) and status != "Budget Request":
            flags.append(f"{rid} '{name}': project name has no store number.")

    ids = {pid_key(p["id"]) for p in out}

    # manual placeholders (e.g. PENDING-071) and multi-year builds the export may drop
    for pr in prior:
        k = pid_key(pr["id"])
        if k in ids:
            continue
        if not isinstance(pr["id"], int):
            code = (re.match(r"(\d{3})", pr.get("store", "")) or [None, None])[1]
            match = next((p for p in out if code and p["store"].startswith(code)
                          and pr["type"].lower() in p["name"].lower()
                          and any(w in p["name"].lower() for w in pr["name"].lower().split()[-2:])), None)
            if match:
                for kx in ("notes", "requestDate"):
                    if pr.get(kx) and not match.get(kx):
                        match[kx] = pr[kx]
                flags.append(f"Placeholder {pr['id']} looks like it now has Fexa ID {match['id']} "
                             f"({match['name']}); replaced it and moved its notes. Please confirm.")
            else:
                out.append(dict(pr))
        elif pr["id"] in CFG["restore_if_missing"]:
            out.append(dict(pr))
            flags.append(f"{pr['id']} {pr['name']} was missing from the export — restored last known figures.")

    # keep the prior report's ordering; new projects go to the end of their status group
    order = {pid_key(p["id"]): i for i, p in enumerate(prior)}
    cid = CFG["contingency_project_id"]
    out.sort(key=lambda p: (p["id"] != cid,
                            STATUSES.index(p["status"]) if p["status"] in STATUSES else 99,
                            order.get(pid_key(p["id"]), 10_000)))
    unknown_status = sorted({p["status"] for p in out if p["status"] not in STATUSES})
    for s in unknown_status:
        flags.append(f"Unrecognised status '{s}' in export — the report has no badge colour for it.")
    return out, flags, removed, inferred


def parse_contingency(path):
    df = read_sheet(path, ["task name", "line item", "name"])
    c_name = find_col(df, r"name|line item")
    c_det = find_col(df, r"description|detail", required=False)
    c_amt = find_col(df, r"amount|cost", r"projected")
    c_proj = find_col(df, r"projected", required=False)
    c_dept = find_col(df, r"department", required=False)
    items, dropped = [], []
    for _, r in df.iterrows():
        nm, amt = clean(r[c_name]), num(r[c_amt])
        if not nm:
            if amt:
                dropped.append(f"blank-name row {money(amt)} (export total row)")
            continue
        if amt == 0:
            dropped.append(f"'{nm}' ($0 line)")
            continue
        items.append({"name": nm, "detail": clean(r[c_det]) if c_det else "",
                      "department": clean(r[c_dept]) or None if c_dept else None,
                      "projected": num(r[c_proj]) or None if c_proj else None,
                      "amount": amt})
    return items, dropped


# ───────────────────────── diffing ─────────────────────────
def diff_projects(prior, new):
    pb = {pid_key(p["id"]): p for p in prior}
    nb = {pid_key(p["id"]): p for p in new}
    d = {"status": [], "new": [], "removed": [], "cost": [], "budget": [], "assignee": [],
         "closed": [], "post": []}
    for k, p in nb.items():
        o = pb.get(k)
        if not o:
            d["new"].append(p)
            if p["status"] == "Closed":
                d["closed"].append(p)
            continue
        if o["status"] != p["status"]:
            d["status"].append((p, o["status"], p["status"]))
            if p["status"] == "Closed":
                d["closed"].append(p)
            elif p["status"] == "Post-Construction":
                d["post"].append(p)
        if abs(o["cost"] - p["cost"]) > 0.005:
            d["cost"].append((p, o["cost"], p["cost"]))
        if abs(o["budget"] - p["budget"]) > 0.005:
            d["budget"].append((p, o["budget"], p["budget"]))
        if o["assignee"] != p["assignee"]:
            d["assignee"].append((p, o["assignee"], p["assignee"]))
    d["removed"] = [p for k, p in pb.items() if k not in nb]
    return d


def log_key(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def diff_log(prior, new):
    pk = {}
    for l in prior:
        pk.setdefault(log_key(l["name"]), []).append(l)
    added, changed = [], []
    seen = set()
    for l in new:
        cands = pk.get(log_key(l["name"]), [])
        o = next((c for c in cands if id(c) not in seen), None)
        if not o:
            added.append(l)
            continue
        seen.add(id(o))
        if abs(o["amount"] - l["amount"]) > 0.005:
            changed.append((l, o["amount"], l["amount"]))
    removed = [l for l in prior if id(l) not in seen]
    return added, changed, removed


# ───────────────────────── rendering index.html ─────────────────────────
def js_array(items, grouped=False):
    def lit(o):
        return json.dumps(o, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    if not grouped:
        return "[\n" + ",\n".join("  " + lit(o) for o in items) + "\n]"
    lines, cur = [], None
    for o in items:
        grp = "CONTINGENCY" if o["id"] == CFG["contingency_project_id"] else o["status"].upper()
        if grp != cur:
            lines.append(f"\n  /* {grp} */")
            cur = grp
        ordered = {k: o[k] for k in BASE_KEYS if k in o}
        ordered.update({k: v for k, v in o.items() if k not in ordered})
        lines.append("  " + lit(ordered) + ",")
    if lines:
        lines[-1] = lines[-1].rstrip(",")
    return "[" + "\n".join(lines) + "\n]"


def meeting_notes():
    notes = []
    for f in sorted((ROOT / "meeting-notes").glob("*.md"), reverse=True):
        if re.match(r"\d{4}-\d{2}-\d{2}\.md$", f.name):
            notes.append({"date": f.stem, "text": f.read_text().strip()})
    return notes


def render(projects, log, as_of):
    path = ROOT / "index.html"
    h = path.read_text()
    d = dt.date.fromisoformat(as_of)
    long_d = d.strftime("%B ") + str(d.day) + d.strftime(", %Y")
    short_d = d.strftime("%b ") + str(d.day) + d.strftime(", %Y")

    def sub(pattern, repl, count_expected):
        nonlocal h
        h, n = re.subn(pattern, lambda m: repl, h)
        if n != count_expected:
            raise RuntimeError(f"index.html template changed: /{pattern[:40]}/ matched {n}x, expected {count_expected}")

    sub(r"const projects = \[[\s\S]*?\n\];", "const projects = " + js_array(projects, True) + ";", 1)
    sub(r"const contingencyLog = \[[\s\S]*?\n\];", "const contingencyLog = " + js_array(log) + ";", 1)
    sub(r"const meetingNotes = \[[\s\S]*?\n\];", "const meetingNotes = " + js_array(meeting_notes()) + ";", 1)
    sub(r"<title>McCoy's Project Summary — [^<]*</title>", f"<title>McCoy's Project Summary — {short_d}</title>", 1)
    sub(r'<div class="report-date">As of [^<]*</div>', f'<div class="report-date">As of {long_d}</div>', 1)
    sub(r"PROJECT DATA — week of [^\n]*", f"PROJECT DATA — week of {short_d}", 1)
    h, n = re.subn(r"const today=new Date\('\d{4}-\d{2}-\d{2}'\)", f"const today=new Date('{as_of}')", h)
    if n < 1:
        raise RuntimeError("pacing 'today' date not found in index.html")
    path.write_text(h)


# ─────────────── private inbox repo (files in, emails out) ───────────────
# A Google Apps Script in the mccoysfleet@gmail.com account uploads each Fexa
# attachment to <inbox>/incoming/<kind>/<YYYY-MM-DD>_<HHmmss>_<gmail id>.<ext>
# and sends any email this script leaves in <inbox>/outbox/.
INBOX = Path(os.environ.get("INBOX_DIR", ROOT / "inbox"))


def find_export(kind, as_of, days=6):
    """Newest uploaded attachment of this kind received on the report date or up to `days` after it
    (so a late or resent Fexa email still counts toward that Thursday's report)."""
    d = INBOX / "incoming" / kind
    last = (dt.date.fromisoformat(as_of) + dt.timedelta(days=days)).isoformat()
    files = sorted(f for f in d.glob("*_*") if f.suffix.lower() in (".xls", ".xlsx", ".csv")
                   and as_of <= f.name[:10] <= last) if d.exists() else []
    return files[-1] if files else None


def send(subject, html_body, no_email=False):
    slug = re.sub(r"[^\w]+", "_", subject)[:60]
    if no_email:
        OUT.mkdir(exist_ok=True)
        f = OUT / f"{slug}.html"
        f.write_text(html_body)
        print(f"[no-email] wrote {f}")
        return
    box = INBOX / "outbox"
    box.mkdir(parents=True, exist_ok=True)
    f = box / f"{now_ct():%Y%m%d-%H%M%S}_{slug}.json"
    f.write_text(json.dumps({"to": CFG["notify_to"], "subject": subject, "html": html_body}, ensure_ascii=False))
    print(f"queued email for Gmail: {subject}")


# ───────────────────────── email bodies ─────────────────────────
CSS = ("font-family:-apple-system,Segoe UI,Arial,sans-serif;font-size:13px;color:#222;")
TH = "background:#1B5C2E;color:#fff;text-align:left;padding:6px 8px;font-size:11px;text-transform:uppercase;"
TD = "padding:5px 8px;border-bottom:1px solid #e3e6e0;vertical-align:top;"

def table(headers, rows, right=()):
    h = "".join(f'<th style="{TH}{"text-align:right;" if i in right else ""}">{esc(x)}</th>' for i, x in enumerate(headers))
    body = "".join("<tr>" + "".join(
        f'<td style="{TD}{"text-align:right;white-space:nowrap;" if i in right else ""}">{c}</td>'
        for i, c in enumerate(r)) + "</tr>" for r in rows)
    return f'<table style="border-collapse:collapse;width:100%;margin:6px 0 16px;">{h}{body}</table>'

def plink(p):
    if isinstance(p["id"], int):
        return f'<a href="{fexa_url(p["id"])}" style="color:#1B5C2E;font-weight:600;">{p["id"]}</a>'
    return esc(p["id"])

def var_cell(budget, cost):
    v = budget - cost
    col = "#1B5C2E" if v >= 0 else "#c0392b"
    word = "under" if v >= 0 else "over"
    pct = f" ({abs(v) / budget * 100:.0f}%)" if budget else ""
    return f'<span style="color:{col};font-weight:600;">{money(abs(v))} {word}{pct}</span>'

def logged_for(p, log):
    """Contingency-log lines already booked for this project (matched on name)."""
    k = log_key(p["name"])
    return [l for l in log if log_key(l["name"]) == k]


def contingency_line(p, log=()):
    over = p["cost"] - p["budget"]                     # + = over budget (draw), - = under (return)
    booked = sum(l["amount"] for l in logged_for(p, log))
    word = "Over" if over > 0 else "Under"
    base = (f"Original Budget {money(p['budget'])} - Actual Spend {money(p['cost'])} - "
            f"{money(abs(over))} {word} Budget")
    if not booked:
        return base, over
    adj = round(over - booked, 2)
    return (f"{base} - {money(booked, True)} already in log - "
            f"{'no further entry needed' if abs(adj) < 0.01 else 'add ' + money(adj, True)}"), adj

def h2(t):
    return f'<h3 style="color:#1B5C2E;margin:18px 0 4px;font-size:15px;">{esc(t)}</h3>'

def bullets(items):
    return "<ul style='margin:4px 0 12px;padding-left:18px;'>" + "".join(f"<li>{x}</li>" for x in items) + "</ul>"


def module_email(d, flags, removed, inferred, as_of, prior_as_of, log=()):
    parts = [f'<div style="{CSS}">',
             f'<div style="background:#1B5C2E;color:#F5C400;padding:10px 14px;font-weight:700;font-size:16px;">'
             f'McCoy\'s Project Summary — weekly closeout check ({as_of})</div>',
             f'<p>Compared with the report published {prior_as_of}. The full report will be rebuilt and pushed '
             f'to GitHub after the 11am <b>Current Budget Year Contingency</b> email arrives — update the '
             f'contingency log in Fexa before then.</p>']

    parts.append(h2(f"Closed since last report ({len(d['closed'])})"))
    if d["closed"]:
        rows = []
        for p in d["closed"]:
            text, adj = contingency_line(p, log)
            booked = logged_for(p, log)
            note = "".join(f'<div style="font-size:11px;color:#7a5500;">Already in log: {esc(l["detail"] or l["name"])} '
                           f'({money(l["amount"], True)})</div>' for l in booked)
            rows.append([plink(p), esc(p["name"]) + (' <span style="background:#F5C400;font-size:10px;padding:1px 5px;'
                         'border-radius:8px;">carryover — includes prior-year spend</span>' if p["carryover"] else ""),
                         esc(p["assignee"]), money(p["budget"]), money(p["cost"]), var_cell(p["budget"], p["cost"]),
                         f'<b>{money(adj, True)}</b>' if abs(adj) >= 0.01 else "—",
                         f'<span style="font-size:11px;color:#555;">{esc(text)}</span>{note}'])
        tb, tc = sum(p["budget"] for p in d["closed"]), sum(p["cost"] for p in d["closed"])
        rows.append(["", "<b>Total</b>", "", f"<b>{money(tb)}</b>", f"<b>{money(tc)}</b>", var_cell(tb, tc), "", ""])
        parts.append(table(["ID", "Project", "PM", "Budget", "Final cost", "Over / under", "Log entry (+draw / −return)",
                            "Contingency log wording"], rows, right=(3, 4, 6)))
        parts.append('<p style="font-size:11px;color:#555;">Log entry = over/under minus anything already booked in '
                     'the contingency log for that project. Carryover projects: enter only the 2026 portion.</p>')
    else:
        parts.append("<p>No projects moved to Closed this week.</p>")

    if d["post"]:
        parts.append(h2(f"Moved to Post-Construction ({len(d['post'])}) — costs may still change"))
        parts.append(table(["ID", "Project", "PM", "Budget", "Cost to date", "Over / under"],
                           [[plink(p), esc(p["name"]), esc(p["assignee"]), money(p["budget"]), money(p["cost"]),
                             var_cell(p["budget"], p["cost"])] for p in d["post"]], right=(3, 4)))

    other = []
    for p, a, b in d["status"]:
        if p not in d["closed"] and p not in d["post"]:
            other.append(f"{plink(p)} {esc(p['name'])}: status {esc(a)} → <b>{esc(b)}</b>")
    for p in d["new"]:
        other.append(f"{plink(p)} {esc(p['name'])}: <b>new</b> ({esc(p['status'])}, {money(p['budget'])}, {esc(p['assignee'])})")
    for p in d["removed"]:
        other.append(f"{esc(p['id'])} {esc(p['name'])}: <b>no longer in the export</b> (was {esc(p['status'])})")
    for p, a, b in d["budget"]:
        other.append(f"{plink(p)} {esc(p['name'])}: budget {money(a)} → <b>{money(b)}</b>")
    for p, a, b in d["cost"]:
        other.append(f"{plink(p)} {esc(p['name'])}: cost {money(a)} → <b>{money(b)}</b>")
    for p, a, b in d["assignee"]:
        other.append(f"{plink(p)} {esc(p['name'])}: assignee {esc(a)} → <b>{esc(b)}</b>")
    parts.append(h2(f"Other changes ({len(other)})"))
    parts.append(bullets(other) if other else "<p>None.</p>")

    if removed:
        parts.append(h2("Removed by standing rules"))
        parts.append(bullets([esc(x) for x in removed]))
    if inferred:
        parts.append(h2("Stores filled in (blank in Fexa)"))
        parts.append(bullets([esc(x) for x in inferred]))
    if flags:
        parts.append(h2("Needs a look"))
        parts.append(bullets([esc(x) for x in flags]))
    parts.append("</div>")
    return "".join(parts)


def publish_email(as_of, log_added, log_changed, log_removed, ledger, log_net, notes, cfg_budget):
    parts = [f'<div style="{CSS}">',
             f'<div style="background:#1B5C2E;color:#F5C400;padding:10px 14px;font-weight:700;font-size:16px;">'
             f'Project Summary Report published — {as_of}</div>']
    parts.append(f"<p>Contingency remaining: <b>{money(cfg_budget - log_net)}</b> "
                 f"(budget {money(cfg_budget)}, net drawn {money(log_net)}).</p>")
    if ledger is not None and abs(ledger - log_net) > 1:
        parts.append(f'<p style="color:#7a5500;">The 5am Fexa ledger for project {CFG["contingency_project_id"]} '
                     f'showed {money(ledger)}; the 11am contingency log nets to {money(log_net)} '
                     f'(difference {money(log_net - ledger, True)}). The report uses the contingency log.</p>')
    lines = ([f"<b>Added</b>: {esc(l['name'])} — {money(l['amount'], True)}" for l in log_added] +
             [f"<b>Changed</b>: {esc(l['name'])} — {money(a)} → {money(b)}" for l, a, b in log_changed] +
             [f"<b>Removed</b>: {esc(l['name'])} — {money(l['amount'])}" for l in log_removed])
    parts.append(h2("Contingency log changes"))
    parts.append(bullets(lines) if lines else "<p>No changes to the contingency log this week.</p>")
    if notes:
        parts.append(h2("Notes"))
        parts.append(bullets([esc(n) for n in notes]))
    parts.append("</div>")
    return "".join(parts)


# ───────────────────────── commands ─────────────────────────
def get_input(args, kind, as_of):
    """Returns (path, key, received), "done" if already handled, or None if not uploaded yet."""
    if args.file:
        return Path(args.file), f"local:{Path(args.file).name}", now_ct().isoformat()
    f = find_export(kind, as_of)
    if not f:
        return None
    runs = load("runs.json", {"processed": {}})
    if f.name in runs["processed"] and not args.force:
        print(f"Already processed {f.name} — nothing to do.")
        return "done"
    return f, f.name, f.name[:10]


def mark_processed(mid, kind, received):
    runs = load("runs.json", {"processed": {}})
    runs["processed"][mid] = {"kind": kind, "received": received, "handled": now_ct().isoformat()}
    save("runs.json", runs)


def gh_output(**kv):
    f = os.environ.get("GITHUB_OUTPUT")
    for k, v in kv.items():
        print(f"{k}={v}")
        if f:
            with open(f, "a") as fh:
                fh.write(f"{k}={v}\n")


def cmd_module(args):
    as_of = args.as_of or now_ct().date().isoformat()
    w = CFG["module_window"]
    if args.scheduled and not in_window(w["start"], w["end"]):
        print("Outside the Thursday module window — skipping.")
        return
    if args.thursday_only and dt.date.fromisoformat(as_of).weekday() != 3:
        print(f"{as_of} is not a Thursday — file kept in the inbox; run the workflow manually with this date to use it.")
        return
    got = get_input(args, "module", as_of)
    if got == "done":
        return
    if not got:
        if args.scheduled and now_ct().strftime("%H:%M") >= w["alert_after"] and not load("staging/alerted_module.json"):
            send(f"Fexa module export not received ({as_of})",
                 f'<div style="{CSS}"><p>The <b>2026 Project Report - Module</b> email from Fexa had not arrived '
                 f'in {CFG["mailbox"]} (or had not been uploaded) by {w["alert_after"]} CT. The 11am run will still publish using last '
                 f'week\'s project data plus the new contingency log. To catch up, run the <i>Thursday 5am – '
                 f'module export</i> workflow manually once the email arrives.</p></div>', no_email=args.no_email)
            save("staging/alerted_module.json", {"as_of": as_of})
            gh_output(changed="true", message=f"Module export missing alert {as_of}")
        else:
            print("Module email not found yet — will retry on the next scheduled run.")
        return
    path, mid, received = got
    prior = load("projects.json")
    meta = load("meta.json", {"as_of": "2026-10-23"})
    projects, flags, removed, inferred = parse_module(path, prior, as_of)
    d = diff_projects(prior, projects)
    cid = CFG["contingency_project_id"]
    ledger = next((p["cost"] for p in projects if p["id"] == cid), None)
    save("staging/module.json", {"as_of": as_of, "source": path.name,
                                 "message_id": mid, "received": received, "ledger_net_drawn": ledger,
                                 "flags": flags, "removed": removed, "inferred": inferred, "projects": projects})
    # draft preview: today's project data + last week's contingency log
    body = module_email(d, flags, removed, inferred, as_of, meta["as_of"], load("contingency_log.json", []))
    n = len(d["closed"])
    send(f"Projects closed this week: {n} — McCoy's report {as_of}", body, no_email=args.no_email)
    mark_processed(mid, "module", received)
    gh_output(changed="true", message=f"Stage module export {as_of}: {n} closed, {len(flags)} flags")


def cmd_contingency(args):
    as_of = args.as_of or now_ct().date().isoformat()
    w = CFG["contingency_window"]
    if args.scheduled and not in_window(w["start"], w["end"]):
        print("Outside the Thursday contingency window — skipping.")
        return
    if args.thursday_only and dt.date.fromisoformat(as_of).weekday() != 3:
        print(f"{as_of} is not a Thursday — file kept in the inbox; run the workflow manually with this date to use it.")
        return
    meta = load("meta.json", {"as_of": "2026-10-23"})
    if args.scheduled and meta.get("as_of") == as_of and not args.force:
        print(f"Report already published for {as_of}.")
        return
    got = get_input(args, "contingency", as_of)
    if got == "done":
        return
    notes = []
    prior_log = load("contingency_log.json")
    if not got:
        if args.scheduled and now_ct().strftime("%H:%M") < w["publish_by"]:
            print("Contingency email not found yet — will retry.")
            return
        notes.append("The Current Budget Year Contingency email had not arrived — published with last week's "
                     "contingency log. Re-run the 11am workflow manually once it arrives.")
        log, mid, received = prior_log, None, None
    else:
        path, mid, received = got
        log, dropped = parse_contingency(path)
        if dropped:
            notes.append("Left out of the log: " + "; ".join(dropped) + ".")

    stage = load("staging/module.json")
    if stage and stage["as_of"] == as_of:
        projects, ledger = stage["projects"], stage.get("ledger_net_drawn")
        notes += [f"Data check: {f}" for f in stage["flags"]]
    else:
        projects, ledger = load("projects.json"), None
        notes.append("No module export was staged for this date — project figures are from the last report.")

    cid = CFG["contingency_project_id"]
    log_net = round(sum(l["amount"] for l in log), 2)
    cproj = next((p for p in projects if p["id"] == cid), None)
    if cproj is None:
        raise SystemExit(f"Contingency project {cid} missing from project data.")
    cproj["cost"] = log_net                      # contingency log is authoritative over the ledger
    added, changed, removed_l = diff_log(prior_log, log)
    groups = {}
    for l in log:
        groups.setdefault(log_key(l["name"]), []).append(l)
    for g in groups.values():
        if len(g) > 1 and any(l in added for l in g):
            notes.append(f"'{g[0]['name']}' now has {len(g)} lines in the contingency log "
                         f"({', '.join(money(l['amount'], True) for l in g)}; net {money(sum(l['amount'] for l in g), True)}). "
                         "Check it isn't counted twice.")

    render(projects, log, as_of)
    save("projects.json", projects)
    save("contingency_log.json", log)
    save(f"snapshots/{as_of}.json", {pid_key(p["id"]): {"status": p["status"], "budget": p["budget"],
                                     "cost": p["cost"], "assignee": p["assignee"]} for p in projects})
    save("meta.json", {"as_of": as_of, "published": now_ct().isoformat(),
                       "module_message": stage.get("message_id") if stage else None,
                       "contingency_message": mid})
    for f in ("staging/module.json", "staging/alerted_module.json"):
        (DATA / f).unlink(missing_ok=True)
    if mid:
        mark_processed(mid, "contingency", received)
    send(f"Project Summary Report published — {as_of}",
         publish_email(as_of, added, changed, removed_l, ledger, log_net, notes, cproj["budget"]),
         no_email=args.no_email)
    gh_output(changed="true", message=f"Project Summary Report {as_of}")


def cmd_render(args):
    meta = load("meta.json", {"as_of": "2026-10-23"})
    render(load("projects.json"), load("contingency_log.json"), args.as_of or meta["as_of"])
    gh_output(changed="true", message=f"Re-render report {args.as_of or meta['as_of']}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["module", "contingency", "render"])
    ap.add_argument("--file")
    ap.add_argument("--as-of")
    ap.add_argument("--no-email", action="store_true")
    ap.add_argument("--scheduled", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--thursday-only", action="store_true", help="ignore uploads not received on a Thursday")
    a = ap.parse_args()
    {"module": cmd_module, "contingency": cmd_contingency, "render": cmd_render}[a.command](a)


if __name__ == "__main__":
    main()
