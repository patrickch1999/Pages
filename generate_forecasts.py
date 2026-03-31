"""
WYSIWYG Editor — Daily Forecast Generator
Fetches Jira data for MWIP-884 and MWIP-903, computes the schedule,
and writes wysiwyg_forecast.html to disk.
"""

import json, math, os, sys
from datetime import date, timedelta
from base64 import b64encode

try:
    import requests
except ImportError:
    sys.exit("Missing 'requests'. Run: pip install requests")
# ── Jira credentials (from environment / GitHub Secrets) ──────────────────────
JIRA_BASE  = os.environ.get("JIRA_BASE_URL", "https://madewithintent.atlassian.net")
JIRA_EMAIL = os.environ.get("JIRA_EMAIL", "")
JIRA_TOKEN = os.environ.get("JIRA_API_TOKEN", "")

if not JIRA_EMAIL or not JIRA_TOKEN:
    sys.exit("JIRA_EMAIL and JIRA_API_TOKEN environment variables must be set.")

AUTH = b64encode(f"{JIRA_EMAIL}:{JIRA_TOKEN}".encode()).decode()
HEADERS = {"Authorization": f"Basic {AUTH}", "Accept": "application/json"}

# ── Constants ──────────────────────────────────────────────────────────────────
TODAY          = date.today()
PPD            = 18        # pixels per working day on Gantt
LABEL_W        = 260       # px for row label column
GAP            = 8         # px gap between label and track
BANK_HOLIDAYS  = {date(2026,4,3),date(2026,4,6),date(2026,5,4),date(2026,5,25),date(2026,8,31)}
HARRY_LEAVE_S  = date(2026,4,13)
HARRY_LEAVE_E  = date(2026,4,17)
EXCL_STATUSES  = {"In Review","Done","QA / UAT","Won't Do","WONT DO"}
EXCLUDED_KEYS  = {"MWIP-971"}   # excluded from estimation by request

# ── Date helpers ───────────────────────────────────────────────────────────────
def is_bh(d):        return d in BANK_HOLIDAYS
def is_hl(d):        return HARRY_LEAVE_S <= d <= HARRY_LEAVE_E
def is_wd_gen(d):    return d.weekday() < 5                          # Mon-Fri only
def is_wd(d, a=None):
    if d.weekday() >= 5: return False
    if is_bh(d):         return False
    if a == "Harry Pegrum" and is_hl(d): return False
    return True

def wd_offset(d):
    """Count Mon-Fri days from TODAY (inclusive=0) to d. BH count — only weekends excluded."""
    count = 0; cur = TODAY
    while cur < d:
        if is_wd_gen(cur): count += 1
        cur += timedelta(1)
    return count

def px(d):          return wd_offset(d) * PPD
def next_wd(d, a=None):
    d += timedelta(1)
    while not is_wd(d, a): d += timedelta(1)
    return d

def end_date(start, n, a=None):
    n = max(1, math.ceil(n)); counted = 0; d = start
    while True:
        if is_wd(d, a): counted += 1
        if counted == n: return d
        d += timedelta(1)

def wds_remaining(a, b):
    count = 0; d = a
    while d <= b:
        if is_wd(d): count += 1
        d += timedelta(1)
    return count

def fmt(d):   return d.strftime("%-d %b %Y")
def fmt_s(d): return d.strftime("%-d %b")
def iso(d):   return d.isoformat()

# ── Bank holiday guard ─────────────────────────────────────────────────────────
if TODAY in BANK_HOLIDAYS:
    print(f"Today ({fmt(TODAY)}) is a UK bank holiday — skipping run.")
    sys.exit(0)

# ── Fetch Jira ─────────────────────────────────────────────────────────────────
def jira_search(jql):
    def jira_search(jql):
    url     = f"{JIRA_BASE}/rest/api/3/search/jql"
    payload = {
        "jql": jql,
        "fields": ["summary", "assignee", "status", "timeoriginalestimate"],
        "maxResults": 100,
    }
    r = requests.post(url, headers={**HEADERS, "Content-Type": "application/json"},
                      json=payload, timeout=30)
    r.raise_for_status()
    return r.json()["issues"]
print("Fetching Jira data…")
raw884 = jira_search("parent = MWIP-884 ORDER BY created ASC")
raw903 = jira_search("parent = MWIP-903 ORDER BY created ASC")
print(f"  884: {len(raw884)} tasks, 903: {len(raw903)} tasks")

def parse_task(issue, epic):
    f = issue["fields"]
    return {
        "key":           issue["key"],
        "summary":       f.get("summary", ""),
        "assignee":      f["assignee"]["displayName"] if f.get("assignee") else None,
        "status":        f["status"]["name"],
        "estimate_secs": f.get("timeoriginalestimate"),
        "epic":          epic,
        "blockedUntil903": issue["key"] == "MWIP-968",
        "finalStage":      issue["key"] == "MWIP-972",
        "excluded_by_request": issue["key"] in EXCLUDED_KEYS,
    }

raw_tasks = (
    [parse_task(i, "884") for i in raw884 if i["key"] != "MWIP-971"] +
    [parse_task(i, "903") for i in raw903]
)
TASKS = {t["key"]: t for t in raw_tasks}

# Add MWIP-971 for change-tracking only (not scheduling)
for i in raw884:
    if i["key"] == "MWIP-971":
        f = i["fields"]
        TASKS["MWIP-971"] = {
            "key": "MWIP-971",
            "summary": f.get("summary",""),
            "assignee": f["assignee"]["displayName"] if f.get("assignee") else None,
            "status": f["status"]["name"],
            "estimate_secs": f.get("timeoriginalestimate"),
            "epic": "884",
            "excluded_by_request": True,
            "blockedUntil903": False, "finalStage": False,
        }

# ── Change detection ───────────────────────────────────────────────────────────
STATE_FILE = "forecast_state.json"
try:
    with open(STATE_FILE) as f: prev = json.load(f)
    prev_tasks = prev.get("tasks", {})
except Exception:
    prev_tasks = None

changes = []
if prev_tasks is None:
    changes = ["First run — baseline established."]
else:
    curr_keys = {k for k in TASKS if not TASKS[k].get("excluded_by_request")}
    prev_keys = {k for k in prev_tasks if k != "MWIP-971"}
    for k in curr_keys - prev_keys:
        changes.append(f"{k} added: {TASKS[k]['summary']}")
    for k in prev_keys - curr_keys:
        changes.append(f"{k} removed")
    for k in curr_keys & prev_keys:
        c, p = TASKS[k], prev_tasks[k]
        if c["status"] != p.get("status"):
            changes.append(f"{k} status: {p.get('status')} → {c['status']}")
        if c["estimate_secs"] != p.get("estimate_secs"):
            cd = round(c["estimate_secs"]/28800,1) if c["estimate_secs"] else "none"
            pd = round(p["estimate_secs"]/28800,1) if p.get("estimate_secs") else "none"
            changes.append(f"{k} estimate: {pd}d → {cd}d")
        ca = c["assignee"] or "Unassigned"; pa = p.get("assignee") or "Unassigned"
        if ca != pa: changes.append(f"{k} assignee: {pa} → {ca}")

# Save state
state_out = {"run_date": iso(TODAY), "tasks": {}}
for k, v in TASKS.items():
    state_out["tasks"][k] = {
        "summary": v["summary"], "assignee": v["assignee"] or "Unassigned",
        "status": v["status"], "estimate_secs": v["estimate_secs"],
    }
with open(STATE_FILE, "w") as f:
    json.dump(state_out, f, indent=2)
print(f"  Changes detected: {len(changes)}")

# ── Schedule computation ───────────────────────────────────────────────────────
def is_excl(v): return v.get("excluded_by_request") or v["status"] in EXCL_STATUSES

def compute_schedule(tol=False):
    def edays(v):
        if not v["estimate_secs"]: return 1
        d = v["estimate_secs"] / 28800.0
        if tol and v["status"] != "In Progress" and v["status"] not in EXCL_STATUSES:
            d *= 1.2
        return d

    sched = {}

    # 903 — Damon sequential
    t903_ip = [(k,v) for k,v in sorted(TASKS.items()) if v["epic"]=="903" and not is_excl(v) and v["status"]=="In Progress"]
    t903_td = [(k,v) for k,v in sorted(TASKS.items()) if v["epic"]=="903" and not is_excl(v) and v["status"]!="In Progress"]
    cur903 = TODAY
    for k, v in t903_ip:
        e = end_date(TODAY, edays(v), v["assignee"])
        sched[k] = {"start": TODAY, "end": e, "days": edays(v)}
        cur903 = next_wd(e, v["assignee"])
    for k, v in t903_td:
        s = cur903
        while not is_wd(s, v["assignee"]): s += timedelta(1)
        e = end_date(s, edays(v), v["assignee"])
        sched[k] = {"start": s, "end": e, "days": edays(v)}
        cur903 = next_wd(e, v["assignee"])

    t903_comp = max((sched[k]["end"] for k in sched if TASKS[k]["epic"]=="903"), default=TODAY)

    # 884 — pre-pass: lock assignee cursors to end of in-progress tasks
    acur = {}
    def proc(keys, blocked=False):
        for k in keys:
            v = TASKS[k]
            if is_excl(v) or v.get("excluded_by_request"): continue
            a = v["assignee"] or "Unassigned"
            d = edays(v)
            if v["status"] == "In Progress":
                s = TODAY
            else:
                s = acur.get(a, TODAY)
                while not is_wd(s, a): s += timedelta(1)
                if blocked:
                    unblock = next_wd(t903_comp, a)
                    if unblock > s: s = unblock
            e = end_date(s, d, a)
            sched[k] = {"start": s, "end": e, "days": d}
            nxt = next_wd(e, a)
            if a not in acur or nxt > acur[a]: acur[a] = nxt

    ip_keys = [k for k,v in sorted(TASKS.items()) if v["epic"]=="884" and not is_excl(v) and not v.get("excluded_by_request") and not v.get("finalStage") and not v.get("blockedUntil903") and v["status"]=="In Progress"]
    td_keys = [k for k,v in sorted(TASKS.items()) if v["epic"]=="884" and not is_excl(v) and not v.get("excluded_by_request") and not v.get("finalStage") and not v.get("blockedUntil903") and v["status"]!="In Progress"]
    b9_keys = [k for k,v in sorted(TASKS.items()) if v["epic"]=="884" and not is_excl(v) and not v.get("excluded_by_request") and not v.get("finalStage") and v.get("blockedUntil903")]

    proc(ip_keys); proc(td_keys); proc(b9_keys, blocked=True)

    # Final stage (MWIP-972)
    all_ends = [sched[k]["end"] for k in sched]
    latest = max(all_ends) if all_ends else TODAY
    fs = next_wd(latest)
    for k, v in TASKS.items():
        if v.get("finalStage") and not is_excl(v):
            e = end_date(fs, edays(v), None)
            sched[k] = {"start": fs, "end": e, "days": edays(v)}

    return sched, t903_comp

BASE_SCHED, T903_BASE = compute_schedule(False)
TOL_SCHED,  T903_TOL  = compute_schedule(True)
BASE_884_END = max(BASE_SCHED[k]["end"] for k in BASE_SCHED if TASKS[k]["epic"]=="884")
TOL_884_END  = max(TOL_SCHED[k]["end"]  for k in TOL_SCHED  if TASKS[k]["epic"]=="884")
BASE_CP      = wds_remaining(TODAY, BASE_884_END)
T903_REM     = wds_remaining(TODAY, T903_BASE)
T903_COUNT   = len([k for k,v in TASKS.items() if v["epic"]=="903" and not is_excl(v)])
BASE_968_S   = BASE_SCHED.get("MWIP-968",{}).get("start", T903_BASE)
TOL_968_S    = TOL_SCHED.get("MWIP-968",{}).get("start",  T903_TOL)
RUN_DT       = TODAY.strftime("%-d %b %Y")

print(f"  884 baseline: {fmt(BASE_884_END)} | 903 baseline: {fmt(T903_BASE)}")

# ── Gantt helpers ──────────────────────────────────────────────────────────────
GANTT_END = TOL_884_END + timedelta(5)
TW_DAYS   = wd_offset(GANTT_END) + 1
TW        = max(TW_DAYS * PPD, 660)

tick_wds = []
cur = TODAY; idx = 0
while cur <= GANTT_END:
    if is_wd_gen(cur):
        if idx % 2 == 0: tick_wds.append(cur)
        idx += 1
    cur += timedelta(1)

bh_in_range = [d for d in sorted(BANK_HOLIDAYS) if TODAY <= d <= GANTT_END and is_wd_gen(d)]
hl_days = [HARRY_LEAVE_S + timedelta(i) for i in range(5) if is_wd_gen(HARRY_LEAVE_S + timedelta(i))]

def bar_px(start, end):
    s = px(start)
    w = (wd_offset(end) - wd_offset(start) + 1) * PPD
    return s, max(w, PPD)

ASSIGNEE_ORDER_884 = ["Geoffrey Viljoen","Damian Dawber","Harry Pegrum","Unassigned"]

def bar_colour(k, v):
    if v["epic"] == "903": return "#1D9E75"
    if v["status"] == "In Progress": return "#EF9F27"
    if v.get("blockedUntil903"): return "#E24B4A"
    if not v["assignee"]: return "#9b59b6"
    return "#378ADD"

def gantt_rows(sched):
    rows884 = []
    for a in ASSIGNEE_ORDER_884:
        tasks_a = [(k,v) for k,v in TASKS.items()
                   if v["epic"]=="884" and (v["assignee"] or "Unassigned")==a
                   and k in sched and not is_excl(v) and not v.get("excluded_by_request")]
        tasks_a.sort(key=lambda x: sched[x[0]]["start"])
        rows884 += tasks_a
    rows903 = sorted(
        [(k,v) for k,v in TASKS.items() if v["epic"]=="903" and k in sched and not is_excl(v)],
        key=lambda x: sched[x[0]]["start"]
    )
    return rows884, rows903

def build_panel(sched, panel_id):
    rows884, rows903 = gantt_rows(sched)
    all_rows = rows884 + [None] + rows903
    ROW_H = 26; SEP_H = 10

    h = f'<div id="{panel_id}" class="gantt-panel">\n'
    # X-axis
    h += f'<div style="display:flex;align-items:flex-end;gap:{GAP}px;margin-bottom:4px">'
    h += f'<div style="width:{LABEL_W}px;flex-shrink:0"></div>'
    h += f'<div style="position:relative;width:{TW}px;height:18px">'
    for td in tick_wds:
        h += f'<span style="position:absolute;left:{px(td)}px;font-size:10px;color:#888780;transform:translateX(-50%);white-space:nowrap">{fmt_s(td)}</span>'
    h += '</div></div>\n'

    # Wrapper with BH stripes
    h += '<div style="position:relative">\n'
    for bh in bh_in_range:
        left = LABEL_W + GAP + px(bh)
        h += f'<div style="position:absolute;top:0;bottom:0;left:{left}px;width:{PPD}px;background:repeating-linear-gradient(45deg,#c9c8c1 0px,#c9c8c1 3px,#eae9e3 3px,#eae9e3 7px);opacity:.7;pointer-events:none;z-index:1"></div>\n'

    for row in all_rows:
        if row is None:
            h += f'<div style="height:{SEP_H}px"></div>\n'; continue
        k, v = row; s_info = sched[k]
        sp, w = bar_px(s_info["start"], s_info["end"])
        colour = bar_colour(k, v)
        border = "border:1.5px dashed #a32d2d;" if v.get("blockedUntil903") else ""
        a_name = v["assignee"] or "Unassigned"
        tt = f'{k}: {fmt_s(s_info["start"])} → {fmt_s(s_info["end"])} ({round(s_info["days"],1)}d)'
        lbl_title = v["summary"] + " (" + a_name + ")"
        summ = v["summary"][:48]

        h += f'<div style="display:flex;align-items:center;gap:{GAP}px;height:{ROW_H}px;margin-bottom:3px">'
        h += f'<div class="row-label" title="{lbl_title}">'
        h += f'<a href="https://madewithintent.atlassian.net/browse/{k}" target="_blank">{k}</a> {summ}'
        h += f'<br><span style="font-size:10px;color:#888780">{a_name}</span></div>'
        h += f'<div style="position:relative;width:{TW}px;height:16px;flex-shrink:0">'
        for td in tick_wds:
            h += f'<div style="position:absolute;top:0;bottom:0;left:{px(td)}px;width:.5px;background:#e8e7e0;pointer-events:none"></div>'
        if a_name == "Harry Pegrum" and hl_days:
            hl_s = px(hl_days[0]); hl_e = px(hl_days[-1]) + PPD
            h += f'<div class="leave-block" style="left:{hl_s}px;width:{hl_e-hl_s}px" title="Harry leave 13–17 Apr"></div>'
        h += f'<div class="bar" style="left:{sp}px;width:{w}px;background:{colour};{border}" title="{tt}"></div>'
        h += '</div></div>\n'

    h += '</div></div>\n'
    return h

# ── Tables ─────────────────────────────────────────────────────────────────────
def badge(status):
    m = {"In Progress": ("b-progress","In Progress"), "In Review": ("b-review","In Review"),
         "QA / UAT": ("b-review","QA / UAT"), "Backlog": ("b-todo","To Do"),
         "Ready for Development": ("b-todo","To Do"), "WONT DO": ("b-excl","Won't Do"),
         "Won't Do": ("b-excl","Won't Do")}
    cls, label = m.get(status, ("b-todo", status))
    return f'<span class="badge {cls}">{label}</span>'

def link(k): return f'<a href="https://madewithintent.atlassian.net/browse/{k}" target="_blank">{k}</a>'

def table_884(sched):
    order = ["MWIP-886","MWIP-887","MWIP-888","MWIP-889","MWIP-890","MWIP-891",
             "MWIP-893","MWIP-967","MWIP-968","MWIP-969","MWIP-971","MWIP-972"]
    rows = ""
    for k in order:
        if k not in TASKS: continue
        v = TASKS[k]; a = v["assignee"] or "Unassigned"
        est = f'{round(v["estimate_secs"]/28800,1)}d' if v["estimate_secs"] else "—"
        if v.get("excluded_by_request"):
            rows += f'<tr><td>{link(k)}</td><td class="assignee">{a}</td><td class="est">{est}</td><td colspan="2" class="excl">excl. from estimation by request</td><td><span class="badge b-excl">excl.</span></td></tr>'
        elif is_excl(v):
            rows += f'<tr><td>{link(k)}</td><td class="assignee">{a}</td><td class="est">{est}</td><td colspan="2" class="excl">excl. from schedule</td><td>{badge(v["status"])}</td></tr>'
        elif k in sched:
            s = sched[k]
            rows += f'<tr><td>{link(k)}</td><td class="assignee">{a}</td><td class="est">{est}</td><td>{fmt_s(s["start"])}</td><td>{fmt_s(s["end"])}</td><td>{badge(v["status"])}</td></tr>'
    return rows

def table_903(sched):
    order = ["MWIP-952","MWIP-953","MWIP-954","MWIP-955","MWIP-956",
             "MWIP-957","MWIP-958","MWIP-959","MWIP-960","MWIP-961","MWIP-962"]
    rows = ""
    for k in order:
        if k not in TASKS: continue
        v = TASKS[k]
        est = f'{round(v["estimate_secs"]/28800,1)}d' if v["estimate_secs"] else "—"
        if is_excl(v):
            rows += f'<tr><td>{link(k)}</td><td class="est">{est}</td><td colspan="2" class="excl">excl. from schedule</td><td>{badge(v["status"])}</td></tr>'
        elif k in sched:
            s = sched[k]
            rows += f'<tr><td>{link(k)}</td><td class="est">{est}</td><td>{fmt_s(s["start"])}</td><td>{fmt_s(s["end"])}</td><td>{badge(v["status"])}</td></tr>'
    return rows

# ── Changes banner ─────────────────────────────────────────────────────────────
if not changes:
    changes_html = '<div class="no-changes">✅ No changes detected since last run — schedule unchanged.</div>'
else:
    items = "".join(f"<li>{c}</li>" for c in changes)
    changes_html = f'<div class="changes-box"><div class="changes-title">⚠️ {len(changes)} change{"s" if len(changes)!=1 else ""} since last run</div><ul class="changes-list">{items}</ul></div>'

# ── Build HTML ─────────────────────────────────────────────────────────────────
base_panel = build_panel(BASE_SCHED, "panel-base")
tol_panel  = build_panel(TOL_SCHED,  "panel-tol")

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>WYSIWYG Editor — Epic Forecast · {RUN_DT}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;color:#1a1a18;background:#f5f4f0;padding:24px}}
h1{{font-size:18px;font-weight:600;margin-bottom:4px}}
.subtitle{{font-size:12px;color:#73726c;margin-bottom:20px}}
.no-changes{{background:#e1f5ee;border:.5px solid #0f6e56;border-radius:10px;padding:10px 14px;margin-bottom:16px;font-size:12px;color:#0a4a39;font-weight:500}}
.changes-box{{background:#faeeda;border:.5px solid #ba7517;border-radius:10px;padding:10px 14px;margin-bottom:16px}}
.changes-title{{font-size:13px;font-weight:600;color:#633806;margin-bottom:6px}}
.changes-list{{list-style:none;font-size:12px;color:#854f0b;line-height:1.7}}
.changes-list li::before{{content:"↳ ";font-weight:600}}
.metrics{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:16px}}
.metric{{background:#fff;border:.5px solid #d3d1c7;border-radius:10px;padding:12px 14px}}
.metric-label{{font-size:11px;color:#73726c;margin-bottom:3px}}
.metric-val{{font-size:22px;font-weight:600;line-height:1.2}}
.metric-sub{{font-size:11px;color:#888780;margin-top:2px}}
.info-box{{border-radius:10px;padding:10px 14px;margin-bottom:14px}}
.info-box.amber{{background:#faeeda;border:.5px solid #ba7517}}
.info-box.blue{{background:#e6f1fb;border:.5px solid #185fa5}}
.info-title{{font-size:13px;font-weight:600;margin-bottom:3px}}
.amber .info-title{{color:#633806}}.amber .info-body{{font-size:12px;color:#854f0b;line-height:1.5}}
.blue .info-title{{color:#0c447c}}.blue .info-body{{font-size:12px;color:#185fa5;line-height:1.5}}
.section-label{{font-size:11px;font-weight:600;color:#73726c;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px}}
.tabs{{display:flex;border-bottom:.5px solid #d3d1c7}}
.tab{{font-size:12px;font-weight:500;padding:8px 16px;cursor:pointer;color:#73726c;border:.5px solid transparent;border-bottom:none;border-radius:8px 8px 0 0;background:transparent;user-select:none}}
.tab.active{{background:#fff;border-color:#d3d1c7;color:#1a1a18;margin-bottom:-.5px}}
.tab:hover:not(.active){{background:#f1efe8}}
.chart-wrap{{background:#fff;border:.5px solid #d3d1c7;border-top:none;border-radius:0 0 10px 10px;padding:16px;margin-bottom:16px;overflow-x:auto}}
.gantt-panel{{display:none}}.gantt-panel.active{{display:block}}
.row-label{{font-size:11px;color:#3d3d3a;width:{LABEL_W}px;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.4}}
.row-label a{{color:#185fa5;text-decoration:none}}
.bar{{position:absolute;height:100%;border-radius:3px;cursor:default}}
.leave-block{{position:absolute;height:100%;background:repeating-linear-gradient(45deg,#f09595 0,#f09595 3px,#fcebeb 3px,#fcebeb 7px);border-radius:2px;opacity:.7;pointer-events:none;z-index:2}}
.tables{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.table-card{{background:#fff;border:.5px solid #d3d1c7;border-radius:10px;padding:14px 16px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{font-size:11px;font-weight:600;color:#73726c;text-align:left;padding:4px 6px 6px;border-bottom:.5px solid #d3d1c7}}
td{{padding:5px 6px;border-bottom:.5px solid #f1efe8;color:#3d3d3a;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
.badge{{display:inline-block;font-size:10px;font-weight:500;padding:2px 6px;border-radius:4px;white-space:nowrap}}
.b-progress{{background:#faeeda;color:#854f0b}}.b-review{{background:#e1f5ee;color:#0f6e56}}
.b-todo{{background:#f1efe8;color:#5f5e5a}}.b-excl{{background:#f1efe8;color:#888780;font-style:italic}}
.est{{color:#73726c;text-align:right}}.assignee{{font-size:11px;color:#888780}}.excl{{color:#888780;font-style:italic}}
.assumptions{{background:#fff;border:.5px solid #d3d1c7;border-radius:10px;padding:14px 16px;margin-top:14px}}
.assumptions ul{{list-style:none;display:grid;grid-template-columns:1fr 1fr;gap:8px 24px;font-size:12px;color:#3d3d3a;line-height:1.6}}
.assumptions li::before{{content:"⚠ ";color:#854f0b;font-weight:600}}
.oos-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px 24px;font-size:12px;color:#3d3d3a;line-height:1.6}}
.oos-item{{display:flex;gap:8px;align-items:flex-start}}.oos-x{{color:#a32d2d;font-weight:600;flex-shrink:0}}
.legend{{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;color:#73726c;margin-top:12px;align-items:center}}
.lsq{{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}}
.footer{{margin-top:20px;font-size:11px;color:#888780;border-top:.5px solid #d3d1c7;padding-top:12px}}
</style>
</head>
<body>
{changes_html}
<h1>WYSIWYG Editor — Epic Forecast</h1>
<p class="subtitle">Run date: {RUN_DT} · Epics: MWIP-884 (Editor Runtime) &amp; MWIP-903 (Widget Library) · Scheduling origin: {fmt_s(TODAY)}</p>
<div class="metrics">
  <div class="metric"><div class="metric-label">MWIP-884 Forecast</div><div class="metric-val">{fmt_s(BASE_884_END)}</div><div class="metric-sub">Baseline completion</div></div>
  <div class="metric"><div class="metric-label">884 Critical Path</div><div class="metric-val">{BASE_CP}d</div><div class="metric-sub">Working days remaining</div></div>
  <div class="metric"><div class="metric-label">MWIP-903 Forecast</div><div class="metric-val">{fmt_s(T903_BASE)}</div><div class="metric-sub">Baseline completion</div></div>
  <div class="metric"><div class="metric-label">903 Remaining</div><div class="metric-val">{T903_REM}d</div><div class="metric-sub">{T903_COUNT} active tasks</div></div>
</div>
<div class="info-box amber">
  <div class="info-title">⚠️ MWIP-968 dependency on MWIP-903</div>
  <div class="info-body"><strong>Integrate widgets into editor shell</strong> cannot start until all MWIP-903 tasks complete.<br>
  Baseline: 903 completes <strong>{fmt_s(T903_BASE)}</strong> → 968 from <strong>{fmt_s(BASE_968_S)}</strong> &nbsp;·&nbsp;
  +20%: 903 completes <strong>{fmt_s(T903_TOL)}</strong> → 968 from <strong>{fmt_s(TOL_968_S)}</strong></div>
</div>
<div class="info-box blue">
  <div class="info-title">ℹ️ Tolerance variant (+20%)</div>
  <div class="info-body">+20% multiplies all <em>To Do</em> estimates. <em>In Progress</em> unchanged.
  Baseline 884: <strong>{fmt_s(BASE_884_END)}</strong> · Tolerance 884: <strong>{fmt_s(TOL_884_END)}</strong></div>
</div>
<p class="section-label">Gantt Chart</p>
<div class="tabs">
  <div class="tab active" onclick="switchTab('base')">Baseline</div>
  <div class="tab" onclick="switchTab('tol')">With +20% Tolerance</div>
</div>
<div class="chart-wrap">
{base_panel}
{tol_panel}
<div class="legend">
  <span><span class="lsq" style="background:#378ADD"></span>MWIP-884 (To Do)</span>
  <span><span class="lsq" style="background:#EF9F27"></span>In Progress</span>
  <span><span class="lsq" style="background:#1D9E75"></span>MWIP-903</span>
  <span><span class="lsq" style="background:#E24B4A;border:1.5px dashed #a32d2d"></span>Blocked (awaiting 903)</span>
  <span><span class="lsq" style="background:#9b59b6"></span>Unassigned</span>
  <span><span class="lsq" style="background:repeating-linear-gradient(45deg,#f09595 0,#f09595 3px,#fcebeb 3px,#fcebeb 7px)"></span>Harry leave 13–17 Apr</span>
  <span><span class="lsq" style="background:repeating-linear-gradient(45deg,#c9c8c1 0,#c9c8c1 3px,#eae9e3 3px,#eae9e3 7px)"></span>UK Bank Holiday</span>
</div>
</div>
<p class="section-label">Task Detail</p>
<div class="tables">
  <div class="table-card">
    <div class="section-label" style="margin-bottom:10px">MWIP-884 · Editor Runtime</div>
    <table><thead><tr><th>Key</th><th>Assignee</th><th class="est">Est.</th><th>Start</th><th>End</th><th>Status</th></tr></thead>
    <tbody>{table_884(BASE_SCHED)}</tbody></table>
  </div>
  <div class="table-card">
    <div class="section-label" style="margin-bottom:10px">MWIP-903 · Widget Library</div>
    <table><thead><tr><th>Key</th><th class="est">Est.</th><th>Start</th><th>End</th><th>Status</th></tr></thead>
    <tbody>{table_903(BASE_SCHED)}</tbody></table>
  </div>
</div>
<div class="assumptions" style="margin-top:14px">
  <div class="section-label" style="margin-bottom:10px">Out-of-Scope Exclusions</div>
  <div class="oos-grid">
    <div class="oos-item"><span class="oos-x">✕</span><span>MWIP-971: Editing live campaigns in draft mode — excluded by request</span></div>
    <div class="oos-item"><span class="oos-x">✕</span><span>Performance optimisation &amp; load testing</span></div>
    <div class="oos-item"><span class="oos-x">✕</span><span>Third-party integrations not listed in epics</span></div>
    <div class="oos-item"><span class="oos-x">✕</span><span>Post-launch bug fixes and ongoing maintenance</span></div>
    <div class="oos-item"><span class="oos-x">✕</span><span>Documentation and internal training materials</span></div>
    <div class="oos-item"><span class="oos-x">✕</span><span>Platform/infra provisioning and CI/CD pipeline changes</span></div>
  </div>
</div>
<div class="assumptions">
  <div class="section-label" style="margin-bottom:10px">Assumptions &amp; Limitations</div>
  <ul>
    <li>Estimates are original Jira estimates — no burn-down on in-progress tasks</li>
    <li>In Progress tasks are scheduled to start today regardless of queue position</li>
    <li>Sequential queuing per assignee: each task starts after the previous ends</li>
    <li>MWIP-968 cannot start until ALL MWIP-903 tasks are complete</li>
    <li>MWIP-972 (Testing &amp; migration) starts only after every other task completes</li>
    <li>Harry Pegrum leave 13–17 Apr 2026 excluded from his working days</li>
    <li>UK bank holidays excluded: 3 Apr, 6 Apr, 4 May, 25 May, 31 Aug 2026</li>
    <li>Unassigned tasks (excl. 972) scheduled from today with no queue dependency</li>
  </ul>
</div>
<div class="footer">
  Generated: {TODAY.strftime("%A %-d %B %Y")} · Bank holidays: 3 Apr, 6 Apr, 4 May, 25 May, 31 Aug 2026 ·
  Harry leave: 13–17 Apr 2026 · 1 working day = 8h
</div>
<script>
function switchTab(id){{
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.gantt-panel').forEach(p=>p.classList.remove('active'));
  if(id==='base'){{document.querySelectorAll('.tab')[0].classList.add('active');document.getElementById('panel-base').classList.add('active');}}
  else{{document.querySelectorAll('.tab')[1].classList.add('active');document.getElementById('panel-tol').classList.add('active');}}
}}
document.getElementById('panel-base').classList.add('active');
</script>
</body></html>"""

out = "wysiwyg_forecast.html"
with open(out, "w") as f:
    f.write(HTML)
print(f"Written: {out} ({len(HTML):,} chars)")
