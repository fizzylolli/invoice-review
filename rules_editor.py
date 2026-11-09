#!/usr/bin/env python3
"""
Invoice Review app window.

Start it by double-clicking "Start Invoice Review.bat" (or run: python rules_editor.py).
A page opens in your browser. It runs only on your own computer, nothing is uploaded.
  1. Choose the folder of invoices and click Run.
  2. See what needs a person, and open the spreadsheet.
  3. Under "Rules", switch rules on and off or add your own, and Save.

Uses only what comes with Python.
"""
import csv, json, os, subprocess, sys, threading, webbrowser
from datetime import datetime
from urllib.parse import parse_qs, urlparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import invoice_pipeline as ip

HERE = os.path.dirname(os.path.abspath(__file__))
RULES_PATH = os.path.join(HERE, "rules.json")
PIPELINE = os.path.join(HERE, "invoice_pipeline.py")
OUT_DIR = os.path.join(HERE, "out")


LAST_FOLDER = {}   # the invoice folder of the latest run
DECISIONS_PATH = os.path.join(OUT_DIR, "decisions.json")


def load_decisions():
    try:
        with open(DECISIONS_PATH, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def save_decisions(d):
    """Approvals live next to the results. approvals.csv is the same list for anyone who wants to open it in Excel."""
    os.makedirs(OUT_DIR, exist_ok=True)
    tmp = DECISIONS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=1)
    os.replace(tmp, DECISIONS_PATH)
    try:
        with open(os.path.join(OUT_DIR, "approvals.csv"), "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["file", "vendor", "invoice_number", "amount", "decision", "when"])
            for v in d.values():
                w.writerow([v.get("file"), v.get("vendor"), v.get("invoice_number"), v.get("amount"), "approved", v.get("when")])
    except OSError:
        pass                                              # the CSV copy is optional (it may be open in Excel)


def clean_rules(data):
    """Accept only well-formed settings from the browser; anything odd falls back to defaults."""
    rules = ip.load_rules(None)

    def num(v, default):
        try:
            return float(v)
        except (TypeError, ValueError):
            return default

    def strs(v):
        return [str(x).strip() for x in v if str(x).strip()] if isinstance(v, list) else []

    for key in rules:
        if key == "custom":
            continue
        src = data.get(key, {}) if isinstance(data.get(key), dict) else {}
        rules[key]["enabled"] = bool(src.get("enabled", rules[key]["enabled"]))
    t = data.get("threshold", {}); rules["threshold"]["amount"] = max(0.0, num(t.get("amount"), 5000))
    m = data.get("line_item_math", {}); rules["line_item_math"]["tolerance"] = max(0.0, num(m.get("tolerance"), 0.01))
    f = data.get("foreign_currency", {}); rules["foreign_currency"]["eur_to_usd"] = max(0.01, num(f.get("eur_to_usd"), 1.10))
    rules["portfolio_company"]["names"] = strs(data.get("portfolio_company", {}).get("names"))
    o = data.get("overdue", {}); rules["overdue"]["days"] = max(0, int(num(o.get("days"), 30)))
    rules["new_vendor"]["known_vendors"] = strs(data.get("new_vendor", {}).get("known_vendors"))
    custom = []
    for c in (data.get("custom") if isinstance(data.get("custom"), list) else [])[:50]:
        if not isinstance(c, dict):
            continue
        custom.append({
            "name": str(c.get("name", "")).strip()[:80] or "Untitled rule",
            "enabled": bool(c.get("enabled", True)),
            "contains": strs(c.get("contains")),
            "reason": str(c.get("reason", "")).strip()[:300],
            "next_step": str(c.get("next_step", "")).strip()[:300],
            "priority": 1 if str(c.get("priority")) == "1" else 2,
        })
    rules["custom"] = custom
    return rules


INVOICE_EXT = (".pdf", ".jpg", ".jpeg", ".png", ".eml")


def python_exe():
    """python.exe, even when this app was started with pythonw.exe (which has no console)."""
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe"):
        cand = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(cand):
            return cand
    return exe


def list_dir(path):
    """Folders inside `path` for the in-page folder browser. Empty path = the drives (or / on Mac/Linux)."""
    path = (path or "").strip().strip('"')
    if not path:
        if os.name == "nt":
            drives = os.listdrives() if hasattr(os, "listdrives") else [f"{c}:\\" for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{c}:\\")]
            return {"path": "", "parent": "", "invoices": 0, "dirs": [{"name": d, "path": d} for d in drives]}
        path = os.sep
    path = os.path.abspath(path)
    while path and not os.path.isdir(path):
        nxt = os.path.dirname(path)
        if nxt == path:
            return list_dir("")
        path = nxt
    dirs, files = [], 0
    try:
        for name in sorted(os.listdir(path), key=str.lower):
            full = os.path.join(path, name)
            if name.startswith((".", "$")):
                continue
            if os.path.isdir(full):
                dirs.append({"name": name, "path": full})
            elif name.lower().endswith(INVOICE_EXT):
                files += 1
    except OSError:
        pass
    parent = os.path.dirname(path)
    return {"path": path, "parent": "" if parent == path else parent, "invoices": files, "dirs": dirs}


PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Invoice Review</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1f2937;--mute:#6b7280;--line:#d9dde3;--accent:#1d4ed8;--accent-ink:#fff;--ok:#166534;--bad:#9b1c1c;--warn:#92400e;--off:#c7ccd4;--chip:#eef2ff}
@media (prefers-color-scheme:dark){:root{--bg:#111418;--card:#1a1f26;--ink:#e6e9ee;--mute:#9aa3af;--line:#2d3540;--accent:#6ea0ff;--accent-ink:#0b1220;--ok:#6ee7a0;--bad:#fca5a5;--warn:#fcd34d;--off:#3a4350;--chip:#1e2a44}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:1180px;margin:0 auto;padding:24px 16px 60px}
h1{font-size:24px;margin:0 0 4px} h2{font-size:17px;margin:0 0 8px}
p.sub{margin:0 0 18px;color:var(--mute)}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:12px}
.row{display:flex;gap:12px;align-items:flex-start}
.grow{flex:1;min-width:0}
.title{font-weight:600} .desc{color:var(--mute);font-size:13.5px;margin-top:2px}
.sw{position:relative;flex:none;width:44px;height:26px;margin-top:2px}
.sw input{position:absolute;inset:0;opacity:0;cursor:pointer;margin:0;width:100%;height:100%}
.sw span{position:absolute;inset:0;background:var(--off);border-radius:13px;transition:.15s}
.sw span:after{content:"";position:absolute;top:3px;left:3px;width:20px;height:20px;border-radius:50%;background:#fff;transition:.15s}
.sw input:checked+span{background:var(--accent)} .sw input:checked+span:after{left:21px}
.sw input:focus-visible+span{outline:2px solid var(--accent);outline-offset:2px}
.params{margin-top:10px;display:grid;gap:8px}
.off .params,.off .title,.off .desc{opacity:.55}
label.f{display:block;font-size:13px;color:var(--mute);margin-bottom:3px}
input[type=text],input[type=number],textarea,select{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:7px;background:var(--bg);color:var(--ink);font:inherit}
textarea{min-height:72px;resize:vertical}
input:focus,textarea:focus,select:focus{outline:2px solid var(--accent);outline-offset:1px}
.two{display:grid;grid-template-columns:1fr 1fr;gap:10px} @media(max-width:600px){.two{grid-template-columns:1fr}}
button{font:inherit;border-radius:8px;padding:9px 16px;border:1px solid var(--line);background:var(--card);color:var(--ink);cursor:pointer}
button.primary{background:var(--accent);color:var(--accent-ink);border-color:var(--accent);font-weight:600}
button.big{padding:11px 22px;font-size:16px}
button.link{border:none;background:none;color:var(--bad);padding:4px 0}
button:disabled{opacity:.55;cursor:wait}
button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.inline{display:flex;gap:8px;flex-wrap:wrap} .inline input{flex:1;min-width:240px}
#status{color:var(--mute);font-size:14px;margin-left:10px} #status.ok{color:var(--ok)} #status.err{color:var(--bad)}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 14px}
.chip{background:var(--chip);border-radius:999px;padding:5px 12px;font-size:13.5px}
table{width:100%;border-collapse:collapse;font-size:14px}
th{text-align:left;color:var(--mute);font-weight:600;font-size:12.5px;padding:6px 8px;border-bottom:1px solid var(--line)}
td{padding:9px 8px;border-bottom:1px solid var(--line);vertical-align:top}
.pri1{color:var(--bad);font-weight:600} .pri2{color:var(--warn);font-weight:600}
.num{text-align:right;white-space:nowrap}
th.sortable{cursor:pointer;user-select:none;white-space:nowrap} th.sortable:hover{text-decoration:underline} td.fa{color:var(--mute);font-weight:700}
tr.flagrow td{background:rgba(220,38,38,.10)} tr.c-dup td{background:rgba(147,51,234,.14)} tr.c-math td{background:rgba(245,158,11,.22)} tr.c-cur td{background:rgba(37,99,235,.13)} tr.c-ent td{background:rgba(13,148,136,.15)} tr.c-unr td{background:rgba(107,114,128,.20)} tr.c-lim td{background:rgba(220,38,38,.10)}
.legend{display:flex;gap:8px 14px;flex-wrap:wrap;margin:0 0 12px;font-size:13px} .legend button{cursor:pointer;border:1px solid transparent;border-radius:6px;background:none;color:inherit;font:inherit;padding:2px 6px;display:inline-flex;align-items:center;gap:6px} .legend .sw{width:14px;height:14px;border-radius:3px;display:inline-block;border:1px solid rgba(0,0,0,.25)} .legend button.on{border-color:#d97706;background:rgba(245,158,11,.15)}
a.fl{color:inherit;text-decoration:underline dotted;cursor:pointer}
tr.hl td{font-weight:700;border-top:2px solid #111;border-bottom:2px solid #111} tr.hl td:first-child{box-shadow:inset 5px 0 0 #111} button.chip{cursor:pointer;color:inherit;border:1px solid var(--line,#ccc);font:inherit} tr.apprrow td{background:rgba(22,101,52,.10)} button.sm{padding:4px 10px;font-size:13px;white-space:nowrap} td.act{white-space:nowrap} td.fy{color:var(--bad);font-weight:700} td.fn{color:var(--ok);font-weight:700}
.small{font-size:13px;color:var(--mute)}
details>summary{cursor:pointer;font-weight:600;font-size:17px;padding:2px 0}
.browse{margin-top:10px;border:1px solid var(--line);border-radius:8px;overflow:hidden}
.browse .path{padding:8px 10px;background:var(--bg);font-size:13px;word-break:break-all;border-bottom:1px solid var(--line)}
.browse ul{list-style:none;margin:0;padding:0;max-height:240px;overflow:auto}
.browse li button{display:block;width:100%;text-align:left;border:none;border-radius:0;border-bottom:1px solid var(--line);background:var(--card)}
.browse li button:hover{background:var(--chip)}
.browse .foot{padding:8px 10px;display:flex;gap:10px;align-items:center;justify-content:space-between;flex-wrap:wrap}
pre{background:#0f172a;color:#e2e8f0;padding:12px;border-radius:8px;overflow:auto;max-height:260px;font-size:12.5px;margin:10px 0 0;white-space:pre-wrap}
</style></head><body><main>
<h1>Invoice Review</h1>
<p class="sub">Pick a folder of invoices, click Run, and see what needs a person before anything gets paid. Nothing is paid by this tool.</p>

<div class="card">
  <h2>1. Choose the folder of invoices</h2>
  <div class="inline"><input type="text" id="folder" placeholder="Click Browse, or paste a folder path" aria-label="Folder of invoices">
    <button id="browse">Browse...</button></div>
  <div id="browsebox" class="browse" hidden>
    <div class="path" id="bpath"></div><ul id="blist"></ul>
    <div class="foot"><span class="small" id="bcount"></span><span><button id="bup">Up one level</button> <button id="buse" class="primary">Use this folder</button></span></div>
  </div>
  <div style="margin-top:14px"><button id="run" class="primary big">Run</button><span id="status" role="status"></span></div>
  <pre id="log" hidden></pre>
</div>

<div class="card" id="results" hidden>
  <h2 id="headline"></h2>
  <div class="chips" id="chips"></div>
  <div class="legend" id="legend" aria-label="Colour key"></div>
  <div class="inline" style="margin:4px 0 8px"><label style="display:flex;align-items:center;gap:6px;white-space:nowrap"><input type="checkbox" id="onlyflag" style="width:auto;margin:0"> Only what still needs a person</label>
    <span class="small">Click a column heading to sort. Clearing a flag records that you checked it. It does not pay anything.</span></div>
  <div style="overflow-x:auto"><table><thead id="allhead"></thead><tbody id="allrows"></tbody></table></div>
  <details style="margin-top:12px"><summary><b>Review queue: what to do next</b></summary>
    <table><thead><tr><th>Urgency</th><th>Invoice</th><th class="num">Amount</th><th>Why it was flagged</th><th>Suggested next step</th></tr></thead><tbody id="rows"></tbody></table>
  </details>
  <div class="inline" style="margin-top:14px">
    <button id="openx" class="primary">Open the spreadsheet</button>
    <button id="opend">Open the Monday email</button>
    <button id="openf">Open the results folder</button>
    <button id="openi">Open the invoice folder</button>
  </div>
  <p class="small" style="margin-bottom:0">The spreadsheet has a tab for every invoice, a review queue with a decision dropdown, a summary, and the rules used.</p>
  <div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--line)">
    <div class="title">Send to AP</div>
    <div class="desc" id="apdesc"></div>
    <div class="inline" style="margin-top:8px">
      <input type="text" id="apto" placeholder="AP email address" aria-label="Email address for accounts payable" inputmode="email">
      <button id="sendap" class="primary">Send to AP</button>
      <button id="copyap">Copy the list</button>
    </div>
    <div class="small" id="apnote" style="margin-top:6px"></div>
  </div>
  <div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--line)">
    <div class="title">Email this summary</div>
    <div class="desc">Opens a ready-to-send draft in Gmail. You press Send. Attach the spreadsheet if you want the full list.</div>
    <div class="inline" style="margin-top:8px">
      <input type="text" id="mailto" placeholder="your email address" aria-label="Email address to send to" inputmode="email">
      <button id="emailme" class="primary">Email me this summary</button>
      <button id="copymail">Copy the text</button>
    </div>
    <div class="small" id="mailnote" style="margin-top:6px"></div>
  </div>
</div>

<div class="card">
  <details id="rulesbox"><summary>Rules: what gets flagged</summary>
    <p class="sub" style="margin-top:8px">Switch rules on or off and change their settings. Click <b>Save rules</b> and the next run uses them.</p>
    <div id="rules"></div>
    <h2 style="margin-top:20px">Your own rules</h2>
    <p class="sub" style="margin-top:-2px">Flag an invoice when its text mentions certain words, for example "legal services" or "retainer".</p>
    <div id="custom"></div>
    <button id="add">+ Add a rule</button>
    <div style="margin-top:16px"><button class="primary" id="save">Save rules</button><span id="rstatus" class="small" style="margin-left:10px"></span></div>
  </details>
</div>
<p class="small"><button id="quit">Quit the app</button> Closing this tab does not stop it. Use Quit when you are done.</p>
</main>
<script>
let R = null, INFO = null, bdir = '';
const $ = (s, el=document) => el.querySelector(s);
const h = (tag, attrs={}, ...kids) => { const e = document.createElement(tag);
  for (const [k,v] of Object.entries(attrs)) { if (k==='class') e.className=v; else if (k.startsWith('on')) e[k]=v; else if (v!==undefined && v!==null) e.setAttribute(k,v); }
  kids.flat().forEach(c => e.append(c.nodeType ? c : document.createTextNode(c))); return e; };
const lines = t => t.split(/\n|,/).map(x => x.trim()).filter(Boolean);
const money = (n, cur) => n == null ? '' : (cur === 'EUR' ? '€' : '$') + Number(n).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
let uid = 0;
function toggle(checked, onchange, label) { const i = h('input', {type:'checkbox', 'aria-label':label}); i.checked = checked; i.onchange = () => onchange(i.checked);
  return h('label', {class:'sw'}, i, h('span')); }
function field(label, input) { const id = 'f' + (++uid); input.id = id; return h('div', {}, h('label', {class:'f', for:id}, label), input); }
function numIn(val, set, step='any') { const i = h('input', {type:'number', step, min:'0', value:val}); i.oninput = () => set(i.value === '' ? 0 : Number(i.value)); return i; }
function area(val, set) { const t = h('textarea'); t.value = val; t.oninput = () => set(t.value); return t; }
function textIn(val, set) { const i = h('input', {type:'text', value:val}); i.oninput = () => set(i.value); return i; }
function sel(val, set) { const s = h('select', {}, h('option', {value:'1'}, 'Check first: possible error'), h('option', {value:'2'}, 'Needs approval')); s.value = String(val); s.onchange = () => set(s.value); return s; }

const PARAMS = {
  threshold: c => field('Flag invoices over this amount (US dollars)', numIn(c.amount, v => c.amount = v)),
  line_item_math: c => field('Allowed difference between lines and total (dollars)', numIn(c.tolerance, v => c.tolerance = v, '0.01')),
  foreign_currency: c => field('Euro to dollar rate used for the dollar limit', numIn(c.eur_to_usd, v => c.eur_to_usd = v, '0.01')),
  portfolio_company: c => field('Company names (one per line)', area(c.names.join('\n'), v => c.names = lines(v))),
  overdue: c => field('Flag when this many days or more past due', numIn(c.days, v => c.days = v, '1')),
  new_vendor: c => field('Approved vendors (one per line). Leave empty and this rule does nothing.', area(c.known_vendors.join('\n'), v => c.known_vendors = lines(v))),
};
function render() {
  const box = $('#rules'); box.replaceChildren();
  for (const [key, [title, desc]] of Object.entries(INFO)) {
    const c = R[key]; const card = h('div', {class:'card' + (c.enabled ? '' : ' off')});
    const body = h('div', {class:'grow'}, h('div', {class:'title'}, title), h('div', {class:'desc'}, desc));
    if (PARAMS[key]) body.append(h('div', {class:'params'}, PARAMS[key](c)));
    card.append(h('div', {class:'row'}, toggle(c.enabled, v => { c.enabled = v; card.classList.toggle('off', !v); }, title), body)); box.append(card);
  }
  const cb = $('#custom'); cb.replaceChildren();
  R.custom.forEach((c, i) => {
    const card = h('div', {class:'card' + (c.enabled ? '' : ' off')});
    const body = h('div', {class:'grow'}, h('div', {class:'params'},
      field('Rule name', textIn(c.name, v => c.name = v)),
      field('Flag when the invoice mentions any of these words (one per line)', area(c.contains.join('\n'), v => c.contains = lines(v))),
      field('What to tell the reviewer', textIn(c.reason, v => c.reason = v)),
      h('div', {class:'two'}, field('Suggested next step', textIn(c.next_step, v => c.next_step = v)), field('How urgent', sel(c.priority, v => c.priority = Number(v)))),
      h('div', {}, h('button', {class:'link', onclick: () => { R.custom.splice(i, 1); render(); }}, 'Delete this rule'))));
    card.append(h('div', {class:'row'}, toggle(c.enabled, v => { c.enabled = v; card.classList.toggle('off', !v); }, 'Rule ' + c.name), body)); cb.append(card);
  });
}
async function api(path, body) {
  const r = await fetch(path, body === undefined ? {} : {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)});
  if (!r.ok) throw new Error((await r.text()) || r.status); return r.json();
}
function say(msg, cls) { const s = $('#status'); s.textContent = msg; s.className = cls || ''; }

// folder browser
async function browseTo(path) {
  const d = await api('/api/ls?path=' + encodeURIComponent(path)); bdir = d.path;
  $('#bpath').textContent = d.path || 'This computer'; $('#bcount').textContent = d.path ? d.invoices + ' invoice file(s) in this folder' : '';
  $('#buse').disabled = !d.path; $('#bup').disabled = !d.path;
  const ul = $('#blist'); ul.replaceChildren();
  d.dirs.forEach(n => ul.append(h('li', {}, h('button', {onclick: () => browseTo(n.path)}, '📁 ' + n.name))));
  if (!d.dirs.length) ul.append(h('li', {class:'small', style:'padding:10px'}, 'No folders here.'));
  window._up = d.parent; 
}
$('#browse').onclick = async () => { $('#browsebox').hidden = false; await browseTo($('#folder').value || ''); };
$('#bup').onclick = () => browseTo(window._up || '');
$('#buse').onclick = () => { $('#folder').value = bdir; $('#browsebox').hidden = true; };

$('#add').onclick = () => { R.custom.push({name:'New rule', enabled:true, contains:[], reason:'', next_step:'', priority:2}); render(); };
$('#save').onclick = async () => {
  try { const out = await api('/api/rules', R); R = out.rules; render(); $('#rstatus').textContent = 'Saved. The next run will use these rules.'; }
  catch (e) { $('#rstatus').textContent = 'Could not save: ' + e.message; }
};
$('#run').onclick = async () => {
  const btn = $('#run'), log = $('#log'); btn.disabled = true; say('Running... scans take a few seconds each.'); log.hidden = true; $('#results').hidden = true;
  try {
    await api('/api/rules', R);
    const out = await api('/api/run', {folder: $('#folder').value});
    if (!out.ok) { log.hidden = false; log.textContent = out.output; say('The run hit a problem. Details below.', 'err'); return; }
    showResults(out.summary, out.decisions); say('Done.', 'ok');
  } catch (e) { log.hidden = false; log.textContent = e.message; say('Run failed.', 'err'); }
  finally { btn.disabled = false; }
};
let DEC = {}, SORT = {col: null, dir: 1};
const keyOf = r => r.file + '|' + (r.invoice_number || '') + '|' + (r.amount === '' || r.amount == null ? '' : r.amount);
const isApproved = r => !!DEC[keyOf(r)];
const needsReview = r => r.flag === 'Y' && !isApproved(r);
const COLS = [['flag','Flag'], ['vendor','Vendor'], ['file','File'], ['invoice_number','Invoice #'], ['due_date','Due date'], ['amount','Amount'], ['reason','Reason'], ['','']];
function sortValue(r, col) {
  if (col === 'flag') return r.flag !== 'Y' ? 2 : (isApproved(r) ? 1 : 0);
  if (col === 'amount') return r.amount === '' || r.amount == null ? null : (r.usd_amount != null ? r.usd_amount : r.amount);
  return (r[col] || '').toString().toLowerCase();
}
function sortedRows() {
  const rows = ((LAST && LAST.rows) || []).slice();
  if (!SORT.col) return rows;
  return rows.sort((a, b) => {
    const x = sortValue(a, SORT.col), y = sortValue(b, SORT.col);
    if (x === null && y === null) return 0; if (x === null) return 1; if (y === null) return -1;
    return (x < y ? -1 : x > y ? 1 : 0) * SORT.dir;
  });
}
function drawHead() {
  const tr = h('tr', {});
  COLS.forEach(([col, label]) => {
    const th = h('th', {class: (col === 'amount' ? 'num ' : '') + (col ? 'sortable' : '')}, label + (SORT.col === col && col ? (SORT.dir === 1 ? ' \u25B2' : ' \u25BC') : ''));
    if (col) th.onclick = () => { SORT = SORT.col === col ? {col, dir: -SORT.dir} : {col, dir: 1}; drawAll(); };
    tr.append(th);
  });
  $('#allhead').replaceChildren(tr);
}
async function decide(r, action) {
  try { const out = await api('/api/decide', {key: keyOf(r), action, row: {file: r.file, vendor: r.vendor, invoice_number: r.invoice_number, amount: r.amount}}); DEC = out.decisions || {}; refresh(); }
  catch (e) { say('Could not save that decision: ' + e.message, 'err'); }
}
function approveBtn(r) {
  const b = h('button', {class:'sm'}, 'Clear'); let timer = null;
  b.onclick = () => {
    if (!timer) { b.textContent = 'Confirm?'; timer = setTimeout(() => { b.textContent = 'Clear'; timer = null; }, 4000); return; }
    clearTimeout(timer); b.disabled = true; decide(r, 'approve');
  };
  return b;
}
function drawAll() {
  drawHead();
  const tb = $('#allrows'); tb.replaceChildren();
  const only = $('#onlyflag').checked;
  sortedRows().filter(r => !only || needsReview(r)).forEach(r => {
    const st = r.flag !== 'Y' ? 'N' : (isApproved(r) ? 'A' : 'Y');
    tb.append(h('tr', {class: (st === 'Y' ? (catOf(r) ? catOf(r).cls : 'flagrow') : st === 'A' ? 'apprrow' : '') + (HL && HL.test(r) ? ' hl' : '')},
      h('td', {class: st === 'Y' ? 'fy' : st === 'A' ? 'fa' : 'fn'}, st === 'A' ? 'Cleared' : r.flag),
      h('td', {}, r.vendor), h('td', {}, h('a', {class:'fl', title:'Open this invoice', onclick: () => api('/api/open', {what:'invoice', file: r.file}).catch(e => say(e.message, 'err'))}, r.file)), h('td', {}, r.invoice_number), h('td', {}, r.due_date),
      h('td', {class:'num'}, r.amount === '' ? '' : money(r.amount, r.currency)), h('td', {}, r.reason),
      h('td', {class:'act'}, st === 'Y' ? approveBtn(r) : st === 'A' ? h('button', {class:'sm', onclick: () => decide(r, 'undo')}, 'Undo') : '')));
  });
}
$('#onlyflag').onchange = drawAll;
let HL = null;
function highlight(test) {
  HL = (HL && HL.id === test.id) ? null : test;      // click the same chip again to turn it off
  drawAll();
  if (HL) { const first = document.querySelector('#allrows tr.hl'); if (first) first.scrollIntoView({behavior:'smooth', block:'center'}); }
  document.querySelectorAll('button.chip').forEach(b => b.style.outline = (HL && b.dataset.id === HL.id) ? '2px solid #d97706' : '');
  drawLegend();
}
const CATS = [
  {id:'dup', cls:'c-dup', color:'rgba(147,51,234,.45)', label:'Duplicate (pay once)', re:/duplicate of|also received/},
  {id:'over', cls:'c-math', color:'rgba(245,158,11,.55)', label:'Totals do not add up (overbilling)', re:/line items add/},
  {id:'unr', cls:'c-unr', color:'rgba(107,114,128,.5)', label:'Could not read a field', re:/could not read|unreadable|not readable/i},
  {id:'ent', cls:'c-ent', color:'rgba(13,148,136,.5)', label:'Other company / entity', re:/billed to/},
  {id:'cur', cls:'c-cur', color:'rgba(37,99,235,.45)', label:'Foreign currency', re:/billed in /},
  {id:'lim', cls:'c-lim', color:'rgba(220,38,38,.4)', label:'Over the dollar limit / other', re:/./}];
const catOf = r => r.flag === 'Y' ? CATS.find(c => c.re.test(r.reason || '')) : null;
const T_OVER = {id:'over', test: r => /line items add/.test(r.reason || '')};
const T_DUP = {id:'dup', test: r => /duplicate of|also received/.test(r.reason || '')};
function drawLegend() {
  const lg = $('#legend'); lg.replaceChildren();
  const rows = ((LAST && LAST.rows) || []).filter(r => r.flag === 'Y');
  CATS.forEach(c => {
    const n = rows.filter(r => catOf(r) === c).length; if (!n) return;
    const b = h('button', {'data-id': c.id, title:'Click to highlight these rows', onclick: () => highlight({id: c.id, test: r => catOf(r) === c})},
      h('span', {class:'sw', style:'background:' + c.color}), `${c.label} (${n})`);
    if (HL && HL.id === c.id) b.classList.add('on');
    lg.append(b);
  });
  lg.append(h('span', {style:'display:inline-flex;align-items:center;gap:6px'}, h('span', {class:'sw', style:'background:rgba(22,101,52,.4)'}), 'Cleared by you'));
}
const openQueue = () => (LAST ? LAST.queue : []).filter(q => !isApproved({file: q.file, invoice_number: q.invoice_number, amount: q.amount}));
function drawQueue() {
  const tb = $('#rows'); tb.replaceChildren();
  openQueue().forEach(q => tb.append(h('tr', {},
    h('td', {class: q.priority === 1 ? 'pri1' : 'pri2'}, q.priority === 1 ? 'Check first' : 'Needs approval'),
    h('td', {}, h('div', {}, q.vendor || '(unreadable)'), h('div', {class:'small'}, (q.invoice_number || '') + '  ' + q.file)),
    h('td', {class:'num'}, money(q.amount, q.currency)), h('td', {}, q.why), h('td', {}, q.next_step))));
}
function refresh() {
  const s = LAST, rows = s.rows || [];
  const need = rows.filter(needsReview).length, appr = rows.filter(r => r.flag === 'Y' && isApproved(r)).length;
  $('#headline').textContent = `${s.files} invoices checked: ${need} need a look` + (appr ? `, ${appr} cleared by you` : '') + `, ${s.clear} clear to pay`;
  drawAll(); drawQueue(); drawAp();
}

function apRows() {
  return ((LAST && LAST.rows) || []).filter(r => r.flag !== 'Y' || isApproved(r))
    .map(r => Object.assign({}, r, {status: r.flag === 'Y' ? 'approved' : 'clear'}))
    .sort((a, b) => (a.due_date || '9999').localeCompare(b.due_date || '9999'));
}
function apLine(r) { return `${r.due_date || 'no date'}  ${r.vendor || '(unreadable)'}  ${r.invoice_number || ''}  ${money(r.amount, r.currency)}${r.status === 'approved' ? ' *' : ''}`; }
function apText(rows, limit) {
  const sum = cur => rows.filter(r => (r.currency === 'EUR') === (cur === 'EUR')).reduce((t, r) => t + (Number(r.amount) || 0), 0);
  const eur = sum('EUR');
  const head = `${rows.length} invoices reviewed and ready to pay: ${money(sum('USD'))}` + (eur ? ` plus ${money(eur, 'EUR')}` : '') + ', earliest due first.';
  const foot = ['', '* flagged, then cleared by a reviewer after checking.', 'Nothing has been paid by the review tool. Full list: ap_queue.csv.'];
  const L = [head, ''];
  let used = L.join('\n').length + foot.join('\n').length + 60;
  for (let i = 0; i < rows.length; i++) {
    const line = apLine(rows[i]);
    if (limit && used + line.length + 1 > limit) { L.push(`...and ${rows.length - i} more. See the attached ap_queue.csv.`); break; }
    L.push(line); used += line.length + 1;
  }
  return L.concat(foot).join('\n');
}
function drawAp() {
  const rows = apRows(), open = ((LAST && LAST.rows) || []).filter(needsReview).length;
  const appr = rows.filter(r => r.status === 'approved').length;
  $('#apdesc').textContent = rows.length
    ? `${rows.length} invoices are ready for AP (${rows.length - appr} never flagged, ${appr} flagged then cleared). ` + (open ? `${open} flagged invoices are still open and are left out.` : 'Nothing is left to review.')
    : 'No invoices are ready for AP yet.';
}
try { $('#apto').value = localStorage.getItem('apto') || ''; } catch (e) {}
$('#sendap').onclick = async () => {
  const rows = apRows(), to = $('#apto').value.trim(), note = $('#apnote');
  if (!rows.length) { note.textContent = 'There is nothing ready for AP yet.'; return; }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) { note.textContent = "Type AP's email address first."; return; }
  try { localStorage.setItem('apto', to); } catch (e) {}
  try { await api('/api/ap', {rows: rows.map(r => ({file: r.file, vendor: r.vendor, invoice_number: r.invoice_number, due_date: r.due_date, amount: r.amount, currency: r.currency, status: r.status}))}); }
  catch (e) { note.textContent = 'Could not save the list: ' + e.message; return; }
  const su = `Invoices ready to pay: ${rows.length} (week of ${LAST.run_date})`;
  window.open('https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(to) + '&su=' + encodeURIComponent(su) + '&body=' + encodeURIComponent(apText(rows, 1800)), '_blank', 'noopener');
  note.textContent = 'A draft opened in Gmail. ap_queue.csv is saved in the results folder: attach it, review the draft, then press Send.';
};
$('#copyap').onclick = async () => {
  const rows = apRows(); if (!rows.length) return;
  try { await navigator.clipboard.writeText(apText(rows, 0)); $('#apnote').textContent = 'Copied the full list.'; }
  catch (e) { $('#apnote').textContent = 'Could not copy. Use Send to AP instead.'; }
};
function showResults(s, decisions) {
  LAST = s; DEC = decisions || {}; SORT = {col: null, dir: 1};
  const chips = $('#chips'); chips.replaceChildren();
  HL = null; 
  if (s.dup_exposure) chips.append(h('button', {class:'chip', 'data-id':'dup', title:'Click to highlight these invoices', onclick: () => highlight(T_DUP)}, money(s.dup_exposure) + ' in duplicate invoices caught (show)'));
  if (s.overbilled) chips.append(h('button', {class:'chip', 'data-id':'over', title:'Click to highlight this invoice', onclick: () => highlight(T_OVER)}, money(s.overbilled) + ' of overbilling caught (show)'));
  chips.append(h('span', {class:'chip'}, money(s.usd_clear) + ' clear to pay'));
  refresh(); drawLegend();
  $('#results').hidden = false; $('#results').scrollIntoView({behavior:'smooth'});
}
let LAST = null;
function digestText(s) {
  const L = [`${s.files} invoices checked: ${openQueue().length} need a look, ${s.clear} clear to pay (${money(s.usd_clear)}).`, ''];
  if (s.dup_exposure) L.push(`Caught: ${money(s.dup_exposure)} in duplicate invoices.`);
  if (s.overbilled) L.push(`Caught: ${money(s.overbilled)} of overbilling.`);
  const qs = openQueue(), hi = qs.filter(q => q.priority === 1), med = qs.filter(q => q.priority !== 1);
  if (hi.length) { L.push('', 'CHECK FIRST'); hi.forEach(q => L.push(`- ${q.vendor || 'Unreadable'} ${q.invoice_number || ''} ${money(q.amount, q.currency)}: ${q.next_step}`)); }
  if (med.length) L.push('', `NEEDS APPROVAL (${med.length}): over the dollar limit, foreign currency or another entity. Details are in the spreadsheet.`);
  L.push('', 'Nothing has been paid. Full list: invoice_review.xlsx, Review Queue tab.');
  return L.join('\n');
}
function mailSubject(s) { return `Invoice review ${s.run_date}: ${openQueue().length} need a look, ${s.clear} clear`; }
try { $('#mailto').value = localStorage.getItem('mailto') || ''; } catch (e) {}
$('#emailme').onclick = () => {
  if (!LAST) return;
  const to = $('#mailto').value.trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) { $('#mailnote').textContent = 'Type your email address first.'; return; }
  try { localStorage.setItem('mailto', to); } catch (e) {}
  const url = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(to) + '&su=' + encodeURIComponent(mailSubject(LAST)) + '&body=' + encodeURIComponent(digestText(LAST));
  window.open(url, '_blank', 'noopener');
  $('#mailnote').textContent = 'A draft opened in Gmail. Review it and press Send.';
};
$('#copymail').onclick = async () => {
  if (!LAST) return;
  try { await navigator.clipboard.writeText(mailSubject(LAST) + '\n\n' + digestText(LAST)); $('#mailnote').textContent = 'Copied. Paste it into any email.'; }
  catch (e) { $('#mailnote').textContent = 'Could not copy automatically. Open the Monday email and copy from there.'; }
};
$('#openx').onclick = () => api('/api/open', {what:'xlsx'}).catch(e => say(e.message, 'err'));
$('#opend').onclick = () => api('/api/open', {what:'digest'}).catch(e => say(e.message, 'err'));
$('#openi').onclick = () => api('/api/open', {what:'invdir'}).catch(e => say(e.message, 'err'));
$('#openf').onclick = () => api('/api/open', {what:'folder'}).catch(e => say(e.message, 'err'));
$('#quit').onclick = async () => { try { await api('/api/quit', {}); } catch (e) {} document.body.innerHTML = '<main><h1>Stopped</h1><p class="sub">You can close this tab. Double-click Start Invoice Review to open it again.</p></main>'; };
api('/api/rules').then(d => { R = d.rules; INFO = d.info; $('#folder').value = d.folder || ''; render(); });
</script></body></html>
"""


class Handler(BaseHTTPRequestHandler):
    server_version = "RulesEditor"

    def log_message(self, *a):
        pass

    def _ok_host(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        return host in ("127.0.0.1", "localhost")

    def _send(self, code, body, ctype="application/json"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if not self._ok_host():
            return self._send(403, "forbidden", "text/plain")
        url = urlparse(self.path)
        if url.path == "/api/ls":
            return self._send(200, json.dumps(list_dir((parse_qs(url.query).get("path") or [""])[0])))
        if self.path == "/":
            return self._send(200, PAGE, "text/html")
        if self.path == "/api/rules":
            guess = next((p for p in (os.path.join(HERE, "packet"), os.path.join(os.path.dirname(HERE), "packet")) if os.path.isdir(p)), "")
            info = {k: list(v) for k, v in ip.RULE_INFO.items()}
            return self._send(200, json.dumps({"rules": ip.load_rules(RULES_PATH), "info": info, "folder": guess}))
        self._send(404, "not found", "text/plain")

    def do_POST(self):
        if not self._ok_host() or "application/json" not in (self.headers.get("Content-Type") or ""):
            return self._send(403, "forbidden", "text/plain")
        origin = self.headers.get("Origin")
        if origin and origin.split("//")[-1].split(":")[0] not in ("127.0.0.1", "localhost"):
            return self._send(403, "forbidden", "text/plain")
        try:
            body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
        except ValueError:
            return self._send(400, "bad json", "text/plain")

        if self.path == "/api/rules":
            rules = clean_rules(body if isinstance(body, dict) else {})
            with open(RULES_PATH, "w", encoding="utf-8") as f:
                json.dump(rules, f, indent=2)
            return self._send(200, json.dumps({"rules": rules}))

        if self.path == "/api/run":
            folder = str((body or {}).get("folder", "")).strip().strip('"')
            if not folder or not os.path.isdir(folder):
                return self._send(200, json.dumps({"ok": False, "output": f"Folder not found: {folder or '(empty)'}"}))
            LAST_FOLDER["path"] = folder                    # so a file name in the table can be opened
            summary_path = os.path.join(OUT_DIR, "run_summary.json")
            if os.path.exists(summary_path):
                os.remove(summary_path)                      # never show an old run's results by mistake
            flags = 0x08000000 if os.name == "nt" else 0     # CREATE_NO_WINDOW: no black window flashing up
            proc = subprocess.run([python_exe(), PIPELINE, folder, "--out", OUT_DIR],
                                  capture_output=True, text=True, cwd=HERE, creationflags=flags)
            out = {"ok": proc.returncode == 0 and os.path.exists(summary_path),
                   "output": (proc.stdout + proc.stderr).strip()}
            if out["ok"]:
                with open(summary_path, encoding="utf-8") as f:
                    out["summary"] = json.load(f)
                out["decisions"] = load_decisions()
            elif proc.returncode == 0:
                out["output"] += "\n\nNo results were produced. Does the folder contain invoice files (PDF, JPG, PNG or email)?"
            return self._send(200, json.dumps(out))

        if self.path == "/api/decide":
            key = str((body or {}).get("key", ""))[:400]
            row = (body or {}).get("row") if isinstance((body or {}).get("row"), dict) else {}
            if not key or (body or {}).get("action") not in ("approve", "undo"):
                return self._send(400, "bad request", "text/plain")
            d = load_decisions()
            if body["action"] == "approve":
                d[key] = {"file": str(row.get("file", ""))[:200], "vendor": str(row.get("vendor", ""))[:200],
                          "invoice_number": str(row.get("invoice_number", ""))[:100], "amount": row.get("amount"),
                          "when": datetime.now().isoformat(timespec="seconds")}
            else:
                d.pop(key, None)
            save_decisions(d)
            return self._send(200, json.dumps({"decisions": d}))

        if self.path == "/api/ap":
            rows = (body or {}).get("rows")
            if not isinstance(rows, list) or not rows or len(rows) > 1000:
                return self._send(400, "No invoices to send.", "text/plain")
            os.makedirs(OUT_DIR, exist_ok=True)
            path = os.path.join(OUT_DIR, "ap_queue.csv")
            try:
                with open(path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(["file", "vendor", "invoice_number", "due_date", "amount", "currency", "status"])
                    for r in rows:
                        r = r if isinstance(r, dict) else {}
                        w.writerow([str(r.get(k, ""))[:200] for k in ("file", "vendor", "invoice_number", "due_date", "amount", "currency", "status")])
            except PermissionError:
                return self._send(409, "Close ap_queue.csv (it is open in Excel) and try again.", "text/plain")
            return self._send(200, json.dumps({"file": path, "count": len(rows)}))

        if self.path == "/api/open":
            if (body or {}).get("what") == "invoice":
                src = LAST_FOLDER.get("path")
                name = str(body.get("file", ""))
                full = os.path.realpath(os.path.join(src or "", name))
                if not src or not name or not full.startswith(os.path.realpath(src) + os.sep) or not os.path.isfile(full):
                    return self._send(404, "That invoice file is no longer in the folder.", "text/plain")
                if os.name == "nt":
                    os.startfile(full)                                       # noqa: windows only
                else:
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", full])
                return self._send(200, "{}")
            if (body or {}).get("what") == "invdir":
                if not LAST_FOLDER.get("path"):
                    return self._send(404, "Run the review first.", "text/plain")
                if os.name == "nt":
                    os.startfile(LAST_FOLDER["path"])                        # noqa: windows only
                else:
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", LAST_FOLDER["path"]])
                return self._send(200, "{}")
            target = {"xlsx": "invoice_review.xlsx", "digest": "monday_digest.txt", "ap": "ap_queue.csv", "folder": ""}.get((body or {}).get("what"))
            if target is None:
                return self._send(400, "unknown", "text/plain")
            full = os.path.join(OUT_DIR, target)
            if not os.path.exists(full):
                return self._send(404, "Run the review first.", "text/plain")
            if os.name == "nt":
                os.startfile(full)                                           # noqa: windows only
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", full])
            return self._send(200, "{}")

        if self.path == "/api/quit":
            self._send(200, "{}")
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return
        self._send(404, "not found", "text/plain")


def main():
    port = int(os.environ.get("RULES_PORT", "8765"))
    srv = None
    for p in range(port, port + 10):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            port = p
            break
        except OSError:
            continue
    if not srv:
        sys.exit("Could not open a local port. Close other copies of this window and try again.")
    url = f"http://127.0.0.1:{port}/"
    print(f"Invoice Review is open at {url}\nLeave this window running while you use it. Press Ctrl+C to stop.")
    if not os.environ.get("RULES_NO_BROWSER"):
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
