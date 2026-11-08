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
import json, os, subprocess, sys, threading, webbrowser
from urllib.parse import parse_qs, urlparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import invoice_pipeline as ip

HERE = os.path.dirname(os.path.abspath(__file__))
RULES_PATH = os.path.join(HERE, "rules.json")
PIPELINE = os.path.join(HERE, "invoice_pipeline.py")
OUT_DIR = os.path.join(HERE, "out")


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
main{max-width:900px;margin:0 auto;padding:24px 16px 60px}
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
  <table><thead><tr><th>Urgency</th><th>Invoice</th><th class="num">Amount</th><th>Why it was flagged</th><th>Suggested next step</th></tr></thead><tbody id="rows"></tbody></table>
  <div class="inline" style="margin-top:14px">
    <button id="openx" class="primary">Open the spreadsheet</button>
    <button id="opend">Open the Monday email</button>
    <button id="openf">Open the results folder</button>
  </div>
  <p class="small" style="margin-bottom:0">The spreadsheet has a tab for every invoice, a review queue with a decision dropdown, a summary, and the rules used.</p>
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
function sel(val, set) { const s = h('select', {}, h('option', {value:'1'}, 'High: check before anything else'), h('option', {value:'2'}, 'Medium: needs approval')); s.value = String(val); s.onchange = () => set(s.value); return s; }

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
    showResults(out.summary); say('Done.', 'ok');
  } catch (e) { log.hidden = false; log.textContent = e.message; say('Run failed.', 'err'); }
  finally { btn.disabled = false; }
};
function showResults(s) {
  LAST = s;
  $('#headline').textContent = `${s.files} invoices checked: ${s.flagged} need a look, ${s.clear} clear to pay`;
  const chips = $('#chips'); chips.replaceChildren();
  if (s.dup_exposure) chips.append(h('span', {class:'chip'}, money(s.dup_exposure) + ' in duplicate invoices caught'));
  if (s.overbilled) chips.append(h('span', {class:'chip'}, money(s.overbilled) + ' of overbilling caught'));
  chips.append(h('span', {class:'chip'}, money(s.usd_clear) + ' clear to pay'));
  const tb = $('#rows'); tb.replaceChildren();
  s.queue.forEach(q => tb.append(h('tr', {},
    h('td', {class: q.priority === 1 ? 'pri1' : 'pri2'}, q.priority === 1 ? 'High' : 'Medium'),
    h('td', {}, h('div', {}, q.vendor || '(unreadable)'), h('div', {class:'small'}, (q.invoice_number || '') + '  ' + q.file)),
    h('td', {class:'num'}, money(q.amount, q.currency)), h('td', {}, q.why), h('td', {}, q.next_step))));
  $('#results').hidden = false; $('#results').scrollIntoView({behavior:'smooth'});
}
let LAST = null;
function digestText(s) {
  const L = [`${s.files} invoices checked: ${s.flagged} need a look, ${s.clear} clear to pay (${money(s.usd_clear)}).`, ''];
  if (s.dup_exposure) L.push(`Caught: ${money(s.dup_exposure)} in duplicate invoices.`);
  if (s.overbilled) L.push(`Caught: ${money(s.overbilled)} of overbilling.`);
  const hi = s.queue.filter(q => q.priority === 1), med = s.queue.filter(q => q.priority !== 1);
  if (hi.length) { L.push('', 'CHECK THESE FIRST'); hi.forEach(q => L.push(`- ${q.vendor || 'Unreadable'} ${q.invoice_number || ''} ${money(q.amount, q.currency)}: ${q.next_step}`)); }
  if (med.length) L.push('', `NEEDS APPROVAL (${med.length}): over the dollar limit, foreign currency or another entity. Details are in the spreadsheet.`);
  L.push('', 'Nothing has been paid. Full list: invoice_review.xlsx, Review Queue tab.');
  return L.join('\n');
}
function mailSubject(s) { return `Invoice review ${s.run_date}: ${s.flagged} need a look, ${s.clear} clear`; }
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
            elif proc.returncode == 0:
                out["output"] += "\n\nNo results were produced. Does the folder contain invoice files (PDF, JPG, PNG or email)?"
            return self._send(200, json.dumps(out))

        if self.path == "/api/open":
            target = {"xlsx": "invoice_review.xlsx", "digest": "monday_digest.txt", "folder": ""}.get((body or {}).get("what"))
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
