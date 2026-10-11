#!/usr/bin/env python3
"""
Generate the copy-ready Lab Prompts web page (GitHub Pages) from the SAME
single source (course_data.py + data_domainN.py) that drives the lab READMEs:

  docs/prompts/index.html   one section per lab: steps, prompt cards with a
                            Copy button, check formulas, sample prompts, Test It
  docs/index.html           redirect to prompts/
  docs/.nojekyll

Serve with GitHub Pages from the main branch /docs folder.
"""
import datetime, glob, html, importlib, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import course_data as C


def find_repo(start):
    env = os.environ.get("COURSE_REPO")
    if env and os.path.isdir(env):
        return env
    d = start
    for _ in range(8):
        d = os.path.dirname(d)
        if os.path.isdir(os.path.join(d, "courseware")) and os.path.isdir(os.path.join(d, "labs")):
            return d
    return os.path.dirname(os.path.dirname(start))


REPO = find_repo(HERE)
DOCS = os.path.join(REPO, "docs")
GH = f"https://github.com/{C.GITHUB_REPO}"

labs = []
for path in sorted(glob.glob(os.path.join(HERE, "data_domain[0-9]*.py"))):
    mod = importlib.import_module(os.path.basename(path)[:-3])
    for name in dir(mod):
        if re.fullmatch(r"DOMAIN\d+", name):
            labs.extend(getattr(mod, name))
labs.sort(key=lambda l: l["num"])
TOPICS = {t["num"]: t for t in C.TOPICS}


def slug(s):
    return re.sub(r"-+", "-", re.sub(r"[^a-z0-9]+", "-", s.lower())).strip("-")


def folder_name(lab):
    return f"lab-{lab['num']:02d}-{slug(lab['title'])}"


# =========================================================================== parse step blocks
# A step's code block holds prompts, check formulas or reference info. Labelled
# paragraphs ("CHAT ONLY:", "CHECK:", "FOLLOW-UP:" ...) say which.
LABEL = re.compile(r"^(CHAT ONLY|ALLOW EDITING|PROMPT|CHECK FORMULAS|CHECK|FOLLOW-UP EXAMPLE|FOLLOW-UP|"
                   r"OPTION [A-Z]|Then|Expected after refresh)([^:\n]*):[ \t]*(.*)$")
MODES = [("advanced analysis mode", "Advanced analysis"), ("Chat only", "Chat only"),
         ("Allow editing", "Allow editing"), (r"\bPlan\b", "Plan")]
INFO_LINE = re.compile(r"^[A-Za-z_][A-Za-z_ ]*\s+:\s")      # "Table name : tblApplications"
FORMULA_LINE = re.compile(r"^(.*?\S)\s+(=\S.*)$")           # "Active headcount   =COUNTIFS(...)"
CELL_PAIR = re.compile(r"^(A\d+)\s+(.*?)\s+(B\d+)\s+(=.*)$")  # "A3 Total fee   B3 =SUM(...)"


STRUCT = re.compile(r"^(\(?\d+\)\s|[A-Z][A-Za-z_]*( \d+)?:|- |=)")


def unwrap(body):
    """Undo the source's ~100-column hard wraps; keep structural breaks (labels, numbered/indented items)."""
    out = []
    for line in body.split("\n"):
        s, indent = line.lstrip(), len(line) - len(line.lstrip())
        structural = (not out or not s or out[-1].rstrip().endswith(":") or STRUCT.match(s)
                      or (1 <= indent <= 4 and s[0].isupper()))
        if structural:
            out.append(line.rstrip())
        else:
            out[-1] = out[-1].rstrip() + " " + s.rstrip()
    return "\n".join(out)


def modes_in(text):
    hits = sorted((m.start(), name) for pat, name in MODES for m in re.finditer(pat, text))
    return [name for _, name in hits]


def split_formulas(body):
    """One item per line starting with '='; indented continuation lines are joined."""
    out = []
    for line in body.split("\n"):
        s = line.strip()
        if not s:
            continue
        if s.startswith("=") or not out:
            out.append(s)
        else:
            out[-1] += " " + s
    return [("", f) for f in out]


def setup_formulas(body):
    out = []
    for line in body.split("\n"):
        m = CELL_PAIR.match(line.strip())
        if m:
            out.append((f"{m.group(1)} {m.group(2)} · {m.group(3)}", m.group(4)))
        else:
            m = FORMULA_LINE.match(line.strip())
            out.append((m.group(1), m.group(2)))
    return out


def split_changes(body):
    """'One change at a time' blocks: one prompt per numbered item or per line."""
    if re.search(r"^\d\)\s", body, re.M):
        parts = re.split(r"^\d\)\s+", body, flags=re.M)
        return [re.sub(r"\s*\n\s*", " ", p).strip() for p in parts if p.strip()]
    return [l.strip() for l in body.split("\n") if l.strip()]


def parse_step(num, instr, cmd, mode):
    """Return (items, carried_mode). Items: ('prompt', tag, text) | ('formulas', tag, [(label, f)])
    | ('info', tag, text)."""
    found = modes_in(instr)
    step_mode = found[0] if found else mode
    if re.search(r"^(CHAT ONLY|ALLOW EDITING):", cmd, re.M):   # explicit labels mark the exceptions
        step_mode = mode or step_mode
    carry = next((m for m in reversed(found) if m != "Advanced analysis"), mode)
    items = []
    if not cmd.strip():
        return items, carry
    one_by_one = re.search(r"one change at a time|one prompt per change", instr)
    for para in re.split(r"\n\s*\n", cmd.strip("\n")):
        lines = para.rstrip().split("\n")
        m = LABEL.match(lines[0].strip())
        label = (m.group(1) + m.group(2)).strip() if m else None
        body = "\n".join(([m.group(3)] if m and m.group(3) else []) + lines[1:]) if m else para
        body = "\n".join(l.rstrip() for l in body.split("\n")).strip("\n")
        if label is None and one_by_one:
            for chunk in split_changes(body):
                cm = LABEL.match(chunk)
                if cm:
                    items.append(("prompt", cm.group(1).capitalize(), cm.group(3)))
                else:
                    items.append(("prompt", step_mode, chunk))
            continue
        if label is None:
            if num == 1 or INFO_LINE.match(lines[0]):
                items.append(("info", "Reference", body))
            elif all(FORMULA_LINE.match(l.strip()) or l.strip().startswith("=") for l in lines if l.strip()):
                items.append(("formulas", "Type these yourself", setup_formulas(body)))
            elif body.startswith("If ") and 'ask: "' in body:
                cond, ask = body.split('ask: "', 1)
                items.append(("prompt", "If needed · " + cond[3:].strip(" ,"), ask.rstrip('"')))
            else:
                items.append(("prompt", step_mode, body))
        elif label.startswith("CHECK"):
            sub = label[len("CHECK FORMULAS") if label.startswith("CHECK FORMULAS") else len("CHECK"):].strip()
            items.append(("formulas", "Check " + sub if sub else "Check", split_formulas(body)))
        elif label.startswith("Expected"):
            items.append(("info", label, body))
        elif label.startswith("OPTION"):
            if body.lstrip().startswith("="):
                items.append(("formulas", "Option" + label[6:], [("", re.sub(r"\s*\n\s*", " ", body.strip()))]))
            else:
                items.append(("prompt", "Option" + label[6:], body))
        elif label in ("CHAT ONLY", "ALLOW EDITING"):
            items.append(("prompt", label.capitalize(), body))
        elif label.startswith("PROMPT"):
            pm = re.search(r"\(([^)]*)\)", label)
            items.append(("prompt", pm.group(1) if pm else step_mode, body))
        else:  # Then / FOLLOW-UP*
            items.append(("prompt", "Follow-up" if label == "Then" else label.capitalize(), body))
    return items, carry


# =========================================================================== render
def esc(s):
    return html.escape(s, quote=False)


def marked(s):
    """Escape, then highlight <placeholders> the learner must replace."""
    return re.sub(r"&lt;([a-z][^&\n]*?)&gt;", r"<mark>&lt;\1&gt;</mark>", esc(s))


def attr(s):
    return html.escape(s, quote=True)


def copy_btn(text, small=False):
    cls = "copy small" if small else "copy"
    return f'<button class="{cls}" type="button" data-copy="{attr(text)}">Copy</button>'


LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def render_lab(lab):
    n, fname = lab["num"], folder_name(lab)
    topic = TOPICS.get(lab["topic"], {})
    base_blob, base_raw = f"{GH}/blob/main/labs/{fname}", f"{GH}/raw/main/labs/{fname}"
    files = [f'<a href="{base_blob}/Lab-{n:02d}-Instructions.pdf">Instructions PDF</a>',
             f'<a href="{base_blob}/README.md">Lab page on GitHub</a>']
    files += [f'<a href="{base_raw}/{d}">Download {esc(os.path.basename(d))}</a>' for d in lab.get("data_files", [])]

    mode, steps_html, blocks, has_ph, k = None, [], [], False, 0
    for i, (instr, cmd) in enumerate(lab["steps"], 1):
        items, mode_next = parse_step(i, instr, cmd, mode)
        info = "".join(f'<pre class="info">{esc(t)}</pre>' for kind, _, t in items if kind == "info")
        prompts_here = sum(1 for it in items if it[0] == "prompt")
        tail = (f' <span class="see">→ Prompt{"s" if prompts_here > 1 else ""} '
                f'{", ".join(LETTERS[k + j] for j in range(prompts_here))}</span>') if prompts_here else ""
        steps_html.append(f"<li>{esc(instr)}{tail}{info}</li>")
        for kind, tag, body in items:
            if kind == "prompt":
                body = unwrap(body)
                has_ph |= "<mark>" in marked(body)
                blocks.append(
                    f'<article class="prompt"><header><div><span class="pname">Lab {n} · Prompt {LETTERS[k]}</span>'
                    f'<span class="ptool">Step {i} · {esc(tag)}</span></div>{copy_btn(body)}</header>'
                    f"<pre>{marked(body)}</pre></article>")
                k += 1
            elif kind == "formulas":
                rows = []
                for lbl, f in body:
                    has_ph |= "<mark>" in marked(f)
                    lab_html = f'<span class="when">{esc(lbl)}</span>' if lbl else ""
                    rows.append(f'<li>{lab_html}<code class="q">{marked(f)}</code>{copy_btn(f, True)}</li>')
                blocks.append(
                    f'<article class="prompt formulas"><header><div><span class="pname">Lab {n} · Step {i}</span>'
                    f'<span class="ptool">{esc(tag)}{"" if tag.startswith("Option") else " — paste into an empty cell"}</span></div></header>'
                    f'<ul>{"".join(rows)}</ul></article>')
        mode = mode_next

    asks = "".join(
        f'<li><span class="when">{esc(lbl)}</span><span class="q">{esc(p)}</span>{copy_btn(p, True)}</li>'
        for lbl, p in lab.get("sample_prompts", []))
    hint = ('<p class="hint">Replace the <mark>highlighted</mark> parts before you send.</p>' if has_ph else "")
    return (
        f'<section class="lab" data-lab="{n}" id="lab-{n}-section">'
        f'<p class="eyebrow">Lab {n} · Topic {lab["topic"]} · {lab.get("duration", "—")} min</p>'
        f'<h2>{esc(lab["title"])}</h2>'
        f'<p class="use"><b>Topic {lab["topic"]}:</b> {esc(topic.get("title", ""))}<br>'
        f'<b>Use:</b> {esc(lab["services"])}</p>'
        f'<p class="files">{" · ".join(files)}</p>'
        f'<details><summary>The steps</summary><ol>{"".join(steps_html)}</ol></details>'
        f"<h3>The prompts</h3>{hint}{''.join(blocks)}"
        f'<section class="ask"><h3>More prompts to try</h3><p>Copy one at a time, adapt it, and verify '
        f"every result with a formula.</p><ul>{asks}</ul></section>"
        f'<details class="test"><summary>Test it — what you should see</summary><p>{esc(lab["test"])}</p></details>'
        f"</section>")


CSS = """
:root{--bg:#F7F9F7;--panel:#fff;--ink:#1D211F;--muted:#5D6762;--rule:#D5DDD8;
--accent:#107C41;--on-accent:#fff;--accent-bg:#EEF6F1;--ok:#185ABD;--mark:#FCE9B8;--code:#F1F4F2}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
--bg:#161A18;--panel:#1F2421;--ink:#ECF0EE;--muted:#AAB5AF;--rule:#36403B;
--accent:#4CC38A;--on-accent:#0E1A13;--accent-bg:#1E2B24;--ok:#7AB4FF;--mark:#5A4A1E;--code:#262C29}}
:root[data-theme="dark"]{--bg:#161A18;--panel:#1F2421;--ink:#ECF0EE;--muted:#AAB5AF;--rule:#36403B;
--accent:#4CC38A;--on-accent:#0E1A13;--accent-bg:#1E2B24;--ok:#7AB4FF;--mark:#5A4A1E;--code:#262C29}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
font:17px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Calibri,sans-serif}
.wrap{max-width:860px;margin:0 auto;padding:24px 16px 80px}
a{color:var(--accent)}
h1,h2{font-family:Cambria,Georgia,serif;font-weight:600;line-height:1.2}
h1{font-size:30px;margin:4px 0 6px}h2{font-size:26px;margin:4px 0 8px}
h3{font-size:15px;letter-spacing:.06em;text-transform:uppercase;color:var(--accent);margin:28px 0 10px}
.eyebrow{margin:0;color:var(--accent);font-size:13px;font-weight:700;letter-spacing:.08em;text-transform:uppercase}
.sub{color:var(--muted);margin:0 0 18px}
nav{position:sticky;top:0;z-index:2;background:var(--bg);padding:10px 0;
border-bottom:1px solid var(--rule);display:flex;flex-wrap:wrap;gap:6px}
@media (max-width:600px){nav{flex-wrap:nowrap;overflow-x:auto}nav a{flex:none}}
nav a{display:inline-block;min-width:44px;text-align:center;padding:8px 10px;
border:1.5px solid var(--rule);border-radius:999px;color:var(--ink);text-decoration:none;font-weight:600;font-size:15px}
nav a[aria-current="true"]{background:var(--accent);border-color:var(--accent);color:var(--on-accent)}
.lab{padding-top:22px}
.use{color:var(--muted);margin:0 0 6px}
.files{margin:0 0 12px;font-size:15px}
details{background:var(--panel);border:1px solid var(--rule);border-radius:10px;padding:8px 14px;margin:0 0 6px}
summary{cursor:pointer;font-weight:600}
details ol{margin:8px 0 4px;padding-left:22px}details li{margin:8px 0}
details.test{margin-top:24px}details.test p{margin:8px 0 4px}
.see{color:var(--accent);font-weight:700;font-size:15px;white-space:nowrap}
pre.info{margin:6px 0 2px;padding:8px 10px;background:var(--code);border-radius:6px;font-size:13px;color:var(--muted)}
.hint{margin:0 0 10px;color:var(--muted);font-size:15px}
mark{background:var(--mark);color:inherit;border-radius:3px;padding:0 2px}
.prompt{background:var(--panel);border:1px solid var(--rule);border-left:4px solid var(--accent);
border-radius:10px;margin:0 0 14px;overflow:hidden}
.prompt header{display:flex;justify-content:space-between;align-items:center;gap:12px;
padding:10px 14px;background:var(--accent-bg)}
.pname{font-weight:700;margin-right:8px}.ptool{color:var(--muted);font-size:15px}
pre{margin:0;padding:14px;white-space:pre-wrap;word-wrap:break-word;
font:15px/1.5 ui-monospace,Menlo,Consolas,monospace}
.copy{flex:none;font:inherit;font-weight:700;font-size:15px;color:var(--on-accent);background:var(--accent);
border:0;border-radius:8px;padding:9px 18px;cursor:pointer;min-height:40px}
.copy:hover{filter:brightness(1.08)}
.copy.done{background:var(--ok)}
.copy.small{font-size:13px;padding:5px 12px;min-height:32px}
.formulas{border-left-color:var(--ok)}
.formulas ul,.ask ul{list-style:none;padding:0;margin:0}
.formulas li{display:flex;flex-wrap:wrap;align-items:center;gap:6px 12px;padding:10px 14px;border-top:1px solid var(--rule)}
.formulas li:first-child{border-top:0}
code.q{flex:1 1 300px;font:14px/1.5 ui-monospace,Menlo,Consolas,monospace;word-break:break-word}
.ask li{display:flex;flex-wrap:wrap;align-items:center;gap:8px 12px;background:var(--panel);
border:1px solid var(--rule);border-radius:10px;padding:10px 14px;margin:0 0 8px}
.ask .q{flex:1 1 300px}
.when{flex-basis:100%;font-size:13px;font-weight:700;color:var(--accent)}
.ask p{color:var(--muted);margin:0 0 10px;font-size:15px}
footer{margin-top:40px;color:var(--muted);font-size:14px}
.pick{margin:28px 0;padding:18px 20px;border:2px dashed var(--accent);border-radius:12px}
.pick p{font-size:18px;font-weight:600;margin:0 0 10px}
.pick h4{margin:14px 0 4px;font-size:14px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
.pick ul{margin:0;padding-left:20px}.pick li{margin:4px 0}
"""

JS = """
const labs=[...document.querySelectorAll('.lab')];
const links=[...document.querySelectorAll('nav a')];
function show(){
  const m=location.hash.match(/^#lab-(\\d+)$/);
  const n=m?m[1]:null;
  document.getElementById('pick').hidden=!!n;
  labs.forEach(s=>s.hidden=s.dataset.lab!==n);
  links.forEach(a=>a.setAttribute('aria-current',a.getAttribute('href')==='#lab-'+n));
}
addEventListener('hashchange',()=>{show();scrollTo(0,0)});
show();addEventListener('load',()=>scrollTo(0,0));
function fallback(t){const a=document.createElement('textarea');a.value=t;
  a.style.position='fixed';a.style.opacity='0';document.body.appendChild(a);
  a.select();try{document.execCommand('copy')}catch(e){}a.remove()}
document.addEventListener('click',async e=>{
  const b=e.target.closest('.copy');if(!b)return;
  const t=b.dataset.copy;
  try{await navigator.clipboard.writeText(t)}catch(err){fallback(t)}
  const old=b.textContent;b.textContent='Copied \\u2713';b.classList.add('done');
  setTimeout(()=>{b.textContent=old;b.classList.remove('done')},1800);
});
"""


def build_page():
    nav = "".join(f'<a href="#lab-{l["num"]}">{l["num"]}</a>' for l in labs)
    pick = ['<div id="pick" class="pick"><p>Which lab are you on? Click its number above — every prompt is '
            "labelled with its lab, so check it before you copy.</p>"]
    for t in C.TOPICS:
        pick.append(f'<h4>Topic {t["num"]} · {esc(t["title"])}</h4><ul>')
        pick += [f'<li><a href="#lab-{l["num"]}">Lab {l["num"]} — {esc(l["title"])}</a></li>'
                 for l in labs if l["topic"] == t["num"]]
        pick.append("</ul>")
    pick.append("</div>")
    today = datetime.date.today().strftime("%-d %B %Y")
    return f"""<!doctype html>
<html lang="en-SG"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="data:,">
<title>Lab Prompts · Copilot in Excel</title>
<meta name="description" content="Copy-ready prompts for every lab of {esc(C.TITLE)} ({C.COURSE_CODE}).">
<style>{CSS}</style></head><body><div class="wrap">
<p class="eyebrow">{C.COURSE_CODE} · Copilot in Excel · {C.VERSION}</p>
<h1>Lab Prompts</h1>
<p class="sub">{esc(C.TITLE)}. Pick your lab, click <b>Copy</b>, then paste into the Copilot pane in Excel
(or into a cell, for check formulas). Verify every result. <a href="{GH}">Course repository</a>.</p>
<nav aria-label="Labs">{nav}</nav>
{"".join(pick)}
{"".join(render_lab(l) for l in labs)}
<footer>Built from the course source on {today}. BrightPath Talent Solutions and all lab data are
fictitious, for training only. Use only data you are authorised to process. © {esc(C.ORG)}.</footer>
</div><script>{JS}</script></body></html>
"""


os.makedirs(os.path.join(DOCS, "prompts"), exist_ok=True)
with open(os.path.join(DOCS, "prompts", "index.html"), "w", encoding="utf-8") as f:
    f.write(build_page())
with open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8") as f:
    f.write('<!doctype html><meta charset="utf-8"><title>Lab Prompts</title>'
            '<meta http-equiv="refresh" content="0; url=prompts/"><a href="prompts/">Lab Prompts</a>\n')
open(os.path.join(DOCS, ".nojekyll"), "w").close()
print(f"Saved prompts page for {len(labs)} labs to {os.path.join(DOCS, 'prompts', 'index.html')}")
