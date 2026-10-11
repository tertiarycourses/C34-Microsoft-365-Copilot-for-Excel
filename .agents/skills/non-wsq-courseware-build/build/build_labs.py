#!/usr/bin/env python3
"""
Generate one folder per lab from the SAME single source (course_data.py +
data_domainN.py) that drives the PPT, LP and LG:

  labs/lab-NN-<slug>/
    README.md                 lab instructions (Markdown)
    Lab-NN-Instructions.pdf   the same instructions + sample prompts as a PDF
    sample-prompts.txt        copy-ready prompts
    data/<workbook>.xlsx      synthetic mock data (from mock_data.py)

plus labs/README.md as the index. Lab files can never drift out of alignment
with the other artifacts.
"""
import os, re, sys, glob, importlib
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import course_data as C
import mock_data

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (KeepTogether, ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)


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
LABS = os.path.join(REPO, "labs")
os.makedirs(LABS, exist_ok=True)

# ---- collect labs from every data_domainN.py, in domain order
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


def pdf_name(lab):
    return f"Lab-{lab['num']:02d}-Instructions.pdf"


def fence(cmd):
    return "```text\n" + cmd.strip() + "\n```"


# =========================================================================== PDF
def _font(name, candidates, fallback):
    for p in candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(name, p))
                return name
            except Exception:
                pass
    return fallback


SANS = _font("C34Sans", ["/System/Library/Fonts/Supplemental/Arial.ttf", "C:/Windows/Fonts/arial.ttf",
                         "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"], "Helvetica")
SANS_B = _font("C34SansB", ["/System/Library/Fonts/Supplemental/Arial Bold.ttf", "C:/Windows/Fonts/arialbd.ttf",
                            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"], "Helvetica-Bold")
MONO = _font("C34Mono", ["/System/Library/Fonts/Supplemental/Courier New.ttf", "C:/Windows/Fonts/cour.ttf",
                         "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"], "Courier")
if SANS != "Helvetica" and SANS_B != "Helvetica-Bold":
    from reportlab.lib.fonts import addMapping
    addMapping(SANS, 0, 0, SANS)
    addMapping(SANS, 1, 0, SANS_B)

BLUE = colors.HexColor("#1F6FEB")
NAVY = colors.HexColor("#0B3060")
INK = colors.HexColor("#111827")
GREY = colors.HexColor("#555B66")
LIGHT = colors.HexColor("#EEF4FF")
PROMPT_BG = colors.HexColor("#F3F6FA")
GREEN_BG = colors.HexColor("#E8F6EE")
GREEN = colors.HexColor("#1A7F4B")
AMBER_BG = colors.HexColor("#FFF6E5")
LINE = colors.HexColor("#D0D7E2")

ST = {
    "title": ParagraphStyle("t", fontName=SANS_B, fontSize=19, leading=23, textColor=INK, spaceAfter=4),
    "kicker": ParagraphStyle("k", fontName=SANS_B, fontSize=9.5, leading=12, textColor=BLUE, spaceAfter=2),
    "h2": ParagraphStyle("h2", fontName=SANS_B, fontSize=12.5, leading=16, textColor=NAVY, spaceBefore=10, spaceAfter=4),
    "h3": ParagraphStyle("h3", fontName=SANS_B, fontSize=10.5, leading=14, textColor=INK, spaceBefore=6, spaceAfter=2),
    "body": ParagraphStyle("b", fontName=SANS, fontSize=9.6, leading=13.4, textColor=INK, spaceAfter=4, alignment=TA_LEFT),
    "small": ParagraphStyle("s", fontName=SANS, fontSize=8.6, leading=11.5, textColor=INK),
    "smallb": ParagraphStyle("sb", fontName=SANS_B, fontSize=8.6, leading=11.5, textColor=INK),
    "code": ParagraphStyle("c", fontName=MONO, fontSize=8.3, leading=10.6, textColor=NAVY),
    "label": ParagraphStyle("l", fontName=SANS_B, fontSize=8, leading=10, textColor=BLUE),
}


def P(text, style="body"):
    return Paragraph(escape(text), ST[style])


def code_box(text, label=None):
    rows = []
    if label:
        rows.append([Paragraph(escape(label), ST["label"])])
    for line in text.strip("\n").split("\n"):
        lead = len(line) - len(line.lstrip(" "))
        rows.append([Paragraph("&nbsp;" * lead + escape(line.lstrip(" ")) if line.strip() else "&nbsp;", ST["code"])])
    t = Table(rows, colWidths=[170 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PROMPT_BG),
        ("LINEBEFORE", (0, 0), (0, -1), 2.5, BLUE),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 0.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.6),
        ("TOPPADDING", (0, 0), (-1, 0), 5), ("BOTTOMPADDING", (0, -1), (-1, -1), 5),
    ]))
    return t


def callout(title, text, bg, edge):
    t = Table([[Paragraph(escape(title), ST["smallb"])], [Paragraph(escape(text), ST["small"])]], colWidths=[170 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg), ("LINEBEFORE", (0, 0), (0, -1), 2.5, edge),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def grid(header, rows, widths):
    data = [[Paragraph(escape(h), ST["smallb"]) for h in header]]
    data += [[Paragraph(escape(str(c)), ST["small"]) for c in r] for r in rows]
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT), ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def bullets(items):
    return ListFlowable([ListItem(P(i), leftIndent=10, value="•") for i in items], bulletType="bullet", start="•",
                        leftIndent=12, bulletFontName=SANS, bulletFontSize=8)


def build_pdf(lab, topic, path):
    n = lab["num"]

    def on_page(canvas, doc):
        canvas.saveState()
        w, h = A4
        canvas.setFillColor(BLUE)
        canvas.rect(0, h - 9 * mm, w, 9 * mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont(SANS_B, 8.5)
        canvas.drawString(20 * mm, h - 6 * mm, f"{C.TITLE}  ·  {C.COURSE_CODE}")
        canvas.drawRightString(w - 20 * mm, h - 6 * mm, f"Lab {n} of {len(labs)}")
        canvas.setFillColor(GREY)
        canvas.setFont(SANS, 7.5)
        canvas.drawString(20 * mm, 10 * mm, f"{C.ORG}  ·  {C.VERSION} ({C.VERSION_DATE})  ·  Fictional training data")
        canvas.drawRightString(w - 20 * mm, 10 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=17 * mm,
                            bottomMargin=17 * mm, title=f"Lab {n} — {lab['title']}", author=C.ORG,
                            subject=f"{C.TITLE} ({C.COURSE_CODE}) lab instructions and sample prompts")
    S = []
    S.append(Paragraph(escape(f"TOPIC {topic.get('code', '')} · {topic.get('title', '').upper()}"), ST["kicker"]))
    S.append(Paragraph(escape(f"Lab {n} — {lab['title']}"), ST["title"]))
    meta = [
        ("Maps to", lab["objective"]),
        ("Duration", f"{lab.get('duration', '—')} minutes"),
        ("Tools", lab["services"]),
        ("Mock data", ", ".join(lab.get("data_files", [])) or "—"),
        ("Course", f"{C.TITLE} ({C.COURSE_CODE}) · {C.VERSION}"),
    ]
    mt = Table([[Paragraph(escape(k), ST["smallb"]), Paragraph(escape(v), ST["small"])] for k, v in meta],
               colWidths=[26 * mm, 144 * mm])
    mt.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), LIGHT), ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 3),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    S += [Spacer(1, 4), mt]
    S += [Paragraph("Goal", ST["h2"]), P(lab.get("goal", lab["desc"]))]
    S += [Paragraph("What You Will Do", ST["h2"]), P(lab["desc"])]
    S += [Paragraph("What You Will Build", ST["h2"]), P(lab["build"])]
    if lab.get("prerequisites"):
        S += [Paragraph("Prerequisites", ST["h2"]), bullets(lab["prerequisites"])]
    S += [Spacer(1, 6), callout("Data note", C.LAB_NOTE, AMBER_BG, colors.HexColor("#E0A100"))]

    S.append(Paragraph("Steps", ST["h2"]))
    for i, (instr, cmd) in enumerate(lab["steps"], 1):
        block = [Paragraph(f"<b>Step {i}.</b> " + escape(instr), ST["body"])]
        if cmd.strip():
            block.append(code_box(cmd, "PROMPT / CHECK — copy and adapt"))
        block.append(Spacer(1, 5))
        S.append(KeepTogether(block) if len(cmd) < 900 else block[0])
        if len(cmd) >= 900:
            S += block[1:]

    S.append(Paragraph("Sample Prompts", ST["h2"]))
    S.append(P("Use these prompts in the Copilot pane (or adapt them). Name the table and columns, state the output, and "
               "always verify the result. The same prompts are in sample-prompts.txt in this lab folder."))
    sp = lab.get("sample_prompts", [])
    S.append(grid(["#", "Purpose", "Prompt"], [(str(i), lbl, pr) for i, (lbl, pr) in enumerate(sp, 1)],
                  [8 * mm, 28 * mm, 134 * mm]))

    # Keep each short closing section with its heading so no heading is orphaned.
    S.append(KeepTogether([Paragraph("Test It", ST["h2"]), callout("Expected results", lab["test"], GREEN_BG, GREEN)]))
    if lab.get("checkpoint"):
        S.append(KeepTogether([Paragraph("Checkpoint", ST["h2"]), P(lab["checkpoint"])]))
    if lab.get("troubleshooting"):
        S.append(KeepTogether([Paragraph("Troubleshooting", ST["h2"]),
                               grid(["If this happens", "Fix"], lab["troubleshooting"], [58 * mm, 112 * mm])]))
    if lab.get("challenge"):
        S.append(KeepTogether([Paragraph("Challenge", ST["h2"]), P(lab["challenge"])]))
    if lab.get("reflection"):
        S.append(KeepTogether([Paragraph("Reflection", ST["h2"]), P(lab["reflection"])]))
    doc.build(S, onFirstPage=on_page, onLaterPages=on_page)


# =========================================================================== Markdown
def build_md(lab, topic, fname_prev, fname_next):
    n = lab["num"]
    L = [f"# Lab {n} — {lab['title']}\n",
         f"**Course:** {C.TITLE}  ",
         f"**Course Code:** {C.COURSE_CODE}  ",
         f"**Version:** {C.VERSION} ({C.VERSION_DATE})  ",
         f"**Topic {lab['topic']}:** {topic.get('title', '')}  ",
         f"**Maps to:** {lab['objective']}  ",
         f"**Duration:** {lab.get('duration', '—')} minutes  ",
         f"**Tools:** {lab['services']}\n",
         f"> 📄 **Instruction PDF (with sample prompts):** [{pdf_name(lab)}]({pdf_name(lab)})  ",
         "> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  ",
         f"> 🌐 **Prompts web page (one-click copy):** [Lab {n} prompts]({C.PROMPTS_PAGE}#lab-{n})  "]
    for f in lab.get("data_files", []):
        L.append(f"> 📊 **Mock data:** [{f}]({f})  ")
    L.append("")
    L.append("---\n")
    L += ["## Goal\n", lab.get("goal", lab["desc"]) + "\n",
          "## What You Will Do\n", lab["desc"] + "\n",
          "## What You Will Build\n", lab["build"] + "\n",
          "## Prerequisites\n"]
    L += [f"- {x}" for x in lab.get("prerequisites", [])]
    L += ["", f"> **Data note.** {C.LAB_NOTE}\n", "## Steps\n"]
    for i, (instr, cmd) in enumerate(lab["steps"], 1):
        L.append(f"### {i}. {instr}\n")
        if cmd.strip():
            L.append(fence(cmd) + "\n")
    L.append("## Sample Prompts\n")
    L.append("| # | Purpose | Prompt |")
    L.append("|---|---|---|")
    for i, (lbl, pr) in enumerate(lab.get("sample_prompts", []), 1):
        L.append(f"| {i} | {lbl} | {pr.replace('|', '/')} |")
    L.append("")
    L += ["## Test It\n", lab["test"] + "\n"]
    if lab.get("checkpoint"):
        L += ["## Checkpoint and Rejoin Point\n", lab["checkpoint"] + "\n"]
    if lab.get("troubleshooting"):
        L += ["## Troubleshooting\n", "| If this happens | Fix |", "|---|---|"]
        L += [f"| {p} | {f} |" for p, f in lab["troubleshooting"]]
        L.append("")
    if lab.get("challenge"):
        L += ["## Challenge\n", lab["challenge"] + "\n"]
    if lab.get("reflection"):
        L += ["## Reflection\n", lab["reflection"] + "\n"]
    L.append("---\n")
    prev = f"[← Lab {n - 1}](../{fname_prev}/README.md)" if fname_prev else "[← Labs index](../README.md)"
    nxt = f"[Lab {n + 1} →](../{fname_next}/README.md)" if fname_next else "[Labs index →](../README.md)"
    L.append(f"{prev} · {nxt}\n")
    return "\n".join(L)


def build_prompts_txt(lab):
    out = [f"{C.TITLE} ({C.COURSE_CODE}) — Lab {lab['num']}: {lab['title']}",
           "Sample prompts for Copilot in Excel. Copy one at a time, adapt it, and verify every result.", ""]
    for i, (lbl, pr) in enumerate(lab.get("sample_prompts", []), 1):
        out += [f"[{i}] {lbl}", pr, ""]
    out += ["-" * 60, "Step prompts used in this lab", "-" * 60, ""]
    for i, (_instr, cmd) in enumerate(lab["steps"], 1):
        if cmd.strip():
            out += [f"[Step {i}]", cmd.strip(), ""]
    return "\n".join(out)


# =========================================================================== build
written = []
names = [folder_name(l) for l in labs]
for idx, lab in enumerate(labs):
    topic = TOPICS.get(lab["topic"], {})
    fname = names[idx]
    folder = os.path.join(LABS, fname)
    os.makedirs(folder, exist_ok=True)
    mock_data.build_lab_data(lab["num"], folder)
    for rel in lab.get("data_files", []):
        assert os.path.exists(os.path.join(folder, rel)), f"missing mock data {rel} for lab {lab['num']}"
    with open(os.path.join(folder, "README.md"), "w", encoding="utf-8") as f:
        f.write(build_md(lab, topic, names[idx - 1] if idx else None, names[idx + 1] if idx + 1 < len(names) else None))
    with open(os.path.join(folder, "sample-prompts.txt"), "w", encoding="utf-8") as f:
        f.write(build_prompts_txt(lab))
    build_pdf(lab, topic, os.path.join(folder, pdf_name(lab)))
    written.append((lab["num"], fname, lab))

# ---- labs/README.md index
R = [f"# {C.TITLE} — Hands-On Labs\n",
     f"{len(labs)} labs across {len(C.TOPICS)} topics · {C.DAYS} day · "
     f"{C.DAYS * getattr(C, 'INSTRUCTIONAL_MINUTES', C.DAY_MINUTES) / 60:g} instructional hours\n",
     "Every lab has its own folder containing an **instruction PDF with sample prompts**, the same instructions as "
     "`README.md`, a copy-ready `sample-prompts.txt`, and a `data/` folder with the **mock HR / recruitment workbook**. "
     "All data belongs to *BrightPath Talent Solutions*, a fictional Singapore recruitment and staffing firm.\n",
     f"**Prompts web page:** every lab's prompts with one-click Copy buttons — [{C.PROMPTS_PAGE}]({C.PROMPTS_PAGE})\n",
     "Work through the labs in order. Each lab opens its own workbook, so you can rejoin at any lab.\n"]
for t in C.TOPICS:
    R.append(f"\n## Topic {t['num']} — {t['title']}\n")
    R.append("| # | Lab | Instructions | Mock data | You Build |")
    R.append("|---|-----|--------------|-----------|-----------|")
    for n, fname, lab in written:
        if lab["topic"] == t["num"]:
            data = "<br>".join(f"[{os.path.basename(d)}]({fname}/{d})" for d in lab.get("data_files", []))
            R.append(f"| {n} | [{lab['title']}]({fname}/README.md) | [PDF]({fname}/{pdf_name(lab)}) · "
                     f"[prompts]({fname}/sample-prompts.txt) | {data} | {lab['build']} |")
R.append(f"\n---\n\n> {C.LAB_NOTE}\n")
R.append(f"\n_{C.ORG} · {C.COURSE_CODE} · {C.VERSION} ({C.VERSION_DATE})_\n")
with open(os.path.join(LABS, "README.md"), "w", encoding="utf-8") as f:
    f.write("\n".join(R))

print(f"Saved {len(labs)} lab folders (README, PDF, prompts, data) + README.md to {LABS}")
