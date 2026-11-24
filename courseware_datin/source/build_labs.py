"""Build the Datin lab folders and the offline prompt page from shared copy."""

from html import escape
from pathlib import Path
from xml.sax.saxutils import quoteattr

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate

from bog_review import WORKPAPER
from build_bog_workpaper import create as build_bog_workpaper
from course_data import COURSE_CODE, LAB_NOTE, TITLE, TOPICS, VERSION
from labs import LABS
from workbooks import build_workbooks

ROOT = Path(__file__).resolve().parents[1]
NAVY, GOLD = colors.HexColor("#0B1730"), colors.HexColor("#E9BE64")

for name, filename in [("Lab Sans", "Arial.ttf"), ("Lab Sans Bold", "Arial Bold.ttf")]:
    pdfmetrics.registerFont(TTFont(name, "/System/Library/Fonts/Supplemental/" + filename))
pdfmetrics.registerFontFamily("Lab Sans", normal="Lab Sans", bold="Lab Sans Bold")
BODY = ParagraphStyle("Body", fontName="Lab Sans", fontSize=10.5, leading=15, spaceAfter=6, textColor=NAVY)
HEAD = ParagraphStyle("Heading", parent=BODY, fontName="Lab Sans Bold", fontSize=13, leading=16,
                      spaceBefore=8, spaceAfter=5, keepWithNext=True)
TITLE_STYLE = ParagraphStyle("Title", parent=HEAD, fontSize=24, leading=27, spaceBefore=0, spaceAfter=12)
PROMPT = ParagraphStyle("Prompt", parent=BODY, backColor=colors.HexColor("#F2F5F9"),
                        borderPadding=(7, 9, 7, 9), leftIndent=9, rightIndent=9,
                        spaceBefore=5, spaceAfter=12)


def name(lab):
    return f"lab-{lab['num']:02d}-{lab['slug']}"


def filename(lab):
    return f"Lab-{lab['num']:02d}-Instructions.pdf"


def resources(lab):
    return [(f"data/{file}", file) for file in lab["data_files"]] + lab.get("downloads", []) + [
        (rel, desc.title()) for rel, _, desc in lab.get("extra_files", [])
    ]


def reference_links(lab):
    return [(url, label) for label, url in lab["links"]]


def build_pdf(lab, folder):
    compact = lab["num"] in {1, 2}
    body_style = ParagraphStyle("Compact body", parent=BODY, leading=13, spaceAfter=3) if compact else BODY
    heading_style = ParagraphStyle("Compact heading", parent=HEAD, spaceBefore=6, spaceAfter=3) if compact else HEAD
    title_style = ParagraphStyle("Compact title", parent=TITLE_STYLE, fontSize=20, leading=23) if compact else TITLE_STYLE

    def page(canvas, doc):
        width, height = A4
        canvas.setFillColor(NAVY)
        canvas.rect(0, height - 11 * mm, width, 11 * mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Lab Sans Bold", 9.5)
        canvas.drawString(19 * mm, height - 7 * mm, TITLE)
        canvas.drawRightString(width - 19 * mm, height - 7 * mm, f"Datin · Lab {lab['num']}")
        canvas.setStrokeColor(GOLD)
        canvas.line(19 * mm, 15 * mm, width - 19 * mm, 15 * mm)
        canvas.setFillColor(NAVY)
        canvas.setFont("Lab Sans", 9)
        canvas.drawString(19 * mm, 10 * mm, f"{COURSE_CODE} · Datin edition {VERSION}")
        canvas.drawRightString(width - 19 * mm, 10 * mm, str(doc.page))

    def body(text):
        return Paragraph(escape(text), body_style)

    def heading(text):
        return Paragraph(escape(text), heading_style)

    def link(label, url):
        return f'<link href={quoteattr(url)} color="#174D91"><u>{escape(label)}</u></link>'

    document = SimpleDocTemplate(
        str(folder / filename(lab)), pagesize=A4, topMargin=18 * mm, bottomMargin=19 * mm,
        leftMargin=19 * mm, rightMargin=19 * mm,
        title=f"{TITLE} | Lab {lab['num']} | {lab['title']}",
    )
    prompt_page = f"../../prompts/index.html#lab-{lab['num']}"
    flow = [
        body(f"TOPIC {lab['topic']} · {TOPICS[lab['topic'] - 1]['title']}"),
        Paragraph(escape(lab["title"]), title_style),
        body(f"{lab['duration']} minutes · {lab['services']}"),
        body(lab["goal"]),
        heading("What you’ll build"), body(lab["build"]),
        heading("Before you start"), *map(body, lab["prerequisites"]),
        Paragraph(
            f"Copy-ready prompts are in sample-prompts.txt and in {link('the Datin prompt page', prompt_page)}. "
            "Open prompts/index.html from the extracted course folder if your PDF reader blocks local links.",
            BODY,
        ),
        *[Paragraph(link(label, url), BODY) for label, url in lab["links"]],
    ]
    for rel, label in resources(lab):
        flow.append(Paragraph(f'{link(label, rel)} <font size="9">({escape(rel)})</font>', BODY))

    labels = {prompt: label for label, prompt in lab["sample_prompts"]}
    for index, (instruction, prompt) in enumerate(lab["steps"], start=1):
        block = [Paragraph(f"<b>{index}.</b> {escape(instruction)}", body_style)]
        if prompt:
            block.append(Paragraph(f"<b>{escape(labels[prompt])}</b><br/>{escape(prompt)}", PROMPT))
        if index == 1:
            block.insert(0, heading("Steps"))
        flow.append(KeepTogether(block))
    flow.extend([
        heading("Test it"),
        *[Paragraph(f"<b>{escape(label)}</b> · {escape(result)}", body_style) for label, result in lab["checks"]],
        heading("Checkpoint"), body(lab["checkpoint"]),
        heading("Troubleshooting"),
        *[Paragraph(f"<b>{escape(issue)}</b> · {escape(fix)}", body_style) for issue, fix in lab["troubleshooting"]],
        heading("Challenge"), body(lab["challenge"]),
        heading("Reflection"), body(lab["reflection"]),
    ])
    document.build(flow, onFirstPage=page, onLaterPages=page)


def lab_files(lab, folder):
    outputs = [f"{TITLE} · Lab {lab['num']} · {lab['title']}", ""]
    for label, prompt in lab["sample_prompts"]:
        outputs.extend([label, prompt, ""])
    (folder / "sample-prompts.txt").write_text("\n".join(outputs), encoding="utf-8")
    for rel, contents, _ in lab.get("extra_files", []):
        path = folder / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(contents, encoding="utf-8")
    readme = [
        f"# Lab {lab['num']} — {lab['title']}", "",
        f"[Instructions PDF]({filename(lab)}) · [Copy prompts](../../prompts/index.html#lab-{lab['num']})", "",
        "## Goal", "", lab["goal"], "", "## Prerequisites", "",
        *[f"- {item}" for item in lab["prerequisites"]], "", "## Steps", "",
    ]
    for index, (instruction, prompt) in enumerate(lab["steps"], start=1):
        readme.extend([f"### {index}. {instruction}", ""])
        if prompt:
            readme.extend(["```text", prompt, "```", ""])
    readme.extend(["## Test it", "", *[f"- **{a}** · {b}" for a, b in lab["checks"]], "",
                   "## Checkpoint", "", lab["checkpoint"], "", "## Files", ""])
    readme.extend(f"- [{label}]({rel})" for rel, label in resources(lab) + reference_links(lab))
    (folder / "README.md").write_text("\n".join(readme) + "\n", encoding="utf-8")
    build_pdf(lab, folder)


def lab_panel(lab):
    base = f"../labs/{name(lab)}"
    links = [(f"{base}/{filename(lab)}", "Instructions PDF")] + [
        (f"{base}/{rel}", label) for rel, label in resources(lab)
    ] + reference_links(lab)
    links = "".join(f'<li><a href="{escape(url, quote=True)}">{escape(label)}</a></li>' for url, label in links)
    prompts = "".join(f"""
    <article>
      <header><h3>{escape(label)}</h3><button type="button" aria-label="Copy {escape(label, quote=True)}">Copy</button></header>
      <pre>{escape(prompt)}</pre>
    </article>""" for label, prompt in lab["sample_prompts"])
    return f"""
  <section id="lab-{lab['num']}">
    <h2>Lab {lab['num']} · {escape(lab['title'])}</h2>
    <p>{escape(lab['desc'])}</p>
    <ul class="resources">{links}</ul>
{prompts}
  </section>"""


def build_page():
    nav = "".join(f'<a href="#lab-{lab["num"]}">Lab {lab["num"]} · {escape(lab["theme"])}</a>' for lab in LABS)
    html = (Path(__file__).with_name("prompts.html")).read_text(encoding="utf-8")
    (ROOT / "prompts" / "index.html").write_text(
        html.format(title=escape(TITLE), nav=nav, panels="\n".join(map(lab_panel, LABS))), encoding="utf-8"
    )


def main():
    folders = {lab["num"]: ROOT / "labs" / name(lab) for lab in LABS}
    for folder in folders.values():
        folder.mkdir(parents=True, exist_ok=True)
    build_workbooks(folders)
    build_bog_workpaper(folders[4] / WORKPAPER)
    index = [f"# {TITLE} labs", "", "[Open the copy-ready prompt page](../prompts/index.html)", "", LAB_NOTE, ""]
    for lab in LABS:
        folder = folders[lab["num"]]
        lab_files(lab, folder)
        index.append(f"- [Lab {lab['num']} · {lab['title']}]({folder.name}/README.md)")
    (ROOT / "labs" / "README.md").write_text("\n".join(index) + "\n", encoding="utf-8")
    build_page()


if __name__ == "__main__":
    main()
