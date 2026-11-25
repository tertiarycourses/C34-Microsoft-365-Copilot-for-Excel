"""Generate the editable Topic 4 financial-statement review workpaper."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt, RGBColor

import bog_review as B

NAVY = "0B1730"
BLUE = "174882"
GOLD = "A16A27"
SOFT = "EFF3F8"


def link(paragraph, label, url):
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True))
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    properties.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.append(underline)
    run.append(properties)
    value = OxmlElement("w:t")
    value.text = label
    run.append(value)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def table(doc, headings, widths, rows):
    grid = doc.add_table(rows=1, cols=len(headings))
    grid.alignment = WD_TABLE_ALIGNMENT.CENTER
    grid.autofit = False
    grid.style = "Table Grid"
    for col, width in zip(grid.columns, widths, strict=True):
        col.width = Inches(width)
    for i, values in enumerate([headings, *rows]):
        row = grid.rows[0] if i == 0 else grid.add_row()
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        if i == 0:
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for cell, value, width in zip(row.cells, values, widths, strict=True):
            cell.width = Inches(width)
            shade = OxmlElement("w:shd")
            shade.set(qn("w:fill"), NAVY if i == 0 else SOFT if i % 2 else "FFFFFF")
            cell._tc.get_or_add_tcPr().append(shade)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(value)
            run.font.size = Pt(10.5)
            run.bold = i == 0
            if i == 0:
                run.font.color.rgb = RGBColor.from_string("FFFFFF")


def heading(doc, text):
    doc.add_heading(text, level=1)


def create(destination, *, include_worked=True):
    doc = Document()
    props = doc.core_properties
    props.title = "Bank of Ghana 2025 Financial Report Review"
    props.subject = "Datin Topic 4 audit-style classroom workpaper"
    props.author = "Tertiary Infotech Academy Pte Ltd"
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.top_margin = sec.bottom_margin = Inches(0.7)
    sec.left_margin = sec.right_margin = Inches(0.78)
    header = sec.header.paragraphs[0].add_run("TERTIARY INFOTECH ACADEMY  /  C34 DATIN  /  TOPIC 4")
    header.bold = True
    header.font.size = Pt(9)
    header.font.color.rgb = RGBColor.from_string(GOLD)
    footer = sec.footer.paragraphs[0]
    footer.add_run("Bank of Ghana 2025  ·  Financial-statement review  ·  ").font.size = Pt(9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    for name in ("Normal", "Title", "Heading 1", "Heading 2"):
        style = doc.styles[name]
        style.font.name = "Calibri"
        style.font.color.rgb = RGBColor.from_string(NAVY)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.13
    doc.styles["Title"].font.size = Pt(30)
    doc.styles["Heading 1"].font.size = Pt(18)
    doc.styles["Heading 2"].font.size = Pt(13)

    doc.add_heading("Bank of Ghana 2025\nFinancial Report Review", 0)
    doc.add_paragraph("Review workpaper for the year ended 31 December 2025.")
    doc.add_paragraph("Prepared by _________________________    Reviewed by _________________________")
    doc.add_paragraph("Review date _________________________    Source version ________________________")
    p = doc.add_paragraph()
    link(p, "Official Bank of Ghana report and financial statements", B.REPORT_URL)
    p.add_run("  •  ")
    link(p, "Open Microsoft 365 Copilot Chat", B.COPILOT_URL)
    doc.add_paragraph(
        "Use this audit-style walkthrough to review published disclosures. The work is a classroom "
        "analytical review and provides no independent assurance or audit opinion."
    )
    heading(doc, "1. Set the scope and extract")
    doc.add_paragraph(
        "Attach the official PDF in Copilot Chat. Use the original PDF for verification. Record the year, "
        "Bank or Group, signed amount, GH¢’000 units and printed page before using a figure. "
        "For the cited sections, the page position in the PDF viewer is one higher than the printed page."
    )
    for label, prompt in B.PROMPTS[:2]:
        doc.add_heading(label, level=2)
        doc.add_paragraph(prompt)
    doc.add_paragraph("If the upload is unavailable, work in short batches using the relevant pages or table excerpts.")

    doc.add_page_break()
    heading(doc, "2. Record your evidence")
    doc.add_paragraph("Use the blank cells to record your own results. Verify the calculations against the report.")
    table(doc, ("Check", "Printed page and entity", "Your result and evidence"), (1.6, 1.85, 3.24), [
        ("Bank year-on-year loss and OMO cost", "15 and Note 10 on 62", "[2025 • 2024 • change • source line]\n\n"),
        ("Total comprehensive result", "16 · 2025 Bank", "[Loss + OCI = total result]\n\n"),
        ("Balance-sheet equation", "17 · 2025 Bank", "[Assets − liabilities = equity]\n\n"),
        ("Cash bridge", "23 and Note 51(a) on 130 · Bank", "[Opening + movements + FX = closing]\n\n"),
        ("Group loss attribution", "5 and 15 · 2025 Group", "[Total • parent • NCI • explanatory note]\n\n"),
        ("Auditor’s scope", "10–12 and Note 2(a)(i) on 25", "[Opinion • emphasis • key audit matter]\n\n"),
    ])
    doc.add_heading("Evidence and action log", level=2)
    table(doc, ("Finding or discrepancy", "Evidence and status", "Owner and next step"), (2.2, 2.48, 2.01), [
        ("[Describe the issue]", "[Cite page, figure and status]", "[Record owner and follow-up]"),
        ("", "\n", ""),
    ])

    if include_worked:
        doc.add_page_break()
        heading(doc, "3. Worked review")
        doc.add_paragraph("Figures are in GH¢’000. Parentheses indicate negative amounts. Page citations refer to the supplied PDF.")
        table(doc, ("Check and source", "Reperformance and result"), (2.05, 4.64), [
            ("Bank loss and open market operations\nPrinted 15 and 62",
             f"Loss rose in magnitude from {B.amount(B.BANK_LOSS_24)} to {B.amount(B.BANK_LOSS_25)}, an increase of 64.7%. "
             f"OMO cost rose from {B.amount(B.OMO_24)} to {B.amount(B.OMO_25)}, up 94.6% in magnitude. Compare operating income and other cost lines before attributing the full loss movement."),
            ("Bank comprehensive loss\nPrinted 16",
             f"{B.amount(B.BANK_LOSS_25)} + {B.amount(B.BANK_OCI)} = {B.amount(B.COMPREHENSIVE)}. "
             "The comprehensive loss includes the reported OCI loss."),
            ("Bank negative equity\nPrinted 17 and 19",
             f"{B.amount(B.ASSETS)} − {B.amount(B.LIABILITIES)} = {B.amount(B.EQUITY)}. "
             f"Opening equity of {B.amount(B.BANK_OPEN_EQUITY)} plus comprehensive loss of {B.amount(B.COMPREHENSIVE)} produces the same closing equity."),
            ("Bank cash bridge\nPrinted 23 and 130",
             f"{B.amount(B.NET_OPERATING)} − {abs(B.NET_INVESTING):,} − {abs(B.NET_FINANCING):,} = {B.amount(B.NET_CASH_MOVE)} before exchange effects. "
             f"{B.amount(B.OPEN_CASH)} − {abs(B.NET_CASH_MOVE):,} + {B.amount(B.EXCHANGE_EFFECT)} = {B.amount(B.CLOSE_CASH)}."),
            ("Group loss attribution\nPrinted 5 and 15",
             f"The directors’ summary shows {B.amount(B.PARENT_LOSS)}. The statement shows total Group loss of {B.amount(B.GROUP_LOSS)}. "
             f"The amount attributable to equity shareholders plus {B.amount(B.NCI_LOSS)} attributable to non-controlling interest reconciles to the total. "
             "Use the statement’s labels and ask whether the summary label should specify the attribution."),
        ])
        doc.add_heading("Independent auditor’s report", level=2)
        doc.add_paragraph(
            "KPMG’s opinion relates to the accounting basis described in Note 2(a)(i). The emphasis of matter "
            "draws attention to that basis and does not modify the opinion. Investment securities impairment "
            "is a key audit matter. The report addresses it in the context of the overall audit. "
            "See printed pages 10–12 and Note 2(a)(i) on page 25."
        )

    doc.add_page_break()
    heading(doc, f"{4 if include_worked else 3}. Draft and review the Word memo")
    doc.add_heading(B.PROMPTS[2][0], level=2)
    doc.add_paragraph(B.WRITE)
    table(doc, ("Memo section", "Your reviewed draft"), (1.76, 4.93), [
        ("Purpose and scope", "[Entity, period, source, units and limits]\n\n"),
        ("Findings", "[Three verified findings with printed pages and amounts]\n\n\n"),
        ("Follow-up", "[One owner and next action for every unresolved point]\n\n"),
        ("Conclusion", "[What the reviewed evidence supports]\n\n"),
    ])
    doc.add_heading("Review before sharing", level=2)
    for item in (
        "Data control · Use the published PDF in the approved workspace.",
        "Permission control · Attach only the source needed for the task.",
        "Human review · Recalculate figures and approve the final wording.",
        "Traceability · Keep the prompt, cited pages and workpaper together.",
        "Guardrails · Flag missing evidence and keep the draft within the review scope.",
    ):
        doc.add_paragraph(item, style="List Bullet")
    doc.save(destination)


if __name__ == "__main__":
    import sys
    create(Path(sys.argv[1]), include_worked="--learner" not in sys.argv[2:])
