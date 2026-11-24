"""Interactive PowerPoint module for the Bank of Ghana financial review."""

import ccdeck as K
import bog_review as B
from datin_poll import write

STANDALONE = "Topic 4 - Bank of Ghana Financial Report Review.pptx"

CHECKS = (
    (
        "Auditor’s report", "Printed pp. 10–12, 25  ·  PDF pp. 11–13, 26", 11,
        ("Find the opinion and identify the accounting basis in Note 2(a)(i).",
         "Record the emphasis of matter and the investment securities impairment key audit matter.",
         "Write what the published opinion covers and keep your own review scope separate."),
        ("Opinion", "KPMG’s opinion refers to the basis of accounting in Note 2(a)(i)."),
        ("Emphasis of matter", "KPMG draws attention to that basis. The opinion is unmodified in respect of it."),
        ("Key audit matter", "Impairment on investment securities. Read it in the context of the overall audit."),
    ),
    (
        "Performance & attribution", "Printed pp. 5, 15–16, 62  ·  PDF pp. 6, 16–17, 63", 16,
        ("Extract the 2025 and 2024 Bank loss and cost of open market operations. Compare signed values and magnitudes.",
         "Recalculate 2025 Bank loss plus other comprehensive income.",
         "Compare the directors’ Group loss summary with total loss and the attribution lines on page 15."),
        ("Loss and OCI · Bank", f"{B.amount(B.BANK_LOSS_25)} loss + {B.amount(B.BANK_OCI)} OCI = {B.amount(B.COMPREHENSIVE)} total comprehensive loss."),
        ("Open market operations", "Bank OMO cost rose 94.6% in magnitude, from 8,595,549 to 16,730,526. The Bank’s loss was 64.7% larger."),
        ("Group attribution", f"{B.amount(B.PARENT_LOSS)} parent + {B.amount(B.NCI_LOSS)} non-controlling interest = {B.amount(B.GROUP_LOSS)} total. Follow up on the summary label."),
    ),
    (
        "Financial position", "Printed pp. 17, 19  ·  PDF pp. 18, 20", 18,
        ("Use only the 2025 Bank column. Extract total assets, liabilities and equity.",
         "Recalculate assets less liabilities. Compare the result with published equity and explain its sign.",
         "Cross-check closing equity against opening equity plus total comprehensive income."),
        ("Assets", B.amount(B.ASSETS)),
        ("Liabilities", B.amount(B.LIABILITIES)),
        ("Closing equity", f"{B.amount(B.EQUITY)}. Assets less liabilities equals opening equity {B.amount(B.BANK_OPEN_EQUITY)} plus comprehensive loss {B.amount(B.COMPREHENSIVE)}."),
    ),
    (
        "Cash flows", "Printed pp. 23, 130  ·  PDF pp. 24, 131", 24,
        ("Record the three Bank cash flow subtotals and recompute the net movement before exchange effects.",
         "Add the net movement and reported exchange effect to the opening cash balance.",
         "Tie your result to closing cash and distinguish operating inflow from the overall movement."),
        ("Operating cash", f"{B.amount(B.NET_OPERATING)} net inflow."),
        ("Net change before FX", f"{B.amount(B.NET_OPERATING)} − {abs(B.NET_INVESTING):,} − {abs(B.NET_FINANCING):,} = {B.amount(B.NET_CASH_MOVE)}."),
        ("Closing cash", f"{B.amount(B.OPEN_CASH)} − {abs(B.NET_CASH_MOVE):,} + {B.amount(B.EXCHANGE_EFFECT)} exchange effect = {B.amount(B.CLOSE_CASH)}."),
    ),
)


def button(slide, x, y, width, label, target, *, accent=False):
    """Use one shape for the label and the PowerPoint click action."""
    shape = K.card(slide, x, y, width, 0.49, fill=K.ACC if accent else K.WHITE, shadow=False, line=K.RULE)
    write(shape, label, 14, K.INK, bold=True, center=True)
    if isinstance(target, str):
        shape.click_action.hyperlink.address = target
    else:
        shape.click_action.target_slide = target
    return shape


def evidence_rows(slide, items, *, results=False):
    for i, (label, description) in enumerate(items):
        y = 2.38 + i * 1.15
        K.card(slide, 0.64, y, 12.05, 1.02, fill=K.WHITE, shadow=True)
        K.text(slide, 0.91, y + .10, 3.02, .79, label, 19, K.TEAL if results else K.INK,
               bold=True, font=K.SERIF, anchor="m", minpt=17)
        K.text(slide, 4.1, y + .08, 8.3, .85, description, 18, K.INK, anchor="m", minpt=16)


def add_topic(deck):
    """Add an entry, four selectable source routes and their worked checks."""
    entry = deck.new(dark=True)
    K.header(deck, entry, "Topic 4 · Copilot financial report review",
             "Review Bank of Ghana’s 2025 Financial Statements",
             "Open the published PDF in Copilot Chat. Verify the output and document your findings in Word.")
    for i, (title, detail) in enumerate([
        ("Attach the PDF", "Check 2025 versus 2024 and Bank versus Group."),
        ("Verify the evidence", "Recalculate from signed figures in GH¢’000 and cite printed pages."),
        ("Draft the Word memo", "Record findings, open items and review status."),
    ]):
        x = .66 + i * 4.07
        K.card(entry, x, 2.65, 3.91, 2.43, fill=K.PANEL, shadow=False)
        K.text(entry, x + .21, 2.91, 3.49, .5, title, 22, K.WHITE, bold=True, font=K.SERIF)
        K.text(entry, x + .21, 3.66, 3.49, 1.10, detail, 18, K.SOFT, minpt=17)

    menu = deck.new()
    K.header(deck, menu, "Topic 4 · evidence routes", "Choose a Statement or the Auditor’s Report",
             "Choose a route. Check the figures or wording in the PDF before opening the worked result.")
    button(entry, .65, 5.77, 2.30, "Start review", menu, accent=True)
    button(entry, 3.16, 5.77, 2.45, "Open report PDF", B.REPORT_URL)
    button(entry, 5.82, 5.77, 2.62, "Open Copilot Chat", B.COPILOT_URL)
    button(entry, 8.65, 5.77, 2.47, "Open lab prompts", B.PROMPTS_URL)
    entry.notes_slide.notes_text_frame.text = (
        f"Primary source: {B.REPORT_URL}\nMicrosoft attachment instructions: {B.ATTACH_URL}\n"
        "Form three groups of three. Use a public copy of the report in the approved Copilot workspace. "
        "Run the lab in Slide Show and use the buttons to move between routes. PDF page positions in this file are one higher than the printed page numbers cited."
    )

    for i, (title, pages, pdf_page, tasks, *results) in enumerate(CHECKS):
        task = deck.new()
        K.header(deck, task, "Topic 4 · verify with the source", title, pages)
        evidence_rows(task, tuple(zip(("Locate", "Reperform", "Record"), tasks, strict=True)))
        result = deck.new()
        K.header(deck, result, "Topic 4 · worked check", title,
                 f"{pages}   ·   Financial amounts in GH¢’000")
        evidence_rows(result, results, results=True)
        source = f"{B.REPORT_URL}#page={pdf_page}"
        button(task, .65, 6.24, 2.35, "All review routes", menu)
        button(task, 3.21, 6.24, 2.20, "Open source PDF", source)
        button(task, 5.62, 6.24, 2.70, "Show worked check", result, accent=True)
        button(result, .65, 6.24, 2.35, "All review routes", menu, accent=True)
        button(result, 3.21, 6.24, 2.20, "Open source PDF", source)
        button(result, 5.62, 6.24, 2.40, "Back to task", task)
        for s in (task, result):
            s.notes_slide.notes_text_frame.text = f"Bank of Ghana 2025 financial statements. {pages}. {source}"

        x, y = .65 + (i % 2) * 6.13, 2.43 + (i // 2) * 1.59
        shape = K.card(menu, x, y, 5.90, 1.31, fill=K.WHITE, line=K.RULE)
        write(shape, title, 23, K.INK, bold=True, center=True)
        shape.click_action.target_slide = task

    memo = deck.new(dark=True)
    K.header(deck, memo, "Topic 4 · lab handoff", "Turn the Checks into a Word Review Memo",
             "Open the Word workpaper from the lab page. Complete the evidence table before asking Copilot to draft the memo.")
    for i, (label, body) in enumerate([
        ("Purpose and scope", "Name the document, entity, period, units and review limits."),
        ("Three verified findings", "Cite the page and show the recalculation or source line."),
        ("Open items", "Assign a reviewer and a follow-up action for each unresolved point."),
    ]):
        y = 2.48 + i * 1.08
        K.card(memo, .65, y, 12.04, .96, fill=K.PANEL, shadow=False)
        K.text(memo, .9, y + .13, 3.55, .68, label, 20, K.WHITE, bold=True, anchor="m")
        K.text(memo, 4.53, y + .13, 7.87, .68, body, 18, K.SOFT, anchor="m")
    button(menu, .65, 5.99, 2.76, "Draft the Word memo", memo, accent=True)
    button(menu, 3.62, 5.99, 2.20, "Open source PDF", B.REPORT_URL)
    button(memo, .65, 6.24, 2.35, "All review routes", menu)
    button(memo, 3.21, 6.24, 2.62, "Open lab prompts", B.PROMPTS_URL)
    memo.notes_slide.notes_text_frame.text = (
        f"Draft prompt for Word:\n{B.WRITE}\nAsk learners to record an owner and next step for any unresolved point."
    )
    return entry, memo


def security_slide(deck, *, return_to=None):
    s = deck.new(dark=True)
    K.header(deck, s, "Class wrap-up · AI security", "Five Controls for AI in Finance",
             "Apply these controls before attaching evidence, while reviewing the analysis and before sharing a result.")
    for i, (label, description) in enumerate(B.SECURITY):
        y = 2.35 + .75 * i
        K.card(s, .66, y, 12.01, .63, fill=K.PANEL, shadow=False)
        K.text(s, .90, y + .05, 3.05, .51, label, 18, "F2A383", bold=True, anchor="m")
        K.text(s, 4.10, y + .05, 8.25, .51, description, 16, K.WHITE, anchor="m", minpt=15)
    if return_to is not None:
        button(s, .65, 6.32, 2.62, "Back to Topic 4", return_to)
    return s
