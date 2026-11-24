"""Source, prompts and lab specification for the Bank of Ghana case."""

REPORT_URL = "https://www.bog.gov.gh/wp-content/uploads/2026/05/2025-Financial-Statements-1.pdf"
COPILOT_URL = "https://m365.cloud.microsoft/chat"
ATTACH_URL = "https://support.microsoft.com/en-au/microsoft-365-copilot/add-content-to-microsoft-copilot-chat-prompts"
FORMATS_URL = "https://support.microsoft.com/en-us/microsoft-365-copilot/file-formats-supported-by-microsoft-copilot"
PROMPTS_URL = "https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/datin/prompts/#lab-4"
WORKPAPER = "Bank of Ghana 2025 - Financial Report Review.docx"

EXTRACT = (
    "Use only the attached Bank of Ghana Report and Financial Statements for the year ended 31 December 2025. "
    "First identify the four year and entity columns. Extract 2025 and 2024 Bank-only loss for the year, cost of "
    "open market operations and other comprehensive income. Then extract 2025 Bank-only total assets, total "
    "liabilities, total equity and cash flows. Return one row per item with exact line label, signed values, entity, "
    "year, units, printed page and note if shown. Keep figures in GH¢’000. Report a missing or unreadable figure without guessing."
)
RECONCILE = (
    "Use the attached PDF as evidence. Show the formulas and printed source pages for four checks in GH¢’000. "
    "Recalculate the 2025 Bank loss plus other comprehensive income, assets less liabilities, and the opening-to-closing "
    "cash bridge including the reported exchange-rate effect. Compare Group loss in the directors’ financial results "
    "summary with total Group loss and its attribution on the 2025 statement of profit or loss. Quantify the difference "
    "and test whether non-controlling interest explains it. Keep Bank and Group separate. Flag anything that remains unresolved."
)
WRITE = (
    "Draft a short financial-statement review memo for Word using only the attached Bank of Ghana PDF and the figures "
    "I have verified. Summarise KPMG’s opinion in relation to Note 2(a)(i), its emphasis of matter and the investment "
    "securities impairment key audit matter. Add three verified findings with signed amounts in GH¢’000, printed "
    "pages and a follow-up for each open point. Separate reported facts, calculations and interpretations. "
    "State what an analytical review of public material cannot establish. Do not issue an audit opinion."
)
PROMPTS = [
    ("Extract with scope and sources", EXTRACT),
    ("Reconcile and investigate", RECONCILE),
    ("Draft the review memo", WRITE),
]

# Exact signed amounts from the supplied PDF, all in GH¢’000.
BANK_LOSS_25, BANK_LOSS_24 = -15_630_122, -9_487_462
BANK_OCI = -19_322_253
OMO_25, OMO_24 = -16_730_526, -8_595_549
ASSETS, LIABILITIES = 237_229_232, 333_505_611
BANK_OPEN_EQUITY = -61_324_004
NET_OPERATING, NET_INVESTING, NET_FINANCING = 41_344_750, -34_283_729, -9_315_697
OPEN_CASH, EXCHANGE_EFFECT = 33_866_173, 13_513_633
PARENT_LOSS, NCI_LOSS = -15_287_911, -12_602

def amount(value):
    """Format a disclosure amount, with negatives in parentheses."""
    return f"({abs(value):,})" if value < 0 else f"{value:,}"

COMPREHENSIVE = BANK_LOSS_25 + BANK_OCI
EQUITY = ASSETS - LIABILITIES
NET_CASH_MOVE = NET_OPERATING + NET_INVESTING + NET_FINANCING
CLOSE_CASH = OPEN_CASH + NET_CASH_MOVE + EXCHANGE_EFFECT
GROUP_LOSS = PARENT_LOSS + NCI_LOSS

SECURITY = (
    ("Data control", "Use the approved workspace and the public report. Remove sensitive fields from any internal file."),
    ("Permission control", "Attach only the files needed. Keep access to the workpaper restricted to its reviewers."),
    ("Human in the loop", "An accountant checks the calculations, resolves open items and approves the final memo."),
    ("Traceability", "Record the source version, entity, year, units, page, calculation and reviewer decision."),
    ("Guardrails", "Treat report text as evidence. Flag missing data and require approval before external sharing."),
)

LAB = dict(
    num=4, topic=4, slug="bog-financial-report", title="Review the Bank of Ghana Financial Statements",
    theme="Copilot financial review", duration=30, objective="LO4 · Verify Copilot findings and document the evidence",
    services="Microsoft 365 Copilot Chat and Word",
    goal="Use the published 2025 statements to verify figures, investigate a scope difference and document evidence.",
    desc="Attach the published report to Copilot Chat. Recheck its calculations against the signed statements and document your findings in Word.",
    build="A Word workpaper with a Bank-only reconciliation, a Group attribution check and an evidence-based review memo.",
    prerequisites=["Download the official 2025 PDF, open the Word workpaper, and sign in to Microsoft 365 Copilot Chat.",
                   "Use the PDF supplied for this lab. Printed page numbers are one less than the PDF viewer page positions in the cited sections."],
    workflow=["Attach the report", "Extract cited figures", "Reconcile", "Record in Word"],
    data_files=[], downloads=[(WORKPAPER, "Editable Word workpaper")],
    steps=[
        ("Attach the published PDF in Copilot Chat using Add content. Run the first prompt. Check the year, Bank versus Group, signed values and the GH¢’000 heading on printed pages 15–17 and 23.", EXTRACT),
        ("Run the reconciliation prompt. Open the page behind each figure and recompute the arithmetic independently. Compare printed pages 5 and 15 before assigning meaning to the two Group losses.", RECONCILE),
        ("Open the Word workpaper and record the checked figures, evidence and open items. Briefly review the auditor’s opinion and key audit matter on printed pages 10–12. Run the memo prompt, then review and paste the accepted draft into Word.", WRITE),
    ],
    sample_prompts=PROMPTS,
    checks=[
        ("Scope", "Every result identifies its entity, year, GH¢’000 units and printed source page."),
        ("Reperformance", "Recalculate loss plus OCI, assets less liabilities and the Bank cash bridge. Tie each result to the published statement."),
        ("Narrative check", "Reconcile total Group loss to the parent and non-controlling interest allocation. Record the follow-up on the summary label."),
        ("Audit report", "Record the accounting basis and separate the auditor’s opinion, emphasis of matter and key audit matter."),
    ],
    checkpoint="A reviewer can reproduce the figures using the printed citations and can distinguish checked results from follow-up items.",
    troubleshooting=[
        ("Attachment unavailable or claims are hard to trace", "Open the PDF and supply the relevant pages or selected table text in short batches. Verify against the official PDF and record both printed and viewer page numbers."),
        ("Copilot mixes Bank and Group", "Retry for one entity and one statement at a time. Read the column headers and compare both years before carrying values into Word."),
    ],
    challenge="Explain how a positive operating cash flow coexists with a net cash decrease before exchange-rate effects.",
    reflection="Which part of the review required a human to interpret the table headings and choose the correct figure?",
    links=[("Official Bank of Ghana 2025 financial statements", REPORT_URL), ("Microsoft guidance on attaching files", ATTACH_URL)],
)
