"""One fictional Tema finance case for the canvas, worked demo and class handoffs."""

from pptx.util import Pt

CASE_NOTE = "Fictional Ghana distributor. Volumes, timing and budget are illustrative."
WHY = "Use the case to decide who benefits, what data is needed and where finance approval stays."
DEMO_NOTE = "Worked answers for the Tema distributor. Use all nine blocks to test the pilot decision."

# Keep these short for the native nine-block canvas.
CANVAS = (
    "Suppliers, hauliers and the ERP provider furnish matching records.",
    "Rank close exceptions. Match orders, invoices and delivery records.",
    "Sales ledger, delivery notes, returns and approval rules.",
    "Faster, traceable corrections to short deliveries, invoices and credits.",
    "Finance reviews disputes and approves credits and payments.",
    "Customer email and WhatsApp. Internal finance queue.",
    "Retailers and wholesalers in Greater Accra, Ashanti and Western.",
    "Illustrative year-one budget: GHS 39,000 for licences, integration and training. Measure time saved and review effort.",
    "GHS sales of staples, household and personal care products. Faster dispute resolution supports collections.",
)

TOPIC_1_HANDOFF = "At 2:20 pm, use Tema’s H1 sales data to explain revenue and margin by region and product."
TOPIC_2_HANDOFF = "At 4:00 pm, turn H1 dashboard instructions into a skill and test them on Tema’s fresh H2 orders."
COURSE_MAP = {
    "Topic 1 · Agentic AI": "Topic 1 · Triage the close",
    "Identify agent roles, tool access and approval gates. Use Muse for the close brief.": "Map the use case, then use Muse to match the close inbox and calendar. Review the evidence before approval.",
    "Topic 2 · Dashboard": "Topic 2 · Explain performance",
    "Build and restyle an Excel dashboard to explain revenue by region and product.": "Use H1 sales in Excel to reconcile Ghana cedi revenue and margin across regions, channels and products.",
    "Topic 3 · Skills": "Topic 3 · Repeat the report",
    "Save and enable a Copilot skill. Reuse it to build on fresh H2 data.": "Turn the H1 dashboard instructions into a Copilot skill. Test it on fresh H2 orders against the same controls.",
}

GENERAL = {
    4: {"Topic 1 · Meridian worked example": "Topic 1 · Tema canvas demo"},
    7: {
        "Use cases by function; Meridian and the nine-block BMC": "Finance uses and the nine-block canvas",
    },
    59: {
        "Flag a possible supplier duplicate and an unsupported customer credit.": "Check a possible duplicate invoice and a credit-note request against approved records.",
    },
    64: {"Topic 2 starts at 2:20 pm.": TOPIC_1_HANDOFF},
    66: {
        "Show how net revenue moves over time, by region and by product.": "Switch from triage to reporting. Compare Tema’s H1 revenue by region and product.",
        "Check the mix and COGS before drawing conclusions about profitability.": "Reconcile gross profit and margin before explaining differences in sales mix.",
    },
    78: {"Topic 3 starts at 4:00 pm.": TOPIC_2_HANDOFF},
    86: {
        "What You Can Now Verify": "From Close Exceptions to Repeatable Reporting",
        "Trace a close brief to the approved demo inbox and calendar.": "Trace Tema’s close priorities to the approved inbox and calendar. Keep approvals with finance.",
        "Reconcile net revenue and margin across region, product and channel.": "Reconcile H1 revenue and margin, then explain the pattern by region, product or channel.",
        "Enable datin-dark-sales-dashboard and check results on a fresh workbook.": "Enable the dashboard skill, build on H2 orders and verify the figures again.",
    },
}

SOURCE_TEXT = {
    27: {
        "The big jump is copilot → agent: from prompt-and-response to delegating a goal.":
            "Preset rules handle known cases and route other exceptions to a person. RPA is one way to run those steps.",
        "You ask, it answers. It cannot act.": "Explains the invoice-matching policy.",
        "Task bot": "Automated Workflow",
        "One clear job, precise instructions.": "Matches invoices by set rules and routes exceptions.",
        "Copilot": "Human-assisted copilot",
        "Suggests; you steer and approve every step.": "You steer. It helps investigate unusual cases.",
        "Goal in, result out, while you sleep.": "Plans and checks a delegated goal within limits.",
        "more autonomy, more controls": "More discretion, clearer controls",
    },
    57: {
        "TOPIC 1 · IDENTIFY AND EVALUATE · MERIDIAN SCENARIO": "TOPIC 1 · WEST AFRICA FINANCE USE CASE",
        "Meridian Today: 80 Requests a Day, Every One Re-Keyed by Hand": "Tema Distributor: 80 Documents a Day",
        "A worked example of how to scope a repeatable task and quantify an opportunity.": "A fictional Ghana wholesaler wants to shorten its month-end review and catch costly exceptions.",
        "Meridian Services Pte Ltd": "Tema Wholesale Distributor",
        "140 staff keep 30 commercial buildings running across Singapore.": "Sells staples, household and personal care goods across Ghana.",
        "The pain": "Scattered evidence",
        "80 emails and WhatsApps a day. Coordinators spend 38% of their time re-keying them into tickets.": "80 invoices, delivery notes and returns messages arrive daily.",
        "The cost": "Manual exception review",
        "18 minutes per request; an urgent leak at 2 am waits until morning.": "Each exception takes 18 minutes to check and route at close.",
        "The new business": "Proposed finance pilot",
        "Meridian Assist: an always-on agent it runs in-house first, then sells.": "Prioritise cut-off, possible duplicates and credit-note requests.",
        "MERIDIAN'S TOPIC 1 QUESTION": "THE FINANCE CONTROLLER ASKS",
        "Where on our business model would AI agents pay off first, and which idea do we park?": "Can an agent speed up the close while finance keeps control?",
        "Meridian is fictitious and all figures are synthetic.": CASE_NOTE,
    },
    58: {
        "The BMC shows how a company creates, delivers and captures value. Read the customer side, then delivery and money.": "Start with the nine questions. Then use the Tema distributor case to fill in the canvas.",
    },
    65: {
        "Example:  triage agent 5 × 4 × 4 = 80  ·  report agent 75  ·  fault predictor 24 → park it":
            "Tema example: close triage 5 × 4 × 4 = 80  ·  sales dashboard 75  ·  cash forecast 24",
    },
    72: {
        "TOPIC 1 · MERIDIAN DEMO · HOW TO FILL EACH BMC BLOCK": "TOPIC 1 · SAMPLE DEMO · CUSTOMER SIDE",
        "Meridian Fills the Customer Side First": "Who Benefits From a Faster Close?",
        "Watch how Meridian filled the canvas. Use the finance example to relate these choices to accounting.": WHY,
    },
    73: {
        "TOPIC 1 · MERIDIAN DEMO · HOW TO FILL EACH BMC BLOCK": "TOPIC 1 · SAMPLE DEMO · OPERATIONS AND COST",
        "Then the Engine Room: What Agents Do, Need and Cost": "What Will the Pilot Need and Cost?",
        "Watch how Meridian filled the canvas. Use the finance example to relate these choices to accounting.": WHY,
    },
    75: {
        "TOPIC 1 · MERIDIAN'S COMPLETED BMC": "TOPIC 1 · SAMPLE DEMO · COMPLETED CANVAS",
        "Meridian's Canvas, Block by Block": "Sample Answers: Tema Distributor’s Canvas",
        "The first note of each block, readable. The live board holds three per block.": DEMO_NOTE,
        **dict(zip(
            (
                "ServiceNow (ticketing vendor): Holds every service request since 2019.",
                "Request triage (agent): Agent reads 80 emails and WhatsApps a day.",
                "5 years of closed tickets: Root causes and repeat faults per site.",
                "Same-day answers, 24/7: Always-on triage agent replies within minutes.",
                "Agent first, human always reachable: Routine updates by agent.",
                "Email + WhatsApp Business: Agent drafts every reply; coordinator sends at first.",
                "Building managers (clients): Want same-day answers.",
                "Year-1 cost SGD 39,000: Licences 14k, integration 15k, training and change 10k.",
                "Premium '2-hour response' tier: Agent makes the SLA affordable.",
            ), CANVAS,
        )),
    },
}

TABLES = {
    59: (
        ("BMC block / meaning", "Pilot use at Tema", "Benefit / condition"),
        (None, "Retailers and wholesalers want accurate invoices and credits.", "Prioritise recurring disputes by account and deadline."),
        (None, "The agent matches short-delivery and return evidence to invoices.", "Faster, traceable corrections after finance review."),
        (None, "Use approved inbox messages and calendar reminders for the internal queue.", "Staff approve customer communications during the pilot."),
    ),
    60: (
        ("BMC block / meaning", "Pilot use at Tema", "Benefit / condition"),
        (None, "Flag disputes and draft a brief for a finance reviewer.", "A named accountant handles customer follow-up and approvals."),
        (None, "Reduce unresolved discrepancies on GHS wholesale and retail sales.", "Track collections separately from recognised revenue."),
        (None, "Rank close exceptions. Check delivery, duplicates and credits.", "Measure time per case and the accuracy of the exceptions raised."),
    ),
    61: (
        ("BMC block / meaning", "Pilot use at Tema", "Benefit / condition"),
        (None, "Use invoice, sales ledger, delivery and return records plus close rules.", "Use only approved records and link each recommendation to evidence."),
        (None, "Suppliers and hauliers supply evidence. The ERP provider supports controlled access.", "Document record owners and who can authorise access."),
        (None, "Budget GHS 39,000 in year one for tools, integration and training.", "Compare time saved with full costs, including human review."),
    ),
    72: (
        ("BMC block", "Decision to make", "Tema sample answer"),
        (None, "What improvement will the customer see?", "Faster resolution of short deliveries and disputed charges."),
        (None, "Who should see the benefit first?", "Retailers and wholesalers with repeat orders and invoice queries."),
        (None, "Which decisions remain with people?", "The finance manager approves credit notes, payments and postings."),
        (None, "Where does work arrive?", "Email and WhatsApp Business. Use the demo inbox and calendar in class."),
    ),
    73: (
        ("BMC block", "Decision to make", "Tema sample answer"),
        (None, "What can the agent do during the pilot?", "Rank exceptions and check references. Produce a draft brief."),
        (None, "What evidence does the agent need?", "Invoice, purchase order, proof of delivery, returns and an approval matrix."),
        (None, "Who supplies or owns that evidence?", "Suppliers, hauliers and the ERP provider. Finance controls access."),
        (None, "What costs and benefits will be measured?", "Illustrative GHS 39,000 year-one budget. Track review time and accuracy."),
        (None, "How does this support the business?", "Resolve invoice disputes faster to support collections."),
    ),
}


def replace_text(target, new):
    """Preserve theme and run styling when replacing the existing sentence."""
    runs = [r for p in target.text_frame.paragraphs for r in p.runs]
    if not runs:
        target.text = new
        return
    if runs[0].text == "WHY  " and new.startswith("WHY  ") and len(runs) > 1:
        runs = runs[1:]
        new = new[5:]
    runs[0].text = new
    for r in runs[1:]:
        r.text = ""


def patch_text(slide, replacements):
    remaining = dict(replacements)
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text in remaining:
            replace_text(sh, remaining.pop(sh.text))
    if remaining:
        raise ValueError(f"Missing case copy on {slide.part.partname}: {list(remaining)}")


def patch_reference(slide, number):
    patch_text(slide, SOURCE_TEXT.get(number, {}))
    if number in (59, 60, 61):
        patch_text(slide, {
            "TOPIC 1 · MERIDIAN ACROSS THE BMC": "TOPIC 1 · WORKED CANVAS · TEMA DISTRIBUTOR",
            "Meridian is fictitious. These are proposed uses and benefits to test, not proven results.": CASE_NOTE,
        })
    if number in TABLES:
        table = next(sh.table for sh in slide.shapes if sh.has_table)
        data = TABLES[number]
        assert len(table.rows) == len(data)
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                if value is not None:
                    replace_text(table.cell(i, j), value)
    if number == 75:
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text in CANVAS:
                for paragraph in sh.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(11.5)


def decision_slide(deck, slide, kit):
    """Reuse the established reference card typography and footer."""
    kit.header(deck, slide, "Topic 1 · choose the first pilot", "Three Options for the Tema Finance Team",
               "Start where the data is ready, the work repeats and a person can verify the outcome.")
    for i, (label, action, evidence, color) in enumerate((
        ("80 · Pilot close triage", "Lab 1 · Meta Muse", "Track inbox and calendar deadlines. Verify each exception before finance approval.", kit.CHAR),
        ("75 · Build the sales view", "Lab 2 · Excel", "Use reconciled H1 orders to explain revenue and gross margin across markets and products.", kit.WHITE),
        ("24 · Defer cash forecasting", "Review once data is ready", "Confirm payment dates and collection history before forecasting cash.", kit.WHITE),
    )):
        x = 0.6 + i * 4.1
        dark = color == kit.CHAR
        fg = kit.WHITE if dark else kit.INK
        kit.card(slide, x, 2.53, 3.9, 3.8, fill=color)
        kit.text(slide, x + 0.23, 2.8, 3.38, 0.65, label, 20, fg, bold=True, font=kit.SERIF, minpt=16)
        kit.text(slide, x + 0.23, 3.57, 3.38, 0.42, action, 16, "F2B49B" if dark else kit.ACC_D,
                 bold=True, minpt=14)
        kit.text(slide, x + 0.23, 4.24, 3.38, 1.48, evidence, 17, kit.SOFT if dark else kit.INK, minpt=15)
