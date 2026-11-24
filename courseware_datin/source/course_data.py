"""Datin edition content shared by the C34 deck, guide and lesson plan."""

import bog_review as B
from datasets import FIRST, FRESH, headline, records, totals
from labs import EXCEL_URL, SKILLS_URL, SLICERS_URL
from tema_case import CANVAS

TITLE = "Agentic Ai for Financial Data Analysis"
SHORT_TITLE = TITLE
COURSE_CODE = "C34"
COURSE_PAGE = "Datin edition · Open README.md in the courseware_datin folder"
VERSION = "v1.6"
VERSION_DATE = "24 November 2026"
ORG = "Tertiary Infotech Academy Pte Ltd"
UEN = "UEN: 201200696W"
TRAINER = "Assigned Tertiary Infotech Academy Trainer"
DAYS = 1
DAY_MINUTES = 360
INSTRUCTIONAL_MINUTES = 330
DAILY_TIMING = "11:00 am–6:00 pm, including one-hour lunch and two 15-minute afternoon tea breaks"
MODE = "Instructor-led demonstration and hands-on practice"
PARTICIPANTS = "9 accounting professionals from West Africa"
VERSION_HISTORY = [
    ("1.6", VERSION_DATE, "Added the Bank of Ghana financial report review, interactive PowerPoint and AI security wrap-up.", TRAINER),
    ("1.5", VERSION_DATE, "Removed the three interactive knowledge checks and their answer slides.", TRAINER),
    ("1.4", VERSION_DATE, "Added sourced leadership quotes and portraits, and updated the live poll button.", TRAINER),
    ("1.3", VERSION_DATE, "Linked the opening poll to its live results and updated the AI agents quotation.", TRAINER),
    ("1.2", VERSION_DATE, "Connected the Tema finance case, worked canvas, Meta Muse brief and Excel reporting labs.", TRAINER),
    ("1.1", VERSION_DATE, "Reference PowerPoint design and Topic 1 slides, Ghana case imagery, and 11 am–6 pm schedule.", TRAINER),
    ("1.0", VERSION_DATE, "First Datin edition based on C34 v3.2.", TRAINER),
]

LO_TITLES = ["Agentic AI for business", "Interactive sales dashboards", "Reusable Excel skills", "Copilot financial report review"]
LEARNING_OUTCOMES = [
    "LO1: Describe the development of AI, distinguish an AI agent from agentic AI, and explain where tools, skills and approvals fit.",
    "LO2: Build an interactive sales dashboard in Excel, reconcile its GHS metrics and change its theme while preserving interactive behavior.",
    "LO3: Turn a styling prompt into an Excel skill, enable it through OneDrive and reuse it on a fresh workbook with a short request.",
    "LO4: Review a published financial report with Copilot, verify the figures and record the evidence and review limits in Word.",
]
TOPIC_RECAPS = {
    1: [("LO1", "Use the Meta Muse brief to explain what the agent inferred, what needs a tool and what needs human approval.")],
    2: [("LO2", "Demonstrate the same GHS totals and working filters before and after the dark theme.")],
    3: [("LO3", "Show the enabled skill and reconcile a dashboard created from fresh data.")],
    4: [("LO4", "Show the recalculations, printed source references and reviewed Word memo.")],
}

AGENTS = "https://www.anthropic.com/research/building-effective-agents"
HISTORY = ["https://proceedings.neurips.cc/paper_files/paper/2012/hash/c399862d3b9d6b76c8436e924a68c45b-Abstract.html",
           "https://arxiv.org/abs/2005.14165"]


def section(title, definition, why, how, example, use_when, avoid_when, sources, *, brief, table=None):
    return dict(title=title, definition=definition, why=why, how=how, example=example,
                use_when=use_when, avoid_when=avoid_when, sources=sources, brief=brief, compare=table)


def lesson(left, right, *, left_head="WORKED EXAMPLE", right_head="WHAT TO CHECK", flow=None):
    return dict(single=True, prompts=left, verify=right, try_head=left_head, check_head=right_head,
                kicker="CONCEPT IN PRACTICE", flow=flow)


TOPICS = [
    dict(
        num=1, code="01", title="Business Innovation with Agentic AI and AI Agents",
        subtitle="A brief history of AI, business use cases, tools and skills, and one Meta Muse activity.",
        concepts=[
            ("AI has several eras", "Rules, machine learning, deep learning and large language models coexist in modern systems."),
            ("Agentic AI describes behaviour", "A system pursues a goal by planning, acting and checking its results."),
            ("An agent is a deployed worker", "The agent carries out a defined job through the permissions and tools it has."),
            ("Tools and skills", "A tool performs an operation. A skill packages a repeatable way to work and verify the outcome."),
        ],
        sections=[
            section(
                "From Rules to AI Agents",
                "Symbolic systems encoded rules and knowledge. Statistical machine learning learned patterns from examples. Deep learning scaled representation learning, and large language models made fluent language interfaces practical.",
                "These approaches build on and coexist with each other. Businesses still use rules to enforce policy, predictive models to score risk and databases to ground answers. An agent adds an execution layer that can call those existing systems.",
                ["Start with explicit policy rules.", "Use a learned model where a pattern can be inferred from examples.",
                 "Use a language model to interpret instructions and explain results.",
                 "Add tools, a feedback loop and runtime controls to pursue a delegated goal."],
                ["An overdue-invoice threshold starts as a rule.", "A model estimates late payment risk. A language model explains why an invoice was flagged.",
                 "An agent retrieves the invoice and payment records and prepares a draft follow-up."],
                ["Features and next actions depend on context."], ["A stable, fully specified rule already solves the decision."],
                HISTORY,
                table=dict(title="A Short History of AI", headers=["Era", "Capability", "Business example"], rows=[
                    ("1950s–1980s", "Symbolic AI and expert systems", "Apply if–then policy rules"),
                    ("1990s–2000s", "Wider use of statistical ML", "Classify or score records"),
                    ("2010s", "Deep learning at scale", "Recognise images and speech"),
                    ("2020s", "LLMs, then tool-using agents", "Draft, retrieve, act and verify"),
                ]),
                brief=lesson(["Inventory rule → demand forecast → language explanation → agent-generated reorder draft."],
                             ["These dates illustrate wider adoption. The techniques overlap.",
                              "A manager approves a purchase before the agent can commit it."]),
            ),
            section(
                "Agentic AI and AI Agents",
                "Agentic AI is a broad term for goal-directed behaviour involving planning, action and feedback. An AI agent is a concrete worker configured for a particular goal, with a model, tools, context, state and controls.",
                "A single agent can be invoked once or triggered on a schedule. A workflow can coordinate one or more agents with conventional rules and human approvals. More autonomy increases the need for clear permissions, stop rules and evaluation.",
                ["Receive a goal and boundaries.", "Plan a permitted action using approved context.", "Call a tool and inspect its result.",
                 "Repeat within limits, stop once complete, or escalate."],
                ["A collections agent reads a synthetic receivables export and drafts reminders for overdue accounts.",
                 "A coordinator verifies any disputed balance and approves external sends.",
                 "A workflow is a sequence of steps, such as matching an invoice, flagging a mismatch and routing it for review.",
                 "The automated workflow matches invoices using preset rules. It routes unusual exceptions to finance when no rule resolves them. The accountant can then use a copilot to inspect the evidence and steer the outcome. Finance keeps approval.",
                 "Robotic process automation (RPA) is one way to automate those steps, often through existing application interfaces. The workflow is the sequence itself. It can also be run by an integration or a person.",
                 "A goal-directed agent can decide which approved records to inspect next, within the limits the finance team sets."],
                ["The next step changes with data or exceptions."], ["The task grants broad permissions with no business owner."],
                [AGENTS],
                brief=lesson(["Agentic AI describes the plan–act–check behaviour.",
                              "A collections agent is one implementation with a specific job and access."],
                             ["An on-demand agent runs for one task. An always-on agent can use approved event or schedule triggers."],
                             left_head="BEHAVIOUR AND IMPLEMENTATION",
                             flow=dict(title="A Bounded Agent Loop", steps=["Receive goal", "Plan", "Use tool", "Check", "Stop or escalate"])),
            ),
            section(
                "Where Agents Add Business Value",
                "Finance use cases include cut-off review, duplicate invoice checks, receivables follow-up, regional revenue reporting and cash reconciliation.",
                "Start with recurring work, known source data and a human owner. Measure time saved and errors alongside the quality of the result. A good pilot limits the agent to a draft or well-scoped action and gives it a clear stop condition.",
                ["Choose a recurring, time-consuming workflow.", "Define the baseline and how to check the result.",
                 "Grant the least access required.", "Give a person responsibility for exceptions and approvals."],
                ["For weekly sales reporting, compare a dashboard’s totals with the source and measure analyst review time.",
                 "For invoice matching, flag amount or vendor discrepancies to a reviewer."],
                ["Reliable inputs and a repeatable outcome exist."], ["Source data or the accountable reviewer is undefined."],
                [AGENTS],
                table=dict(title="Three Business Use Cases", headers=["Use case", "Agent action", "Human gate"], rows=[
                    ("Revenue cut-off", "Match invoice date to delivery evidence", "Approve close adjustment"),
                    ("Payables", "Flag suspected duplicate supplier invoices", "Review exception before payment"),
                    ("Management reporting", "Reconcile regional and product revenue", "Approve management pack"),
                ]),
                brief=lesson(["Use the sales dashboard as a weekly reporting pilot. Compare minutes per report before and after."],
                             ["Count reconciliation errors and exceptions. Agree who approves the final report."]),
            ),
            section(
                "The Tema Distributor Business Model Canvas",
                "The Business Model Canvas uses nine blocks to show who a business serves, what value it provides, how it operates and how it earns revenue. Read the framework, then use the fictional Tema wholesale distributor case to work through the completed sample.",
                "The distributor receives about 80 sales invoices, delivery notes and returns messages a day. A finance exception takes an illustrative 18 minutes to check and route at close. The pilot tests whether an agent can prioritise exceptions and assemble evidence while accountants retain approval.",
                ["Start with the customer and the desired improvement.",
                 "Map permitted finance records, partners, channels and approval owners.",
                 "Estimate full costs and decide what to measure before the pilot."],
                [f"{label} — {answer}" for label, answer in zip(
                    ("Key Partners", "Key Activities", "Key Resources", "Value Propositions",
                     "Customer Relationships", "Channels", "Customer Segments", "Cost Structure",
                     "Revenue Streams"), CANVAS, strict=True)],
                ["You need to decide which finance task to pilot and how it supports the business."],
                ["The agent cannot trace a recommendation to approved source records."],
                [AGENTS],
                brief=lesson(["Start with close triage."], ["Measure review effort, accuracy and escalation handling."]),
            ),
            section(
                "Always-On Agents in Finance Teams",
                "An always-on agent follows approved triggers between requests. The reference deck profiles Meta Muse, OpenAI Dots, Claude Duet, Grok Bot, Manus Cue and Gemini Spark or Task. The class uses Meta Muse for inbox and calendar tracking. Product access is account- and region-dependent.",
                "The five collaboration patterns describe how people and agents share work. Agreement on a common ticket ID, evidence source and approval owner makes the close brief traceable across hand-offs.",
                ["One accountant and one agent track close deadlines.",
                 "Several payables staff share one invoice-exception agent.",
                 "One controller delegates extraction, matching and review to several agents.",
                 "Multiple finance teams coordinate payables, receivables and general ledger agents.",
                 "A buyer’s payables agent coordinates with a supplier’s receivables agent, with approval at both organisations."],
                ["For the Tema case, Muse matches demo email and calendar records by ticket ID. The finance team checks cut-off, duplicate payment and credit-note evidence."],
                ["A recurring workflow has controlled inputs and a designated reviewer."],
                ["Use of the connected inbox or calendar is outside the approved scope."],
                [AGENTS, "https://muse.ai/"],
                brief=lesson(["Match evidence by ticket ID."], ["Review the proposed action and approval owner."]),
            ),
            section(
                "Tools, Skills and Context",
                "A tool is a callable capability, such as reading a spreadsheet, querying a ledger or creating a chart. APIs and MCP can expose tool capabilities. A skill is a reusable set of instructions, examples and checks for when and how to perform work.",
                "A skill tells an agent how to sequence operations and what success looks like. It uses the access the environment already provides. Memory carries approved state between tasks. Knowledge supplies records and documents to consult. Instructions give the current goal, and runtime controls govern triggers, permissions and limits.",
                ["Identify what the agent must do and which tools permit it.",
                 "Write the sequence and quality checks for a repeatable task as a skill.",
                 "Ground the work in the current data and confirm permissions before any write."],
                ["Excel’s chart operation can plot a linked series.",
                 "A Datin dashboard skill says which KPIs, chart types, colours and filter tests to use."],
                ["A work standard should be reused across workbooks."], ["The user has no permission to access the required data."],
                [AGENTS, SKILLS_URL],
                table=dict(title="A Tool Acts and a Skill Guides", headers=["Element", "Purpose", "Datin example"], rows=[
                    ("Tool", "Read, calculate or edit through an allowed capability", "Insert an Excel chart"),
                    ("Skill", "Store a procedure and its verification checks", "Build and style a sales dashboard"),
                    ("Knowledge", "Supply facts or records", "Read tblDatinSales"),
                    ("Runtime", "Enforce access and stop conditions", "Require approval before sharing"),
                ]),
                brief=lesson(["The model uses the user goal and workbook context.",
                              "The skill tells it which Excel tools to use and how to check the result."],
                             ["Access comes from permissions. Review the result before sharing."],
                             left_head="HOW THE COMPONENTS FIT"),
            ),
        ],
    ),
    dict(
        num=2, code="02", title="Create an Interactive Dashboard in Excel",
        subtitle="Build with Copilot, reconcile sales and profit in Ghana cedi, then apply the dark theme.",
        concepts=[
            ("One table, one grain", "tblDatinSales has one row per order and an Order_ID key."),
            ("Use the right denominator", "Gross margin divides total gross profit by total net sales in the current selection."),
            ("Filters must control the whole view", "The KPIs and all charts should respond to the same Region, Channel and Product selections."),
            ("Protect the analytical structure", "Check the same filter states before and after the theme activity."),
        ],
        sections=[
            section(
                "Sales and Financial Measures",
                "The data represents a fictional Ghana distributor. Sales and finance fields live on the same order row. Net sales is Units × List Price − Discount. Gross profit subtracts cost of goods sold from net sales.",
                "Gross margin is a weighted rate computed from totals. Averaging order-level percentages gives each order equal weight regardless of its value. Gross profit excludes other operating costs and taxes.",
                ["Verify the table name, row grain and unique order IDs.",
                 "Sum Net_Sales_GHS and Gross_Profit_GHS in the selected view.",
                 "Count Order_ID values. Divide total Gross Profit by total Net Sales and format the ratio as a percent."],
                ["60 units at GHS 100 minus GHS 300 in discounts gives GHS 5,700 net sales.",
                 "Cost of goods sold is GHS 3,600. Gross profit is GHS 2,100 and gross margin is 36.84%.",
                 f"Lab 2 all-data baseline is {headline(totals(list(records(FIRST))))}."],
                ["Finance and sales values use the same grain and currency."], ["Sales and finance tables have different grains and are joined without an aggregation plan."],
                [EXCEL_URL],
                brief=lesson(["Net sales = 60 × GHS 100 − GHS 300 = GHS 5,700.",
                              "Gross profit = GHS 5,700 − GHS 3,600 = GHS 2,100."],
                             ["Gross margin = 2,100 ÷ 5,700 = 36.84%.",
                              "Reconcile totals from the source before interpreting a chart."]),
            ),
            section(
                "Charts and Interactive Controls",
                "A dashboard brings the headline metrics and comparisons onto one sheet. Use a monthly line for trends, bars for regional comparisons and doughnuts for product and channel shares of revenue. Region, Channel and Product inputs define the subset shown.",
                "A filter that changes one chart while the KPI still shows all data makes the dashboard misleading. Slicers can work across PivotTables that share a data source. Drop-downs with formula summaries provide another option.",
                ["Start from tblDatinSales and use Month_Start dates in order.",
                 "Connect all KPI and chart summaries to the same filters.",
                 "Start at All, select Greater Accra and Wholesale, then add Rice. Reset all selectors."],
                ["A wholesale manager selects Greater Accra and Wholesale. The KPI totals and each chart update. Adding Rice isolates one product across the selection.",
                 "The Data Guide tab shows independent totals for that selection."],
                ["A manager needs to compare subsets without editing the input table."], ["Users cannot tell which filters apply."],
                [SLICERS_URL, EXCEL_URL],
                brief=lesson(["One table feeds three shared controls, four KPIs and four charts."],
                             ["The chart source is linked to the results, and Month_Start sorts chronologically.",
                              "Test All, Greater Accra and Wholesale, then add Rice and reset filters."],
                             flow=dict(title="Dashboard Data Flow", steps=["Sales table", "Shared filters", "KPI and chart summaries", "Dashboard sheet"])),
            ),
            section(
                "Restyle Without Breaking the Dashboard",
                "A dark theme changes the workbook presentation. The Datin design uses navy panels, gold accents and high-contrast white labels.",
                "Preserving formulas, chart series and filter logic keeps the visual redesign comparable with the earlier view. Legibility depends on font size, contrast and clear selected states as well as palette.",
                ["Save a pre-theme copy and record selected and unfiltered totals.",
                 "Run Prompt 2 on the dashboard sheet.",
                 "Verify the same KPIs, charts and controls at normal zoom."],
                ["Use navy for the canvas, gold for the main chart series and white for labels.",
                 "Compare the two control selections before and after. Check white axis labels and visibly selected filters."],
                ["The figures and interactions are already correct."], ["A redesign has changed formulas or the dashboard’s data sources."],
                [EXCEL_URL],
                brief=lesson(["Navy #0B1730 · gold #E9BE64 · white #FFFFFF.",
                              "Style the canvas, chart text, accents and filter states."],
                             ["Preserve sheet layout, KPI formulas, chart series and filter connections. Confirm by retesting."]),
            ),
        ],
    ),
    dict(
        num=3, code="03", title="Skills",
        subtitle="Convert Prompt 2 into a custom skill, install it through OneDrive and use it on fresh data.",
        concepts=[
            ("A skill packages a repeatable method", "SKILL.md states when the skill applies, what to do and how to verify it."),
            ("The folder name matches the skill name", "Put SKILL.md inside datin-dark-sales-dashboard in the OneDrive folder opened by Excel."),
            ("Enable it for automatic selection", "Refresh and turn the custom skill on, then test with a short prompt."),
            ("Retest on a fresh workbook", "A new July–December dataset proves the skill can adapt to fresh source values."),
        ],
        sections=[
            section(
                "Convert a Prompt into a Skill",
                "A custom Excel skill stores a name and description in YAML frontmatter at the top of SKILL.md, followed by instructions. The description guides selection. The body states the procedure and the expected result.",
                "Prompt 2 restyles a dashboard that already exists. A short create request on a clean workbook needs a build branch before those styling instructions can run. Put both branches in the skill with a clear trigger for each.",
                ["Capture the dark theme and presentation-preservation rules from Prompt 2.",
                 "Add a create-if-missing branch that builds four KPIs, four charts and three filters.",
                 "Add checks that work with either branch."],
                ["Restyle mode targets the dashboard that already exists. Build mode reads tblDatinSales, creates a dashboard and styles it.",
                 f"On the H2 workbook the baseline is {headline(totals(list(records(FRESH))))}."],
                ["A dashboard standard is reused across reporting periods."], ["A one-off request is too specific to reuse."],
                [SKILLS_URL],
                brief=lesson(["Start with Prompt 2. Add the trigger, workbook inputs, create branch and verification steps."],
                             ["On H1 the skill can restyle the existing view. On fresh H2 it first builds the interactive dashboard."]),
            ),
            section(
                "Install, Enable and Invoke the Skill",
                "Excel discovers custom skills in a specific OneDrive folder. Excel’s Custom skills dialog opens the correct account folder. The skill subfolder name must match the SKILL.md metadata name.",
                "Refresh asks Excel to discover the upload. Turning the skill on lets Copilot select it by default when relevant. An explicit @mention lets the learner invoke it and diagnose cases where the short prompt does not trigger it.",
                ["Open the OneDrive skills folder through Excel.", "Upload the named skill folder.", "Refresh and turn on the skill.",
                 "Try a short prompt in a fresh workbook. If needed, choose the explicit @mention and rerun from a clean copy."],
                ["Short request: Create an interactive Datin sales dashboard.",
                 "Check the July–December dates and the filtered GHS totals before keeping the workbook."],
                ["The work account exposes Excel custom skills and supports editing."], ["A test would modify the only copy of a production workbook."],
                [SKILLS_URL],
                brief=lesson(["Create an interactive Datin sales dashboard."],
                             ["Check that the skill is enabled, the dark theme appears and all controls work.",
                              "If automatic selection misses, use the explicit @mention in a fresh copy."],
                             left_head="SHORT PROMPT",
                             flow=dict(title="OneDrive Skill Installation", steps=["Open skills folder", "Upload folder", "Refresh", "Turn on", "Try short prompt"])),
            ),
        ],
    ),

    dict(
        num=4, code="04", title="Analyse a Financial Report with Copilot",
        subtitle="Review the Bank of Ghana’s 2025 financial statements and document your checks in Word.",
        concepts=[
            ("Report scope", "Keep Bank and Group, 2025 and 2024, and original GH¢’000 units separate."),
            ("Arithmetic checks", "Recompute comprehensive income, equity and the cash bridge."),
            ("Attribution", "Reconcile total Group profit or loss to its parent and non-controlling interest components."),
            ("Evidence and review", "Trace each claim to the report, record open items and approve the Word memo."),
        ],
        sections=[
            section(
                "Ground Copilot in the Published Report",
                "Attach the official Bank of Ghana PDF in Microsoft 365 Copilot Chat. Ask for exact row labels, signed amounts, source pages and note numbers.",
                "The PDF contains four adjacent year and entity columns. A plausible answer can still cite the wrong year, subtotal or entity.",
                ["Open the report and identify the 2025 Bank and 2025 Group columns.",
                 "Run the extraction prompt, then check each copied figure against the PDF page.",
                 "Record printed and viewer pages when you share a citation."],
                ["The 2025 Bank operating cash flow appears on printed page 23, which is page 24 in the PDF viewer.",
                 "Note 2(a)(i) begins on printed page 25. The auditor’s emphasis of matter refers to that accounting basis."],
                ["The official source is available and the attachment is authorised."],
                ["A cited figure cannot be located in the report or the response combines Bank and Group."],
                [B.REPORT_URL, B.ATTACH_URL, B.FORMATS_URL],
                brief=lesson(["Attach PDF → extract → verify → record."],
                             ["Scope each figure by year, entity, units and printed page."]),
            ),
            section(
                "Reconcile and Record the Evidence",
                "Reperformance means doing the calculation independently using the published signed amounts.",
                "Recomputation checks the maths. The accounting labels and auditor’s report explain the scope of a figure or conclusion.",
                ["Recompute Bank total comprehensive income, closing equity and closing cash.",
                 "Trace the directors’ Group loss summary to the attribution lines on the profit or loss statement.",
                 "Review the auditor’s stated basis and key audit matter. Record verified findings and unresolved points in Word."],
                ["Assets less liabilities equals negative equity for the Bank. Compare with the Bank equity roll-forward.",
                 "Total Group loss includes parent and non-controlling interest. Cite printed pages 5 and 15 for the label difference."],
                ["The figure, formula and relevant source can be retained in the workpaper."],
                ["The task requires independent audit evidence or assurance beyond the published report."],
                [B.REPORT_URL],
                brief=lesson(["Show the arithmetic and cite the source statement."],
                             ["Separate reported facts, calculations and interpretations."]),
            ),
        ],
    ),
]


def SCHEDULE(lab_titles):
    return {1: ("Agentic AI, financial dashboards, reusable skills and financial report review", [
        ("11:00", "11:10", 10, "admin", "Set up the Datin materials and confirm access to Meta Muse, Excel and OneDrive."),
        ("11:10", "12:30", 80, "topic", "Topic 1. AI history, agents, tools and skills, always-on systems, collaboration patterns and business use cases."),
        ("12:30", "13:30", 60, "lunch", "Lunch break"),
        ("13:30", "13:45", 15, "topic", "Topic 1 continued. Introduce the Business Model Canvas, present the Tema distributor case, show the completed sample and explain the pilot choice."),
        ("13:45", "14:05", 20, "lab", lab_titles([1])),
        ("14:05", "14:20", 15, "break", "Tea break after Topic 1"),
        ("14:20", "14:45", 25, "topic", "Topic 2. Explain the order table, net sales, COGS, gross margin, chart choice and shared filters."),
        ("14:45", "15:45", 60, "lab", lab_titles([2]) + ". Trainer dashboard build and learner theme activity."),
        ("15:45", "16:00", 15, "break", "Tea break after Topic 2"),
        ("16:00", "16:15", 15, "topic", "Topic 3. Explain the skill trigger, build and restyle branches, OneDrive discovery and invocation."),
        ("16:15", "17:05", 50, "lab", lab_titles([3])),
        ("17:05", "17:20", 15, "topic", "Topic 4. Demonstrate the report attachment, statement routes and an independent check."),
        ("17:20", "17:50", 30, "lab", lab_titles([4])),
        ("17:50", "18:00", 10, "recap", "Recap LO1–LO4 and apply five AI security controls before sharing financial outputs."),
    ])}


LG_INTRO = (
    "This Datin edition is for nine accounting professionals from West Africa. Work as the finance team of a fictional wholesale distributor in Tema, Ghana, using only mock month-end and sales data. "
    "First use Meta Muse to prioritise cut-off evidence, a duplicate invoice and a credit-note request. "
    "Then shift from exception triage to management reporting. Use the same distributor’s mock H1 orders to build and restyle a sales dashboard and reconcile its figures in Ghana cedi (GHS)."
)
LG_INTRO2 = (
    "Lab 3 packages the H1 dashboard style and build steps into a OneDrive skill. "
    "Use the distributor’s fresh July–December orders to check whether the same reporting workflow produces accurate new figures. In Topic 4, move from mock sales data to the published Bank of Ghana report and verify a Word review memo."
)
LG_SETUP = dict(
    needs=["Use Microsoft 365 Excel with Copilot editing, a permitted work account and OneDrive. Verify that Manage plugins & skills is visible.",
           "Check Meta Muse access or join the trainer demonstration. Keep the supplied mock material and lab folders together.",
           "Open prompts/index.html from the Datin folder for the copy-ready prompts."],
    verify_text="Open the Lab 2 starter as an editable OneDrive copy with AutoSave on. Check the Sales Data and Data Guide tabs. In Copilot, verify that workbook editing is available.",
    conventions=["GHS means Ghana cedi. Use the Data Guide to reconcile All, Greater Accra/Wholesale and Greater Accra/Wholesale/Rice.",
                 "Use class-only copies and record the before and after values when Copilot modifies a dashboard."],
)
LAB_NOTE = "Use class copies of all workpapers and use the specified official report for Topic 4."
LG_WRAPUP = {"title": "AI Security in Finance", "sections": [{"title": "Five controls", "bullets": [f"{name} · {detail}" for name, detail in B.SECURITY]}]}
LG_GLOSSARY = [
    ("Agentic AI", "Goal-directed behaviour using planning, tools and feedback."),
    ("AI agent", "A configured software worker with a goal and permission boundaries."),
    ("Workflow", "A sequence of steps that people or software can run."),
    ("Automated workflow", "Software runs defined steps. A rule-based example follows preset rules and routes unresolved exceptions for review."),
    ("Copilot", "An assistant that helps within a workflow while a person steers the steps and checks the result."),
    ("Robotic process automation (RPA)", "One way to automate steps, often by operating existing applications through their interfaces."),
    ("Tool", "An operation an agent can call, such as retrieving data or creating a chart."),
    ("Skill", "A reusable instruction package that describes when and how to work."),
    ("COGS", "Cost of goods sold."),
    ("Gross margin", "Total gross profit divided by total net sales for the same selection."),
    ("Slicer", "A button-based filter for an Excel table or compatible PivotTables."),
]
LAB_SHOTS = {2: [("dashboard.png", "Finished Datin Sales Dashboard",
                  "Region, Channel and Product drive all four KPIs and charts.")] }
NEXT_STEPS = {"title": "Reuse the Datin Workflow", "items": [
    "Save the reviewed dashboard and record the filter checks.",
    "Keep the skill folder in OneDrive and test it on a new class workbook.",
]}
THANK_YOU = {"body": "Take the same goal, permissions and verification checks into your next financial dashboard.", "kicker": "DATIN EDITION"}
