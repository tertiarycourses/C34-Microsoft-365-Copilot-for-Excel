"""Topic 3 labs for C34: Use Copilot to Create Dashboards (BrightPath recruitment and workforce data)."""

from mock_data import EXPECTED
from data_domain1 import _LOG, _SETUP_STEP

E5 = EXPECTED[5]
E6 = EXPECTED[6]
_BM = E5["by_month"]
_RD = E6["by_dept"]["Recruitment"]

DOMAIN3 = [
    # ------------------------------------------------------------------ LAB 5
    dict(
        num=5,
        topic=3,
        title="Build a Recruitment KPI Dashboard with Copilot",
        duration=45,
        objective=(
            "LO3: define recruitment KPIs, use Copilot Plan and Allow editing modes to build a dashboard of KPI cards, "
            "PivotTables, charts and slicers, and verify every KPI against an independent check"
        ),
        goal=(
            "Deliver a one-page H1 2026 recruitment dashboard whose five KPI cards, three charts and two slicers all agree "
            "with independently calculated checks."
        ),
        workflow=["Define KPIs", "Build checks", "Plan with Copilot", "Build in Edit", "Verify & slice"],
        desc=(
            "BrightPath's Head of Recruitment wants one page that shows H1 2026 applications, hires, offer acceptance rate, "
            "time to hire and cost per hire, by month, sector, source and recruiter. You will write the checks first, ask "
            "Copilot to plan the dashboard, review the plan against the KPI definitions, let Copilot build it in Allow "
            "editing mode, and then prove each KPI and slicer works."
        ),
        build=(
            "BrightPath_Recruitment_Dashboard_Data.xlsx with a Checks sheet and a Dashboard sheet holding five formula-driven "
            "KPI cards, charts for hires vs target by month, hires by sector and applications by source, and Sector and "
            "Recruiter slicers connected to every PivotTable"
        ),
        services="Excel · Copilot (Plan, Allow editing) · PivotTables, PivotCharts, slicers · BrightPath_Recruitment_Dashboard_Data.xlsx",
        data_files=["data/BrightPath_Recruitment_Dashboard_Data.xlsx"],
        prerequisites=[
            "Topic 2 completed: you can build and verify PivotTables and charts with Copilot.",
            "This lab folder's data/BrightPath_Recruitment_Dashboard_Data.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Recruitment_Dashboard_Data.xlsx")
                + " Read the KPI definitions row on About_This_Data carefully - the dashboard must follow them exactly.",
                "Tables : tblApps (600 applications, Jan-Jun 2026) | tblTargets (monthly hire targets)\n"
                "KPI definitions:\n"
                "  Applications          = number of rows\n"
                "  Hires                 = Furthest_Stage = \"Hired\"\n"
                "  Offer acceptance rate = accepted / (accepted + declined)   (Pending excluded)\n"
                "  Avg time to hire      = average Time_to_Hire_Days of hires\n"
                "  Cost per hire         = total Sourcing_Cost / hires",
            ),
            (
                "Before Copilot builds the dashboard, choose Allow editing and ask Copilot to create a Checks sheet with these "
                "independent KPI formulas, exactly as written. These are the numbers the dashboard must match.",
                "Create a sheet named Checks. Enter these labels and formulas exactly as written, then tell me each result:\n"
                "A1 Applications        B1 =ROWS(tblApps[Application_ID])\n"
                "A2 Hires               B2 =COUNTIFS(tblApps[Furthest_Stage],\"Hired\")\n"
                "A3 Offer acceptance    B3 =COUNTIFS(tblApps[Offer_Accepted],\"Yes\")/(COUNTIFS(tblApps[Offer_Accepted],\"Yes\")+COUNTIFS(tblApps[Offer_Accepted],\"No\"))\n"
                "A4 Avg time to hire    B4 =AVERAGEIFS(tblApps[Time_to_Hire_Days],tblApps[Furthest_Stage],\"Hired\")\n"
                "A5 Cost per hire       B5 =SUM(tblApps[Sourcing_Cost])/B2",
            ),
            (
                "Choose Plan. Paste the dashboard brief below. Read Copilot's plan: does it use the same KPI definitions (especially "
                "excluding Pending offers) and count hires by Hire_Date month? Reply with corrections until the plan matches. "
                "Do not approve yet.",
                "GOAL: Build a one-page H1 2026 recruitment KPI dashboard on the Dashboard sheet.\n"
                "CONTEXT: For BrightPath's Head of Recruitment; monthly review meeting.\n"
                "DATA: tblApps and tblTargets only. Use the KPI definitions on About_This_Data.\n"
                "CONSTRAINTS: Do not change tblApps or tblTargets. KPI cards must be formulas, not typed numbers.\n"
                "OUTPUT:\n"
                "  Row 1: title \"Recruitment Dashboard - Jan to Jun 2026\" and \"Data as at 2026-06-30\".\n"
                "  Row 3: five KPI cards - Applications, Hires, Offer acceptance rate, Avg time to hire (days),\n"
                "         Cost per hire (S$).\n"
                "  Below: (1) column chart of hires by Hire_Date month with a line for Hire_Target from tblTargets,\n"
                "         (2) bar chart of hires by Sector sorted highest first,\n"
                "         (3) bar chart of applications by Source.\n"
                "  Slicers for Sector and Recruiter connected to all PivotTables. Put PivotTables on a sheet Dash_Data.\n"
                "REVIEW: List each KPI formula you will use before building.",
            ),
            (
                "When the plan is right, switch to Allow editing and tell Copilot to carry it out. Watch the steps in the pane; "
                "you can select Stop if it goes off track. When it finishes, open Dash_Data and the Dashboard and review what was created.",
                "The plan is approved. Build the dashboard exactly as planned.",
            ),
            (
                "Ask Copilot to compare each KPI card with Checks!B1:B5 (with no slicer selected) and show each card's formula. "
                "Fix any difference with a targeted follow-up prompt - the most common error is including Pending offers.",
                "CHAT ONLY:\n"
                "Compare each KPI card on the Dashboard with Checks!B1:B5, with no slicer selected. Show each card's\n"
                "formula, its value, the Checks value and whether they match.\n\n"
                "FOLLOW-UP EXAMPLE:\n"
                "The Offer acceptance rate card must exclude Offer_Accepted = \"Pending\". Change its formula to\n"
                "accepted / (accepted + declined) and show it as a percentage with 1 decimal.",
            ),
            (
                "Test interactivity. Ask Copilot to set the Sector slicer to Technology and report every KPI card, then to add "
                "the COUNTIFS checks for Technology and compare. Finally ask it to clear the slicer.",
                "ALLOW EDITING:\n"
                "Set the Sector slicer on the Dashboard to Technology only, then tell me the value on every KPI card.\n\n"
                "CHECK (Allow editing):\n"
                "On the Checks sheet, add these formulas in D1:D2 with labels in C1:C2, and tell me if they match the cards:\n"
                "=COUNTIFS(tblApps[Sector],\"Technology\")\n"
                "=COUNTIFS(tblApps[Sector],\"Technology\",tblApps[Furthest_Stage],\"Hired\")\n\n"
                "ALLOW EDITING:\n"
                "Clear the Sector slicer so the dashboard shows all sectors.\n\n"
                "If a KPI card does not change, ask: \"Make the KPI cards respond to the Sector and Recruiter slicers.\"",
            ),
            (
                "Polish with one prompt per change, then ask Copilot (Chat only) to review the dashboard for clarity. Ask "
                "Copilot to log all prompts on Prompt_Log.",
                "Add a note under the title: \"Source: tblApps; KPI definitions on About_This_Data\".\n"
                "Use one consistent colour scheme and align the charts in a grid.\n"
                "CHAT ONLY: Review this dashboard for a Head of Recruitment. What is unclear or missing?\n\n" + _LOG,
            ),
        ],
        sample_prompts=[
            ("Plan", "Plan a recruitment KPI dashboard on the Dashboard sheet from tblApps with KPI cards for applications, hires, offer acceptance rate, average time to hire and cost per hire. List the formulas before building."),
            ("KPI card", "Create a KPI card that shows Cost per hire = total Sourcing_Cost / hires, formatted as S$ with 2 decimals."),
            ("Chart vs target", "Create a column chart of hires by Hire_Date month with Hire_Target from tblTargets as a line."),
            ("Slicers", "Add slicers for Sector and Recruiter and connect them to all PivotTables on Dash_Data."),
            ("Fix a KPI", "The offer acceptance rate must exclude Pending offers. Update the formula and explain it."),
            ("Layout", "Arrange the KPI cards in one row and the three charts in a grid below them with equal sizes."),
            ("Review", "Review this dashboard for a Head of Recruitment. What is unclear, missing or potentially misleading?"),
            ("Question", "Which recruiter made the most hires in H1 2026, and what was their average time to hire?"),
        ],
        test=(
            f"With no slicer selected the KPI cards match Checks: Applications {E5['apps']}, Hires {E5['hires']}, Offer "
            f"acceptance {E5['oar']}% ({E5['accepted']} accepted of {E5['decided']} decided; {E5['pending']} pending excluded), "
            f"Avg time to hire {E5['avg_tth']} days, Cost per hire S${E5['cph']:.2f}. Hires by month: "
            + ", ".join(f"{k} = {v}" for k, v in _BM.items())
            + f" (none in 2026-01). {E5['top_sector']} has the most hires ({E5['by_sector'][E5['top_sector']]}). With Sector = "
            f"Technology, Applications = {E5['tech_apps']}, Hires = {E5['tech_hires']} and Avg time to hire = "
            f"{E5['tech_avg_tth']:.1f} days. {E5['top_rec']} made the most hires ({E5['top_rec_n']}). No KPI card contains a "
            "typed number."
        ),
        checkpoint="Keep the workbook in work-c34. Lab 6 builds a second dashboard from the workforce workbook in its own folder.",
        troubleshooting=[
            (
                "Copilot builds before you approve the plan.",
                "You were in Allow editing mode. Undo (Ctrl+Z) or restore from version history, switch to Plan and resubmit.",
            ),
            (
                "KPI cards do not change when a slicer is used.",
                "Cards built with COUNTIFS on tblApps ignore slicers. Ask Copilot to base them on the PivotTables (GETPIVOTDATA) "
                "or accept that cards show the H1 total and label them so.",
            ),
            (
                "Hires appear in January.",
                "The chart is counting by Applied_Month. Ask Copilot to group hires by Hire_Date month as the KPI definition says.",
            ),
            (
                "Copilot stops part-way or times out.",
                "Ask it to continue from the last completed step, or split the build into KPI cards, then charts, then slicers.",
            ),
        ],
        challenge=(
            "Ask Copilot to add a sixth KPI card for 'Hires vs target %' (hires / sum of targets) and conditional formatting "
            "that turns it red below 90%. Verify it with your own formula."
        ),
        reflection="Which KPI definition was easiest to get subtly wrong, and how would you stop that happening next month?",
    ),
    # ------------------------------------------------------------------ LAB 6
    dict(
        num=6,
        topic=3,
        title="Build an Interactive Workforce and Attrition Dashboard",
        duration=35,
        objective=(
            "LO3: use Copilot to build a refreshable HR workforce dashboard with helper columns, KPI cards, charts and "
            "slicers, prove it refreshes with new rows, and check it for responsible sharing"
        ),
        goal=(
            "Show headcount, 12-month attrition, new hires and tenure by department and location on one interactive page "
            "that updates correctly when new employees are added."
        ),
        workflow=["Helper columns", "Checks", "Build dashboard", "Refresh test", "Privacy check"],
        desc=(
            "BrightPath's HR Director needs a workforce dashboard for the July leadership meeting. You will ask Copilot to add "
            "period helper columns, type independent checks for headcount and attrition, have Copilot plan and build the "
            "dashboard with slicers, then append eight July new joiners and refresh to prove the dashboard is live. Finally "
            "you will check that no slicer combination exposes a group too small to stay anonymous."
        ),
        build=(
            "BrightPath_Workforce_Headcount.xlsx with helper columns in tblEmployees, a Checks area, and an HR_Dashboard sheet "
            "showing six KPI cards, headcount and attrition-rate charts by department, hires vs exits by quarter, and Location "
            "and Job_Level slicers - refresh-tested with new rows"
        ),
        services="Excel · Copilot (Plan, Allow editing) · PivotTables, slicers, refresh · BrightPath_Workforce_Headcount.xlsx",
        data_files=["data/BrightPath_Workforce_Headcount.xlsx"],
        prerequisites=[
            "Lab 5 completed: you have built and verified a dashboard with Copilot.",
            "This lab folder's data/BrightPath_Workforce_Headcount.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Workforce_Headcount.xlsx")
                + " Do not touch the New_Joiners_Jul2026 sheet until Step 6.",
                "Table : tblEmployees (420 employees hired since 2019; Status Active or Exited)\n"
                "Period: 2025-07-01 to 2026-06-30\n"
                "Attrition rate = exits in period / ((active on 2025-07-01 + active on 2026-06-30) / 2)",
            ),
            (
                "Choose Allow editing and ask Copilot for three helper columns. Then ask it to explain the Active_Start formula "
                "and show one employee who exited before 2025-07-01 (Active_Start must be FALSE).",
                "Add three columns to tblEmployees:\n"
                "1) Active_Start = TRUE if Hire_Date <= 2025-07-01 and (Exit_Date is blank or Exit_Date > 2025-07-01).\n"
                "2) Exit_In_Period = TRUE if Exit_Date is between 2025-07-01 and 2026-06-30 inclusive.\n"
                "3) Tenure_Years = years from Hire_Date to Exit_Date, or to 2026-06-30 if Exit_Date is blank, 1 decimal.\n\n"
                "FOLLOW-UP: Explain the Active_Start formula, and show one employee who exited before 2025-07-01\n"
                "with Hire_Date, Exit_Date and Active_Start.",
            ),
            (
                "Ask Copilot to create a Checks sheet with these independent checks (label in column A, formula in column B), "
                "exactly as written. These are the numbers the dashboard must match.",
                "Create a sheet named Checks. Put these labels in column A and formulas in column B exactly as written,\n"
                "then tell me each result. Build the Attrition rate formula from the cells above it:\n"
                "Active headcount     =COUNTIFS(tblEmployees[Status],\"Active\")\n"
                "Headcount 2025-07-01 =COUNTIFS(tblEmployees[Active_Start],TRUE)\n"
                "Exits in period      =COUNTIFS(tblEmployees[Exit_In_Period],TRUE)\n"
                "Voluntary exits      =COUNTIFS(tblEmployees[Exit_In_Period],TRUE,tblEmployees[Exit_Type],\"Voluntary\")\n"
                "Attrition rate       = Exits in period / ((Headcount 2025-07-01 + Active headcount) / 2)\n"
                "Hires in period      =COUNTIFS(tblEmployees[Hire_Date],\">=\"&DATE(2025,7,1),tblEmployees[Hire_Date],\"<=\"&DATE(2026,6,30))",
            ),
            (
                "Choose Plan, paste the brief, review the plan against the definitions, then switch to Allow editing to build.",
                "Plan an HR workforce dashboard on the HR_Dashboard sheet from tblEmployees for the HR Director.\n"
                "KPI cards (formulas): Active headcount, Exits (last 12 months), Attrition rate, Voluntary attrition rate,\n"
                "Avg tenure of active staff (years), Hires (last 12 months). Use the definitions on About_This_Data.\n"
                "Charts: active headcount by Department (bar, sorted); attrition rate by Department (bar); hires vs exits\n"
                "by quarter for the last 12 months (clustered column); active headcount by Location (bar).\n"
                "Slicers: Location and Job_Level, connected to all PivotTables. Title and \"Data as at 2026-06-30\".\n"
                "Do not show Employee_ID, salary or individual records on the dashboard.",
            ),
            (
                "Ask Copilot to compare the KPI cards with Checks, and fix any difference with a follow-up prompt. Then ask "
                "which Location, Job_Level and Department combinations contain fewer than 5 people.",
                "CHAT ONLY:\n"
                "Compare each KPI card on HR_Dashboard with the Checks sheet and tell me which, if any, do not match.\n\n"
                "FOLLOW-UP (if needed):\n"
                "The attrition rate must use average headcount = (Active_Start count + active count) / 2.\n\n"
                "CHAT ONLY:\n"
                "Which combinations of Location, Job_Level and Department contain fewer than 5 active employees?",
            ),
            (
                "Refresh test: ask Copilot to append the 8 rows from New_Joiners_Jul2026 to tblEmployees and refresh every "
                "PivotTable. Every KPI and chart that counts active staff must update.",
                "ALLOW EDITING:\n"
                "Copy the 8 data rows in New_Joiners_Jul2026!A2:L9 to the first empty rows directly below tblEmployees\n"
                "so the table expands. Then refresh all PivotTables in the workbook and tell me the new Active headcount\n"
                "and the number of rows in tblEmployees.\n\n"
                "Expected after refresh:\n"
                f"Active headcount : {E6['active']} -> {E6['active_after']}\n"
                f"tblEmployees rows: {E6['rows']} -> {E6['rows_after']}\n"
                "New joiners by department: "
                + ", ".join(f"{k} +{v}" for k, v in E6["nj_by_dept"].items()),
            ),
            (
                "Ask Copilot (Chat only) to review the dashboard for privacy and clarity, ask it to act on its valid points, and "
                "ask it to log all prompts on Prompt_Log.",
                "CHAT ONLY:\n"
                "Review HR_Dashboard before I share it with the leadership team. Does any element expose individual\n"
                "employee data or a group smaller than 5? Is each KPI labelled with its period and definition?\n\n"
                + _LOG,
            ),
        ],
        sample_prompts=[
            ("Helper column", "Add Active_Start = TRUE if Hire_Date <= 2025-07-01 and Exit_Date is blank or after 2025-07-01."),
            ("Helper column", "Add Tenure_Years from Hire_Date to Exit_Date, or to 2026-06-30 if Exit_Date is blank, rounded to 1 decimal."),
            ("Plan", "Plan an HR workforce dashboard with KPI cards for active headcount, exits, attrition rate, voluntary attrition, average tenure and hires."),
            ("Chart", "Create a bar chart of 12-month attrition rate by Department, sorted highest first, with data labels as percentages."),
            ("Slicers", "Add slicers for Location and Job_Level and connect them to every PivotTable on the dashboard."),
            ("Refresh", "I added rows to tblEmployees. Refresh all PivotTables and tell me which KPI values changed."),
            ("Privacy", "Which slicer combinations show a group of fewer than 5 employees? Suggest how to suppress them."),
            ("Explain", "Explain how the Attrition rate card is calculated, step by step."),
        ],
        test=(
            f"Before the refresh: Active headcount {E6['active']}, headcount on 2025-07-01 {E6['start_hc']}, exits in period "
            f"{E6['exits_12m']} ({E6['vol_exits']} voluntary), attrition rate {E6['attrition']}% and voluntary attrition "
            f"{E6['vol_attrition']}%, hires in period {E6['hires_12m']}, average tenure of active staff about "
            f"{E6['avg_tenure']} years. {E6['top_attr_dept']} has the highest attrition rate ({_RD['rate']}%, {_RD['exits']} exits) "
            f"and the largest active headcount ({_RD['hc']}). After appending the July joiners and refreshing, Active headcount "
            f"shows {E6['active_after']} and tblEmployees has {E6['rows_after']} rows. The dashboard shows no Employee_ID, salary "
            "or individual record."
        ),
        checkpoint="Keep the workbook in work-c34. Topic 4 starts from the attrition-analysis workbook in the Lab 7 folder.",
        troubleshooting=[
            (
                "The new rows did not join the table and nothing changed on refresh.",
                "Ask Copilot to resize tblEmployees to include the 8 new rows, then refresh all PivotTables again.",
            ),
            (
                "Attrition rate is far too high (e.g. 29 / 237).",
                "The card divides by closing headcount only. Ask Copilot to use the average of opening and closing headcount.",
            ),
            (
                "Helper columns show #VALUE! for blank Exit_Date.",
                "Ask Copilot to treat blank Exit_Date explicitly, e.g. OR([@Exit_Date]=\"\",[@Exit_Date]>DATE(2025,7,1)).",
            ),
            (
                "A chart did not update after the refresh.",
                "It is a static picture or a range copy. Ask Copilot to replace it with a PivotChart linked to tblEmployees.",
            ),
        ],
        challenge=(
            "Ask Copilot to add a Timeline control on Hire_Date and a chart of average Engagement_Score by Department for "
            "active staff, then check one department's average with AVERAGEIFS."
        ),
        reflection="What would happen to trust in this dashboard if one leader could filter it down to a single named team member?",
    ),
]
