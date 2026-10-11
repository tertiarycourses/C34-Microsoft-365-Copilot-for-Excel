"""
"What it does" text for every step prompt on the prompts page, keyed by lab number
and prompt letter (the letters prompt_rows() assigns in step order).

build_prompts_page.py fails the build if a prompt has no explanation or an
explanation has no prompt, so after adding, removing or reordering a lab's prompts,
update the letters here to match.
"""

_LOG = ("Fills in your Prompt_Log: for each prompt you ran, what Copilot did and how you checked it. "
        "This is your record of the lab.")

EXPLAIN = {
    1: {
        "A": "Turns the raw Applications data into an Excel table named tblApplications, so every later prompt can "
             "refer to it by name. Copilot reports the columns and rows so you can confirm nothing was lost.",
        "B": "Asks Copilot to describe the data before you analyse it: what a row is, each column's type, the row "
             "count, the date range and the Stage and Source values. Chat only, so nothing changes.",
        "C": "Looks for data-quality problems - duplicate Application_IDs and misspelt Source labels - and counts "
             "each. It only reports; nothing is fixed yet.",
        "D": "Copilot shows its plan first, then highlights duplicate IDs in orange and adds a cleaned Source_Clean "
             "column. The original values stay untouched.",
        "E": "Checks whether each duplicate is a true copy or two different applications sharing an ID, with the "
             "row numbers of both.",
        "F": "Highlights in yellow the Source cells the cleanup changed and lists their IDs, so you can see exactly "
             "what was fixed.",
        "G": "Confirms the cleanup worked: no 'Linked In', no trailing spaces, and the table still has 242 rows.",
        "H": "Copilot's answer: applications, hires and hire rate by source, sorted by hires. You verify it with "
             "formulas in the next prompt.",
        "I": "Adds two COUNTIFS formulas on a Checks sheet so you can verify Copilot's hire counts with formulas "
             "you can see.",
        "J": "Asks the same question in a new chat with a different model. If the answers differ, the formulas on "
             "Checks decide which is right.",
        "K": "Highlights every hired candidate's row in light green so the hires stand out.",
        "L": "Sorts the table by interview score (highest first), then by days in the pipeline (shortest first).",
        "M": "Filters the table to Technology candidates at the Interview stage.",
        "N": "Asks how many rows the filter shows. Compare it with the count Excel shows at the bottom of the window.",
        "O": "Clears every filter and confirms all 242 rows are back, so later steps see the full data.",
        "P": _LOG,
    },
    2: {
        "A": "Adds a Recruiter_Name column that looks up each requisition's recruiter in tblRecruiters with "
             "XLOOKUP, showing \"Not found\" for an ID with no match.",
        "B": "Copilot explains the XLOOKUP in plain English and shows one row next to its match in tblRecruiters, "
             "so you can check the lookup by eye.",
        "C": "Adds Days_Open: how long each requisition has been open, using the report date in Control!B1 for "
             "those still open.",
        "D": "Shows one open and one closed requisition with the subtraction behind each, so you can confirm both "
             "halves of the rule.",
        "E": "Adds Fill_Rate - positions filled divided by openings - as a percentage.",
        "F": "Adds Est_Fee from the fee rule (filled x monthly salary x 12 x fee rate) in S$, and explains the "
             "formula.",
        "G": "Walks through the fee calculation for one requisition, JOB-2002, so you can check it by hand.",
        "H": "Adds an Ageing_Flag that marks open requisitions older than 45 days as \"Overdue\" and highlights "
             "those rows amber.",
        "I": "Asks how many requisitions are flagged Overdue. The next prompt checks this with an independent "
             "formula.",
        "J": "Puts independent control formulas on the Control sheet. Each pair must match; if one does not, one of "
             "Copilot's columns is wrong.",
        "K": "Asks which recruiter generated the most estimated fees, with every recruiter's total. Chat only.",
        "L": "Adds a SUMIFS check for the top recruiter so a formula confirms Copilot's answer. Replace <name> with "
             "the recruiter Copilot named.",
        "M": _LOG,
    },
    3: {
        "A": "Builds a PivotTable of total hires by month on Pivot_Trend, then a cleanly formatted line chart from "
             "it on the Charts sheet (top left). Answers Q1 Trend.",
        "B": "Adds a SUM of all hires and asks whether it matches the PivotTable grand total, and which month had "
             "the most hires.",
        "C": "Builds a PivotTable of 2026 hires by sector, then a bar chart from it on the Charts sheet (top right). "
             "Answers Q2 Comparison.",
        "D": "Sorts the bars from highest to lowest and makes the title say which sector led.",
        "E": "Adds a SUMIFS of Technology's 2026 hires. It must match Technology's bar.",
        "F": "Builds a 2026 funnel table (Applications to Hires) with SUMIFS on a new sheet, then a funnel chart on "
             "the Charts sheet (bottom left). Answers Q3 Funnel.",
        "G": "Shows the formula behind each funnel number, so you can confirm they are formulas, not typed values.",
        "H": "Adds Hire_Rate and Cost_per_Hire to tblSources, then a combo chart - rate as columns, cost as a line "
             "on a second axis - on the Charts sheet (bottom right). Answers Q4 Efficiency.",
        "I": "Asks for the best hire rate and the lowest cost per hire, with values, to compare with the chart.",
        "J": "Rewrites each chart title to state its finding, e.g. Technology led hiring in H1 2026.",
        "K": "Makes hires the same blue in every chart and adds thousands separators.",
        "L": "Makes the four charts the same size and lines them up in a 2 x 2 grid.",
        "M": "Adds alt text to every chart so screen-reader users know what it shows.",
        "N": "Copilot reviews each chart against the four questions: does the chart type fit, and one improvement. "
             "Chat only.",
        "O": _LOG,
    },
    4: {
        "A": "Lists the messy values in Status, Region and Client and groups the ones that mean the same thing "
             "(e.g. WIP and In Progress). Chat only - this is your cleanup checklist.",
        "B": "Writes the workbook rules into the .Rules sheet. Copilot reads them on every later prompt, so you do "
             "not have to repeat them.",
        "C": "Creates Orders_Clean, a cleaned copy of every order that follows the rules: standard statuses, "
             "regions, client names, dates and money.",
        "D": "Adds check formulas that count orders and distinct values and total the budget, so you can confirm "
             "the cleanup.",
        "E": "Adds Rate (looked up by Owner), Delivered_Value, Budget_Used and an Over_Budget flag to the clean "
             "table.",
        "F": "Adds two more checks - total delivered value and the over-budget count - then lists the over-budget "
             "orders.",
        "G": "Adds a new rule to .Rules: highlight over-budget rows in light red.",
        "H": "Re-applies every workbook rule to tblOrdersClean, including the new highlight rule.",
        "I": _LOG,
    },
    5: {
        "A": "Asks Copilot which of its built-in skills fit this job: the finance analysis and the model audit. "
             "Chat only.",
        "B": "Uses the chart-design skill to create a combo chart: MRR as columns and client count as a line on a "
             "second axis.",
        "C": "Uses chart-design again to polish the chart for a leadership slide: a finding-style title, axis "
             "titles and one data label.",
        "D": "The unit-economics skill plans, then calculates CAC, ARPA, margin, churn, LTV, LTV:CAC and payback on "
             "a new sheet and lists its assumptions. Review the plan before you approve it.",
        "E": "Builds independent check formulas for each unit-economics measure, so you can verify the skill's "
             "numbers.",
        "F": "Compares the check formulas with the Unit_Economics sheet and explains any difference.",
        "G": "Uses the check-model skill to audit Headcount_Budget for typed numbers, inconsistent formulas and "
             "totals that miss rows. It only reports.",
        "H": "Fixes the issues the audit found and gives the corrected FY2027 total cost.",
        "I": "Uses the theme-design skill to give the four sheets one consistent, professional look.",
        "J": _LOG,
    },
    6: {
        "A": "Copilot writes a complete SKILL.md for an interview-calendar skill: read tblInterviews, build a "
             "weekly calendar, flag clashes and summarise. Chat only - you save the text as the skill file.",
        "B": "Runs your new skill to build the interview calendar for the week of 6 July 2026.",
        "C": "Adds a COUNTIFS of that week's interviews. It must match the number of interviews on the calendar.",
        "D": "Runs the skill again for another week with a short prompt, to show it is reusable.",
        "E": "Updates the skill so the summary also shows interview hours per interviewer and the reply names the "
             "busiest day.",
        "F": _LOG,
    },
    7: {
        "A": "Writes the dashboard rules into .Rules - KPI definitions, layout and colours - so every later prompt "
             "follows them.",
        "B": "Builds independent check formulas for the six KPIs. The dashboard cards must match these.",
        "C": "Copilot plans, then builds the dashboard: six KPI cards, four charts, slicers and a timeline. Review "
             "the plan before you approve it.",
        "D": "Uses chart-design to review every chart: finding-style titles, axis titles, useful data labels and "
             "the colours in the rules.",
        "E": "Adds traffic-light colours (green, amber, red) to the Hires vs target card.",
        "F": "Uses theme-design to give the dashboard a clean, modern look while keeping the rule colours.",
        "G": "Compares each KPI card with the Checks sheet and shows each card's formula. Chat only.",
        "H": "Only if needed: fixes the Offer acceptance card so pending offers are left out, as the rules say.",
        "I": "Adds filtered check formulas to confirm the cards change correctly with the Technology slicer and the "
             "Apr-Jun timeline.",
        "J": "Only if needed: links any KPI card that ignores the slicers to the PivotTables with GETPIVOTDATA.",
        "K": "Copilot reviews the dashboard as a Head of Recruitment would - what is unclear, missing or misleading "
             "- and checks it against the rules. Chat only.",
        "L": _LOG,
    },
    8: {
        "A": "Writes the workforce dashboard rules into .Rules: definitions, privacy limits (no individual data, no "
             "group under 5), layout and colours.",
        "B": "Copilot writes a SKILL.md for an hr-workforce-dashboard skill that builds the whole dashboard with "
             "checks. Chat only - you save the text as the skill file.",
        "C": "Runs your skill to plan, then build the workforce dashboard for the 12 months to 30 June 2026. Review "
             "the plan before you approve it.",
        "D": "Compares each KPI card with the Checks sheet and shows the attrition formula. Chat only.",
        "E": "Only if needed: fixes the attrition rate to use average headcount, as the rules say.",
        "F": "Adds 8 new joiners to the employee table, refreshes every PivotTable and reports the new headcount - "
             "the refresh test.",
        "G": "A privacy review before sharing: does anything show an individual or a group smaller than 5, and is "
             "every KPI labelled? Chat only.",
        "H": "Writes a 5-line leadership summary that uses only the dashboard's numbers and labels any explanation "
             "as a hypothesis.",
        "I": _LOG,
    },
}
