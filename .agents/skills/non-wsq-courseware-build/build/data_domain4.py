"""Topic 4 labs for C34: Use Copilot to Draw Insights and Recommendations (BrightPath workforce data)."""

from mock_data import EXPECTED
from data_domain1 import _LOG, _SETUP_STEP

E7 = EXPECTED[7]
E8 = EXPECTED[8]


def _q(name):
    return name.split("_", 1)[1].replace("_", " ").replace(" and ", " & ")


DOMAIN4 = [
    # ------------------------------------------------------------------ LAB 7
    dict(
        num=7,
        topic=4,
        title="Surface Attrition Insights with Copilot and Advanced Analysis",
        duration=40,
        objective=(
            "LO4: use Copilot direct answers and Python-based advanced analysis to surface attrition patterns, verify them "
            "with native formulas and record each as OBSERVED evidence or a HYPOTHESIS with its limitation"
        ),
        goal=(
            "Identify the work-pattern factors most associated with leaving BrightPath, with verified rates, group sizes and "
            "honest limitations in an Insight_Log."
        ),
        workflow=["Ask broadly", "Set the baseline", "Compare groups", "Advanced analysis", "Log insights"],
        desc=(
            "BrightPath lost 14% of a 400-person study group in 12 months. You will start with an open 'What's interesting?' "
            "prompt, set the baseline attrition rate, ask Copilot for specific group comparisons with counts, enter advanced "
            "analysis mode to compare leavers and stayers with Python, re-compute the headline numbers with COUNTIFS and "
            "AVERAGEIFS, and write four labelled insights - including why a high rate in a small team is only a signal."
        ),
        build=(
            "BrightPath_Attrition_Analysis.xlsx with verified attrition rates by department and factor, an advanced-analysis "
            "sheet with reviewed Python output, and an Insight_Log of four OBSERVED or HYPOTHESIS insights with evidence and "
            "limitations"
        ),
        services="Excel · Copilot (Chat only, advanced analysis mode) · Python in Excel · BrightPath_Attrition_Analysis.xlsx",
        data_files=["data/BrightPath_Attrition_Analysis.xlsx"],
        prerequisites=[
            "Labs 1-6 completed, or equivalent skill with Copilot prompts, PivotTables and COUNTIFS.",
            "Python in Excel available for advanced analysis (fallback in Troubleshooting).",
            "This lab folder's data/BrightPath_Attrition_Analysis.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Attrition_Analysis.xlsx")
                + " Note the warning on About_This_Data: these are team-level signals, never a basis for decisions about an individual.",
                "Table   : tblAttrition (400 employees active on 2025-07-01)\n"
                "Outcome : Left_Company = Yes / No (left within 12 months)\n"
                "Template: Insight_Log sheet",
            ),
            (
                "Choose Chat only and ask the open question. Then ask Copilot (Allow editing) to copy its candidate patterns into "
                "Prompt_Log as unverified - do not record them as insights yet; each must be verified first.",
                "CHAT ONLY:\n"
                "What's interesting about tblAttrition in relation to Left_Company? Give up to five patterns, each with\n"
                "the numbers and group sizes behind it. Do not change the workbook.\n\n"
                "ALLOW EDITING:\n"
                "Copy the patterns you just gave me into Prompt_Log, one row each, marked \"Unverified\".",
            ),
            (
                "Set the baseline: ask for the overall rate, then ask Copilot to add the two check formulas in empty cells.",
                "CHAT ONLY:\n"
                "What is the overall attrition rate in tblAttrition (Left_Company = \"Yes\" divided by all employees)?\n\n"
                "CHECK (Allow editing):\n"
                "In empty cells on Prompt_Log, add these formulas with a label beside each, and tell me the results:\n"
                "=COUNTIFS(tblAttrition[Left_Company],\"Yes\")\n"
                "=COUNTIFS(tblAttrition[Left_Company],\"Yes\")/ROWS(tblAttrition[Employee_ID])",
            ),
            (
                "Ask for specific comparisons with counts. Ask Copilot to add the COUNTIFS check for the engagement split, and "
                "note which departments have fewer than 30 people.",
                "CHAT ONLY:\n"
                "Using tblAttrition, show attrition rate AND number of employees for:\n"
                "1) each Department;\n"
                "2) Engagement_Score below 3 vs 3 and above;\n"
                "3) Overtime_Hours_Month above 20 vs 20 or below;\n"
                "4) Years_Since_Promotion 3 or more vs below 3;\n"
                "5) Compa_Ratio below 0.9 vs 0.9 and above;\n"
                "6) each Work_Mode.\n\n"
                "CHECK (Allow editing):\n"
                "In an empty cell on Prompt_Log, add this formula for Engagement_Score below 3 and tell me if it matches:\n"
                "=COUNTIFS(tblAttrition[Engagement_Score],\"<3\",tblAttrition[Left_Company],\"Yes\")/COUNTIFS(tblAttrition[Engagement_Score],\"<3\")",
            ),
            (
                "Go deeper: ask Copilot to enter advanced analysis mode (or select 'Get deeper results using advanced analysis "
                "mode'). Copilot opens a new sheet with Python. Ask Copilot which rows and columns the code uses and how blanks "
                "are handled. Select Stop advanced analysis, then ask Copilot to add two AVERAGEIFS checks of the means.",
                "Enter advanced analysis mode.\n"
                "Compare employees who left with those who stayed on Overtime_Hours_Month, Engagement_Score,\n"
                "Years_Since_Promotion, Compa_Ratio and Training_Hours. Show the mean for each group, the group sizes,\n"
                "and the correlation between Overtime_Hours_Month and Engagement_Score. Explain the method.\n\n"
                "FOLLOW-UP: Which rows and columns does your code use, and how does it handle blanks?\n\n"
                "CHECK (Allow editing):\n"
                "In empty cells on Prompt_Log, add these formulas and tell me if they match your means for leavers:\n"
                "=AVERAGEIFS(tblAttrition[Overtime_Hours_Month],tblAttrition[Left_Company],\"Yes\")\n"
                "=AVERAGEIFS(tblAttrition[Engagement_Score],tblAttrition[Left_Company],\"Yes\")",
            ),
            (
                "Ask Copilot to help label your findings, then ask it to write them into four rows of the Insight_Log sheet. Each "
                "row needs the verified numbers, a label, a possible explanation to test and a limitation. Read every row and "
                "tell Copilot what to change.",
                "CHAT ONLY:\n"
                "Here are my verified findings: <paste your checked rates and means>.\n"
                "For each, write one OBSERVED sentence with the numbers, one HYPOTHESIS that could explain it, and\n"
                "one limitation (sample size, correlation is not causation, overlapping factors). Keep each under 30 words.\n\n"
                "ALLOW EDITING:\n"
                "Write these four findings into the Insight_Log sheet, one per row, with the numbers, the OBSERVED or\n"
                "HYPOTHESIS label, the explanation to test and the limitation.",
            ),
            (
                "Ask Copilot (Chat only) to challenge your Insight_Log, then ask it to log all prompts on Prompt_Log.",
                "CHAT ONLY:\n"
                "Review my Insight_Log. Which statement overstates the evidence or implies causation? How would\n"
                "you reword it?\n\n" + _LOG,
            ),
        ],
        sample_prompts=[
            ("Open question", "What's interesting about this data in relation to Left_Company? Give five patterns with numbers and group sizes."),
            ("Baseline", "What is the overall attrition rate in tblAttrition? Show the numerator and denominator."),
            ("Comparison", "Show attrition rate and headcount for Engagement_Score below 3 vs 3 and above."),
            ("By group", "Show attrition rate and number of employees by Department, sorted highest first."),
            ("Advanced analysis", "Enter advanced analysis mode. Compare leavers and stayers on overtime, engagement, promotion, compa-ratio and training."),
            ("Correlation", "What is the correlation between Overtime_Hours_Month and Engagement_Score, and what does it mean for interpreting the attrition patterns?"),
            ("Outliers", "Are there any managers (Manager_ID) whose teams have unusually high attrition? Only include teams with at least 10 people."),
            ("Challenge", "Which of my insights could be a coincidence or caused by another factor? Explain."),
        ],
        test=(
            f"Overall attrition = {E7['left']} of {E7['rows']} = {E7['rate']}%. Engagement below 3: {E7['eng_lo'][1]}% "
            f"(n = {E7['eng_lo'][0]}) vs {E7['eng_hi'][1]}% (n = {E7['eng_hi'][0]}). Overtime above 20 hours: {E7['ot_hi'][1]}% "
            f"(n = {E7['ot_hi'][0]}) vs {E7['ot_lo'][1]}% (n = {E7['ot_lo'][0]}). Promotion 3+ years ago: {E7['pr_hi'][1]}% vs "
            f"{E7['pr_lo'][1]}%. On-site staff: {E7['by_mode']['On-site']}%. {E7['top_dept']} shows the highest department rate "
            f"({E7['by_dept'][E7['top_dept']]}%) but has only {E7['dept_n'][E7['top_dept']]} people, so it is recorded as a "
            f"small-sample signal. Leavers average {E7['ot_left']} overtime hours and engagement {E7['eng_left']:.2f}, vs "
            f"{E7['ot_stayed']} and {E7['eng_stayed']:.2f} for stayers. Overtime and engagement correlate at about "
            f"{E7['corr_eng_ot']}, so the two factors overlap. Insight_Log has four rows, each labelled and with a limitation."
        ),
        checkpoint=(
            "Keep the workbook in work-c34 and keep your Insight_Log open - Lab 8 turns verified insights like these into "
            "recommendations."
        ),
        troubleshooting=[
            (
                "Advanced analysis mode is not offered.",
                "It needs Python in Excel. Stay in Chat only: direct answers also use Python and can show code. Or build the "
                "comparisons with a PivotTable (Left_Company in columns, Count of Employee_ID, % of row total).",
            ),
            (
                "Copilot gives a rate without a group size.",
                "Follow up: 'Show the number of employees in each group next to each rate.' Never record a rate without n.",
            ),
            (
                "Copilot says a factor 'causes' attrition.",
                "Ask it to restate as an association, and record any causal idea as a HYPOTHESIS to test (e.g. exit interviews).",
            ),
            (
                "Your COUNTIFS rate differs slightly from Copilot's.",
                "Check the boundary (< 3 vs <= 3, > 20 vs >= 20) and whether blanks were excluded. Use your formula.",
            ),
        ],
        challenge=(
            "In advanced analysis mode, ask Copilot to fit a simple logistic regression of Left_Company on the five factors and "
            "explain the output in plain English. Record why this still does not prove cause."
        ),
        reflection="Which insight would you be most careful presenting to managers, and what extra evidence would you want first?",
    ),
    # ------------------------------------------------------------------ LAB 8
    dict(
        num=8,
        topic=4,
        title="Turn Engagement Survey Feedback into Recommendations",
        duration=30,
        objective=(
            "LO4: use Copilot to quantify survey scores and eNPS, classify comment themes and sentiment with a human "
            "spot-check, and draft owned, measurable recommendations and an executive summary from verified findings"
        ),
        goal=(
            "Produce three evidence-backed recommendations and a 150-word executive summary from BrightPath's June engagement "
            "pulse survey, with every number verified."
        ),
        workflow=["Score & eNPS", "By department", "Theme comments", "Recommend", "Summarise"],
        desc=(
            "BrightPath ran an anonymous pulse survey of 160 staff. You will have Copilot calculate eNPS and question averages "
            "and verify them with COUNTIFS and AVERAGE, compare departments while respecting small groups, add Theme and "
            "Sentiment columns from a fixed list and spot-check them, then use the Recommendations template and Copilot to "
            "turn verified findings into owned actions and an executive summary you approve."
        ),
        build=(
            "BrightPath_Engagement_Survey.xlsx with verified eNPS and question averages, Theme and Sentiment columns with a "
            "recorded spot-check, a Recommendations sheet with three owned, measurable actions, and an Executive_Summary"
        ),
        services="Excel · Copilot (Chat only, Allow editing) · COPILOT function (where available) · BrightPath_Engagement_Survey.xlsx",
        data_files=["data/BrightPath_Engagement_Survey.xlsx"],
        prerequisites=[
            "Lab 7 completed: you can verify an insight and label it OBSERVED or HYPOTHESIS.",
            "This lab folder's data/BrightPath_Engagement_Survey.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Engagement_Survey.xlsx")
                + " Responses are anonymous: do not try to identify anyone, and do not report a group with fewer than 5 responses.",
                "Table : tblSurvey (160 responses; Q1-Q6 scored 1-5; eNPS_Score 0-10; Comment)\n"
                "eNPS  = % promoters (9-10) - % detractors (0-6)\n"
                "Templates: Recommendations | Executive_Summary",
            ),
            (
                "Choose Chat only and ask for eNPS and the question averages. Then ask Copilot (Allow editing) to add the check "
                "formulas on a new Checks sheet.",
                "CHAT ONLY:\n"
                "Using tblSurvey, calculate the number of promoters (9-10), passives (7-8) and detractors (0-6) on\n"
                "eNPS_Score, and the eNPS. Then give the average of Q1 to Q6, lowest to highest.\n\n"
                "CHECK (Allow editing):\n"
                "Create a sheet named Checks. Enter these formulas in B1:B4 with a label beside each in A1:A4,\n"
                "then tell me each result:\n"
                "=COUNTIFS(tblSurvey[eNPS_Score],\">=9\")\n"
                "=COUNTIFS(tblSurvey[eNPS_Score],\"<=6\")\n"
                "=(COUNTIFS(tblSurvey[eNPS_Score],\">=9\")-COUNTIFS(tblSurvey[eNPS_Score],\"<=6\"))/ROWS(tblSurvey[Response_ID])*100\n"
                "=AVERAGE(tblSurvey[Q3_Workload_Balance])",
            ),
            (
                "In Chat only, ask for eNPS by department with response counts. Mark any department with fewer than 10 responses as "
                "'interpret with caution' and never report one with fewer than 5.",
                "Show eNPS and number of responses for each Department in tblSurvey, sorted lowest eNPS first.",
            ),
            (
                "Choose Allow editing and ask Copilot to classify comments into a fixed list. If your Excel has the COPILOT "
                "function you may ask Copilot to use the formula option instead. Keep the original Comment column unchanged.",
                "OPTION A (Copilot pane):\n"
                "Add two columns to tblSurvey. Theme: classify each Comment into exactly one of Workload, Career Growth,\n"
                "Manager Support, Pay & Benefits, Recognition, Tools & Systems, Flexibility, or Unclear. Sentiment:\n"
                "Positive, Negative or Mixed. Leave both blank where Comment is blank. Do not change Comment.\n\n"
                "OPTION B (COPILOT function, where available):\n"
                "In tblSurvey, add a Theme column, enter this formula in its first cell and fill it down the column:\n"
                "=COPILOT(\"Classify this staff comment into exactly one theme from the list. Reply with the theme only.\",\n"
                "         [@Comment], \"List:\", \"Workload, Career Growth, Manager Support, Pay & Benefits, Recognition,\n"
                "         Tools & Systems, Flexibility, Unclear\")",
            ),
            (
                "Spot-check: ask Copilot to show 10 classified comments, read them and count any you disagree with. If more than "
                "1 in 10 is wrong, refine the prompt (for example, add an example per theme) and rerun. Then ask for theme counts, "
                "the COUNTIF check, and ask Copilot to record the error rate on Prompt_Log.",
                "CHAT ONLY:\n"
                "Show me 10 random rows of tblSurvey with Comment, Theme and Sentiment only.\n\n"
                "CHAT ONLY:\n"
                "Count comments by Theme and by Sentiment in tblSurvey, highest first. For the top theme, summarise\n"
                "what staff are saying in three bullet points without quoting anything that could identify a person.\n\n"
                "CHECK (Allow editing):\n"
                "On the Checks sheet, add this formula and tell me if it matches your Workload count:\n"
                "=COUNTIF(tblSurvey[Theme],\"Workload\")\n\n"
                "ALLOW EDITING:\n"
                "Add a row to Prompt_Log: \"Theme spot-check: <number> of 10 disagreed\".",
            ),
            (
                "Choose Allow editing to fill the Recommendations sheet: ask Copilot to draft three rows from your verified "
                "findings, then tell Copilot what to change so the owners, targets and dates are realistic. Every Evidence cell "
                "must cite numbers you verified.",
                "Using only these verified findings: <paste eNPS, lowest question, top theme, department results>,\n"
                "draft three recommendations in the Recommendations sheet columns: Finding (OBSERVED), Evidence,\n"
                "Recommendation (a team, process or policy action - not about any individual), Owner (role),\n"
                "Priority, Success KPI and target, Review date (by Dec 2026).\n\n"
                "FOLLOW-UP: In the Recommendations sheet, change <row and column> to <new owner, target or date>.",
            ),
            (
                "Ask Copilot to draft the executive summary on Executive_Summary. Check every number against your Checks sheet, "
                "tell Copilot what wording to change, and ask it to add your decision: Share, Revise or Hold. Then ask Copilot to "
                "log all prompts. Optionally ask Copilot in Word or PowerPoint to turn the summary into a memo or two slides.",
                "Draft a 150-word executive summary for BrightPath's leadership from the Recommendations sheet and the\n"
                "verified results. Include eNPS, the lowest-scoring question, the top comment theme and the three\n"
                "recommendations with owners. Use plain English, no jargon, and do not claim causes.\n\n"
                "FOLLOW-UP: Add a final line to Executive_Summary: \"Decision: <Share / Revise / Hold> - <one-line reason>\".\n\n"
                + _LOG,
            ),
        ],
        sample_prompts=[
            ("Score", "Calculate eNPS from eNPS_Score in tblSurvey: promoters 9-10, passives 7-8, detractors 0-6. Show the counts."),
            ("Averages", "Give the average of each question Q1 to Q6, lowest to highest, with the number of responses."),
            ("Groups", "Show eNPS and number of responses by Department, sorted lowest first."),
            ("Themes", "Classify each Comment into exactly one theme from my list (or Unclear) in a new Theme column. Do not change Comment."),
            ("Sentiment", "Add a Sentiment column: Positive, Negative or Mixed for each non-blank Comment."),
            ("COPILOT function", "=COPILOT(\"Classify the sentiment of this comment as Positive, Negative or Mixed\", [@Comment])"),
            ("Summarise text", "Summarise what staff say about Workload in three bullet points without quoting anything identifying."),
            ("Recommend", "From these verified findings, draft three recommendations with owner, priority, success KPI and review date."),
            ("Executive summary", "Draft a 150-word executive summary for leadership. Do not claim causes; cite the numbers."),
        ],
        test=(
            f"Promoters {E8['promoters']}, passives {E8['passives']}, detractors {E8['detractors']} of {E8['rows']}; eNPS = "
            f"{E8['enps']}. Lowest question {_q(E8['lowest_q'])} ({E8['q_avg'][E8['lowest_q']]:.2f}); highest "
            f"{_q(E8['highest_q'])} ({E8['q_avg'][E8['highest_q']]:.2f}). {E8['low_dept']} has the lowest department eNPS "
            f"({E8['by_dept'][E8['low_dept']]}) from only {E8['dept_n'][E8['low_dept']]} responses, so it is flagged 'interpret "
            f"with caution'. {E8['comments']} responses have a comment; after classification {E8['top_theme']} is the most "
            f"frequent theme (about {E8['theme_counts'][E8['top_theme']]} comments - AI labels can vary by a few, which is why "
            "the spot-check error rate is recorded). The Recommendations sheet has three rows, each with verified evidence, an "
            "owner role, a KPI target and a review date, and Executive_Summary ends with a Share, Revise or Hold decision."
        ),
        checkpoint="This is the final lab. Keep all eight workbooks in work-c34 as your personal reference set.",
        troubleshooting=[
            (
                "Copilot invents new themes outside the list.",
                "Repeat the exact list, add 'reply with one theme from the list only', and ask it to use Unclear when unsure.",
            ),
            (
                "The COPILOT function returns #NAME? or is missing.",
                "It is not available on every licence or channel yet. Use Option A in the Copilot pane instead.",
            ),
            (
                "Copilot's eNPS uses 7-10 as promoters or averages the scores.",
                "Restate the rule: promoters 9-10, detractors 0-6, eNPS = % promoters - % detractors. Trust your COUNTIFS.",
            ),
            (
                "The summary quotes a comment that could identify someone.",
                "Remove it and ask Copilot to paraphrase themes without quotes. Anonymity comes before vividness.",
            ),
        ],
        challenge=(
            "Combine Lab 7 and Lab 8: ask Copilot whether the departments with the highest attrition also have the lowest "
            "Workload Balance scores, then write one recommendation that uses both sources - labelled as a HYPOTHESIS."
        ),
        reflection="Which recommendation would you be confident defending to the HR Director, and which number makes it defensible?",
    ),
]
