"""Lab copy, exact prompts, expected results and the reusable theme skill."""

from bog_review import LAB as BOG_LAB
from datasets import CHANNELS, CHECKPOINT, FIRST, FRESH, headline, records, select, totals

MUSE_URL = "https://muse.ai/"
VIDEO_URL = "https://www.youtube.com/watch?v=YgcEesCDIkw"
EXCEL_URL = "https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel"
SKILLS_URL = "https://support.microsoft.com/en-gb/excel/copilot/copilot-in-excel-skills"
SLICERS_URL = "https://support.microsoft.com/en-us/excel/get-started/use-slicers-to-filter-data"
SKILL_NAME = "datin-dark-sales-dashboard"
SKILL_PATH = f"skills/{SKILL_NAME}/SKILL.md"

MUSE_CONNECT = (
    "Confirm whether you can read Gmail and Google Calendar in this demo account. Tell me which connectors "
    "are available and what approval you need before an action. Do not read anything yet or change settings."
)
MUSE_TRACK = (
    "Prepare a read-only month-end brief for the fictional Tema distributor. In the approved demo Gmail "
    "account find only subjects beginning DATIN DEMO. In the demo Google Calendar check only events "
    "beginning DATIN DEMO for the next seven days, in Accra time (GMT). Match the email and event by "
    "ticket ID. Rank S-104 (sales cut-off), S-109 (supplier duplicate) and S-112 (credit note) by the "
    "recorded deadline. Return the source, deadline, evidence to check and approval owner. Mark missing "
    "details. Do not send or move email, change events, post entries or make payments. Tell me whether "
    "you can repeat this as a read-only scheduled brief. Ask before creating any schedule."
)
MUSE = (
    "You are helping the finance team of a fictional wholesale distributor in Tema, Ghana, at month-end. "
    "Work only from these mock notes. Today is Tuesday. S-104: reconcile a short rice delivery against a "
    "GHS 4,800 sales invoice before the cut-off on Wednesday at 2 pm. S-109: check a possible duplicate "
    "GHS 3,250 supplier invoice before the payment review on Thursday at 10 am. S-112: check a request for "
    "a GHS 620 credit note for returned detergent, with no deadline given. Rank the three items by explicit "
    "deadline. Return ID, reason, evidence to check and one draft recommendation per item. Mark missing "
    "information and name the accounting record or tool needed. Identify who should approve a posting, "
    "payment or customer credit. Draft only. Leave accounts, connectors, records and payments untouched."
)
CREATE = (
    "Create a professional, interactive sales dashboard using the data in this workbook. "
    "Include key sales metrics, charts showing sales trends and performance, and interactive filters "
    "so I can explore the data. Put everything on a new worksheet."
)
THEME = (
    "Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, "
    "and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully "
    "functional. Make sure all labels and charts remain easy to read."
)
REPAIR = (
    "On the dashboard sheet, use only tblDatinSales. Show Net Sales, Gross Profit, Gross Margin "
    "(total Gross Profit divided by total Net Sales) and Order Count (count Order_ID). Format money in GHS. "
    "Show a monthly net revenue trend, regional revenue bars, and doughnut charts for Product and Channel mix. Connect Region, Channel and Product controls to every KPI and chart. "
    "Keep the Sales Data table unchanged and tell me which controls you implemented."
)
CONVERT = (
    f"@skill-builder Turn the theme prompt below into an Excel skill named {SKILL_NAME}. "
    "For an existing interactive sales dashboard, apply the theme while preserving its layout, source data, "
    "formulas, charts and working filters. If the request is to create a dashboard and none exists, first build "
    "one from the workbook sales table with GHS Net Sales, Gross Profit, Gross Margin and Order Count, monthly "
    "trend and regional charts, doughnut charts for product and channel mix, and Region, Channel and Product filters. Apply the theme and verify the filters. "
    "Show the proposed skill for approval. Theme prompt: " + THEME
)
SHORT = "Create an interactive Datin sales dashboard."
EXPLICIT = f"@{SKILL_NAME} {SHORT}"

SKILL = f"""---
name: {SKILL_NAME}
description: Create an interactive Datin sales dashboard in Excel or restyle an existing sales dashboard with the Datin navy, gold and white theme. Use for requests to create an interactive Datin sales dashboard or apply its dark theme.
---

# Datin dark sales dashboard

## Choose the operation

- If the user asks to restyle an existing dashboard, identify the intended dashboard sheet. Keep its layout, chart objects and series, KPI formulas, source table, filter connections and current filter selections.
- If the user asks to create the dashboard and a dashboard is absent, create one new worksheet named Sales Dashboard. Prefer tblDatinSales. Inspect the current workbook and confirm the source and mapping if several plausible tables exist.
- If a dashboard already exists and the request to create another is ambiguous, ask before replacing the existing sheet.

## Build when needed

1. Use one row per order, keyed by Order_ID. Use Net_Sales_GHS and Gross_Profit_GHS for totals. Count Order_ID. Compute Gross Margin as total Gross Profit divided by total Net Sales. Show a dash for zero sales. Format money with the GHS code and margin as a percentage.
2. Add a monthly net revenue line chart using Month_Start, a regional revenue bar chart, a doughnut chart by Product and another doughnut chart by Channel. Use net sales after discounts as net revenue on all four charts.
3. Add working Region, Channel and Product controls, such as connected slicers or drop-down selectors driving formulas. Use the same filters for all charts and all four KPIs. Start with All. Keep source values and table structure intact. Keep chart support calculations on the dashboard.
4. If the Excel session cannot create working controls, identify the unfinished control and explain the next manual step. Report completion after filter behavior is verified.

## Theme from the Datin redesign prompt

- Use navy #0B1730 for the dashboard canvas and #152642 for cards and chart panels.
- Use gold #E9BE64 for key accents and primary series. Use white #FFFFFF for titles, axes, labels and important text. Make gridlines subtle and keep filter states distinguishable with text and contrast.
- For a restyle, edit only the dashboard’s presentation. Preserve the existing chart series, units, formulas, positions and controls. Keep currency, labels and legends legible at normal zoom.

## Verify

- Compare unfiltered sales, gross profit, margin and order count with the source totals.
- Filter Greater Accra and Wholesale, then add Rice when those options exist. Check that all KPIs and charts respond. Restore the starting selections after testing.
- For a restyle, confirm the same selection displays the same numbers before and after. Report what changed and identify any control that still needs manual verification.
"""

def checks(period):
    rows = list(records(period))
    return [
        ("All regions and channels", headline(totals(rows))),
        ("Greater Accra plus Wholesale", headline(totals(select(rows, region="Greater Accra", channel=CHANNELS[0])))),
        ("Greater Accra plus Wholesale plus Rice", headline(totals(select(rows, region="Greater Accra", channel=CHANNELS[0], product="Rice")))),
    ]

LABS = [
    dict(
        num=1, topic=1, slug="meta-muse", title="Prioritise Month-End Exceptions with Meta Muse",
        theme="Meta Muse", duration=20, objective="LO1 · Explain agent behaviour and human approval",
        services="Meta Muse in a browser",
        goal="Use Meta Muse to track demo email and calendar deadlines and prepare a traceable month-end brief.",
        desc="Read a demo Gmail inbox and Google Calendar to prepare a finance exception queue for a Tema distributor. Check cut-off, duplicate payment and credit-note evidence.",
        build="A three-row month-end exception brief showing evidence to check and approval owners.",
        prerequisites=[
            "Use muse.ai and a trainer-approved demo Gmail account and Google Calendar. The Outlook training accounts are for Excel.",
            "Before class, the trainer seeds three demo email subjects prefixed DATIN DEMO with ticket IDs S-104, S-109 and S-112. Put the mock amounts and exception notes from the fallback prompt in the messages. Add matching calendar events for S-104 and S-109 within seven days, S-104 first, in Accra time (GMT). Leave S-112 unscheduled.",
            "If connector access is unavailable, watch the trainer’s connected demonstration or use the inline mock prompt.",
        ],
        workflow=["Connect demo sources", "Track deadlines", "Verify the draft", "Confirm monitoring scope"],
        data_files=[],
        steps=[
            ("Open muse.ai and start a new chat. Under Settings > Connectors, connect only the approved demo Gmail and Google Calendar. Use read-only permissions where offered. Run the connector check.", MUSE_CONNECT),
            ("Ask for a read-only email and calendar brief. Confirm the returned sources and dates in the demo account.", MUSE_TRACK),
            ("If the connectors are unavailable, use this mock-notes variant. Check the three ticket IDs and verify which case has no deadline.", MUSE),
            ("If scheduled tracking is supported, review the proposed demo-only scope and frequency. The trainer may approve one read-only scheduled brief for the demo, then stop it at the end of class. Compare the result with the actual demo sources.", ""),
        ],
        sample_prompts=[
            ("Check email and calendar connectors", MUSE_CONNECT),
            ("Track month-end email and calendar", MUSE_TRACK),
            ("Fallback · Draft from inline mock notes", MUSE),
        ],
        checks=[
            ("Order and evidence", "S-104 ranks before S-109. S-112 has no stated deadline. All three IDs appear. In the connected version, S-104 and S-109 have matched inbox and calendar records."),
            ("Tool and approval", "Check proof of delivery, the AP ledger and the returns record. The finance manager approves postings, payments or credits."),
        ],
        checkpoint="Keep your close brief. Lab 2 moves from individual exceptions to Tema’s H1 sales and margin reporting. Use the independent mock workbook and reconcile it to its Data Guide.",
        troubleshooting=[
            ("Muse unavailable", "Follow the trainer demonstration. Use the same mock brief to check the reasoning together."),
            ("Muse invents a due date", "Ask it to cite the notes and label the missing date as unknown."),
        ],
        challenge="Revise the prompt to define what counts as urgent and compare the resulting priorities.",
        reflection="Which part of this task could run on a schedule, and which decision should a person own?",
        links=[("Open Muse", MUSE_URL)],
    ),
    dict(
        num=2, topic=2, slug="interactive-sales-dashboard", title="Create and Restyle an Interactive Sales Dashboard",
        theme="Interactive Excel dashboard", duration=60, objective="LO2 · Build and verify an interactive dashboard",
        services="Excel for Microsoft 365 with Copilot",
        goal="Build an interactive GHS dashboard, then apply the dark theme while preserving every result and control.",
        desc="The trainer runs Prompt 1 to build the dashboard. Learners run Prompt 2 as the hands-on activity and compare the numbers and filter behavior before and after.",
        build="A new Sales Dashboard worksheet showing KPI cards, a monthly trend, regional revenue bars, doughnut charts for product and channel mix, and three working filters.",
        prerequisites=["Excel for Microsoft 365 and a work account with Copilot editing enabled.",
                       f"Open data/{FIRST.filename} as a working copy in OneDrive. Turn AutoSave on."],
        workflow=["Build", "Check totals", "Apply dark theme", "Retest filters"],
        data_files=[FIRST.filename],
        downloads=[(f"examples/{CHECKPOINT}", "Finished dark dashboard example")],
        steps=[
            ("Open the practice workbook in Excel. On the Sales Data sheet, click inside tblDatinSales. Open Copilot, select Plan if available to review the proposed changes, then allow editing to apply them.", ""),
            ("Watch the trainer run Prompt 1, then run it in your own H1 copy. Check that Copilot creates a new dashboard sheet and links every chart and KPI to the data. Compare the structure with the reference video.", CREATE),
            ("Inspect the dashboard. It should show Net Sales, Gross Profit, Gross Margin and Order Count, plus a monthly trend, regional bars and doughnuts for product and channel mix. If any elements are missing, run the follow-up. Test All, Greater Accra plus Wholesale, and then add Rice. Record each result.", REPAIR),
            ("Learner activity. Save a pre-theme copy. Select the dashboard sheet, run Prompt 2 and inspect the result at 100% zoom.", THEME),
            ("Repeat all three filter states. Check that all four KPI values match the pre-theme checks, every chart responds, and the labels and controls remain readable. Clear the filters and save a working copy.", ""),
        ],
        sample_prompts=[("Prompt 1 · Create the dashboard", CREATE), ("Prompt 2 · Customize the theme", THEME), ("Follow-up · Complete or repair", REPAIR)],
        checks=checks(FIRST) + [("Dark theme", "Navy panels, gold accents and readable white labels. The original layout, charts and controls still work.")],
        checkpoint=f"Use examples/{CHECKPOINT} to check the finished look and test the Region, Channel and Product drop-downs. For the theme activity, use your Prompt 1 dashboard or a trainer-provided pre-theme copy. Keep the theme and verification steps. In Lab 3 you will save them as a skill and test them with Tema’s fresh H2 orders.",
        troubleshooting=[
            ("Figures do not reconcile", "Check the table source, GHS fields and shared filters. Use total profit divided by total sales for margin."),
            ("A chart ignores filters", "For PivotTables sharing a source, check Slicer > Report Connections. For drop-downs, trace each chart to a filtered summary."),
            ("Copilot cannot edit", "Check the account, OneDrive copy, AutoSave and edit mode. Use the working example for the filter demonstration."),
        ],
        challenge="Explain which product has the largest revenue share under Wholesale and what additional costs would be needed to assess net profit.",
        reflection="Which dashboard tests prove a cosmetic change preserved the analysis?",
        links=[("Kevin Stratvert’s dashboard walkthrough", VIDEO_URL), ("Microsoft’s slicer guide", SLICERS_URL)],
    ),
    dict(
        num=3, topic=3, slug="dashboard-skill", title="Install and Reuse the Dashboard Skill",
        theme="Excel custom skills", duration=50, objective="LO3 · Convert a prompt into an enabled skill and reuse it",
        services="Copilot in Excel and OneDrive",
        goal="Convert Prompt 2 into a reusable skill and use a short request to create a verified dashboard on fresh data.",
        desc="Save the theme and its preservation checks in SKILL.md. Add the create-if-missing steps so the same skill can apply the theme during a fresh dashboard build.",
        build="An enabled skill in your OneDrive skills folder and a new interactive dashboard based on July–December sales.",
        prerequisites=[f"Have Prompt 2 and data/{FRESH.filename} open as a fresh OneDrive working copy.",
                       "Use Excel with Manage plugins & skills available on your work account."],
        workflow=["Draft skill", "Upload to OneDrive", "Refresh and enable", "Recreate dashboard"],
        data_files=[FRESH.filename],
        extra_files=[(SKILL_PATH, SKILL, "ready-to-upload skill")],
        steps=[
            ("In Copilot in Excel, use @skill-builder when available to draft a skill from Prompt 2. Review the name, trigger, theme and preservation checks. Check that the draft defines both paths: restyle an existing dashboard or build one if a create request finds none. Choose Save skill once the proposed files are correct.", CONVERT),
            (f"For the manual upload route, use the reviewed file or the provided {SKILL_PATH}. In the Copilot pane choose … > Manage plugins & skills > Custom skills. Select Create OneDrive folder if needed, then Open skills folder. Upload the entire {SKILL_NAME} folder, with SKILL.md inside it. If you saved through @skill-builder, confirm the file is in OneDrive and skip the manual upload.", ""),
            (f"Return to Custom skills and select Refresh. Find {SKILL_NAME} under Manage plugins & skills and turn it on. Open the fresh workbook. Start a new Copilot chat, allow editing and send the short prompt.", SHORT),
            ("Check the theme and controls. Record the skill name if Copilot shows which skill it used. For a controlled run, or if the short request misses the theme, start from a fresh workbook copy and select the skill from + > Choose plugins and skills or use this explicit request.", EXPLICIT),
            ("Check the July–December months and the three filter states below. Verify that the charts and KPI values use GHS, every chart follows the filters, and the source data is intact. Restore All and save the result.", ""),
        ],
        sample_prompts=[("Convert Prompt 2", CONVERT), ("Run the short prompt", SHORT), ("Invoke the skill explicitly", EXPLICIT)],
        slide_prompts=[
            ("Draft with @skill-builder", "Save Prompt 2 as a skill. Add the build-if-missing steps and the filter checks. Copy the full prompt from the Datin prompt page."),
            ("Run the short prompt", SHORT),
            ("Invoke the skill explicitly", EXPLICIT),
        ],
        checks=checks(FRESH) + [
            ("Skill discovery", f"{SKILL_NAME} appears after Refresh and is switched on."),
            ("Fresh dashboard", "July–December 2026 appears in the trend. All four charts react to filters and the theme is applied."),
        ],
        checkpoint="Keep the skill folder and the completed H2 working copy. Use the fresh starter again if you want to repeat the short-prompt test.",
        troubleshooting=[
            ("Skill does not appear", "Use Excel’s Open skills folder destination, match the folder and metadata names, check the .md extension and select Refresh."),
            ("Skill builder or skill menu is missing", "Use the manual upload if Custom skills is available. Otherwise review the supplied file while the trainer demonstrates on an enabled account."),
            ("Short prompt makes a light dashboard", "Choose the skill explicitly on a fresh workbook. Verify the new result and note that automatic selection may vary."),
        ],
        challenge="After verifying the dashboard, restyle it through the skill and check that all three filter states still return the same numbers.",
        reflection="What must a skill specify so a short request behaves consistently on a new workbook?",
        links=[("Microsoft’s Excel skills guide", SKILLS_URL)],
    ),
    BOG_LAB,
]
