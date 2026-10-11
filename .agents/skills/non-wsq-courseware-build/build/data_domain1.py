"""Topic 1 labs for C34: Use Copilot to Analyse Data (BrightPath recruitment data)."""

from mock_data import EXPECTED

E1 = EXPECTED[1]
E2 = EXPECTED[2]


def _money(v):
    return f"S${v:,.0f}" if float(v).is_integer() else f"S${v:,.2f}"


_SETUP_STEP = (
    "Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload "
    "{file} from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title "
    "bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the "
    "About_This_Data sheet before you prompt."
)

_LOG = (
    "LOG (Allow editing):\n"
    "Add one row to Prompt_Log for each prompt I ran in this lab: Mode, Prompt, What Copilot did,\n"
    "How it was verified, and Decision = Keep. I will tell you any row to change to Revise or Undo."
)

DOMAIN1 = [
    # ------------------------------------------------------------------ LAB 1
    dict(
        num=1,
        topic=1,
        title="Explore and Question Recruitment Pipeline Data",
        duration=35,
        objective=(
            "LO1: confirm Copilot access, prepare a named table, and use Chat and Edit modes to question, check, "
            "highlight, sort and filter recruitment pipeline data"
        ),
        goal=(
            "Turn a raw candidate-pipeline export into a named Excel table, find its data-quality issues with Copilot, "
            "and answer pipeline questions with checked numbers."
        ),
        workflow=["Save to OneDrive", "Create the table", "Describe the data", "Check quality", "Ask, highlight, filter"],
        desc=(
            "You are a talent analyst at BrightPath Talent Solutions, a fictional Singapore recruitment and staffing firm. "
            "You will put the H1 2026 candidate-pipeline export into a named table, ask Copilot (Chat only) to describe it, "
            "use Copilot to find duplicate IDs and inconsistent Source labels, then answer pipeline questions and build a "
            "highlighted, sorted and filtered review view. Every number is checked with a formula."
        ),
        build=(
            "BrightPath_Candidate_Pipeline.xlsx in OneDrive with table tblApplications, a Source_Clean column, duplicate "
            "IDs highlighted, Hired rows highlighted, a verified source-by-hires answer and a filled Prompt_Log"
        ),
        services="Excel for Microsoft 365 or Excel for the web · Copilot (Chat only, Allow editing) · BrightPath_Candidate_Pipeline.xlsx",
        data_files=["data/BrightPath_Candidate_Pipeline.xlsx"],
        prerequisites=[
            "A Microsoft 365 account with Copilot in Excel enabled, signed in to Excel or Excel for the web.",
            "A OneDrive or SharePoint folder named work-c34 that you can upload to.",
            "This lab folder downloaded, including data/BrightPath_Candidate_Pipeline.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Candidate_Pipeline.xlsx"),
                "Required state:\n"
                "File      : OneDrive/work-c34/BrightPath_Candidate_Pipeline.xlsx\n"
                "AutoSave  : On\n"
                "Calculation Options : Automatic\n"
                "Sheets    : About_This_Data | Applications | Prompt_Log",
            ),
            (
                "Select the Copilot button and choose Allow editing. Ask Copilot to turn the Applications data into a named "
                f"table. Its reply must confirm the table name tblApplications, 13 columns and {E1['rows']} data rows. If "
                "Copilot is unavailable, use the Troubleshooting table before continuing.",
                "Format the data on the Applications sheet (from A1 to the last row and column) as an Excel table\n"
                "with headers. Name the table tblApplications. Do not change, sort or delete any values.\n"
                "Then tell me the table name, the number of columns and the number of data rows.",
            ),
            (
                "Switch to Chat only (no workbook changes). Paste the prompt below. Copilot's row count must match the "
                f"{E1['rows']} rows it reported in Step 2. You will log every prompt in Step 7.",
                "GOAL: Describe this candidate pipeline before I analyse it.\n"
                "CONTEXT: BrightPath recruitment desk, H1 2026. One row should be one application.\n"
                "DATA: Use only tblApplications.\n"
                "CONSTRAINTS: Do not change the workbook.\n"
                "OUTPUT: What a row represents, each column with its data type, the number of rows,\n"
                "        the Applied_Date range and the distinct values of Stage and Source.\n"
                "REVIEW: Tell me how you counted the rows.",
            ),
            (
                "Still in Chat only, ask Copilot to check data quality. It should report 2 repeated Application_IDs and Source "
                "values that are not in the canonical list on About_This_Data. Then switch to Allow editing (Edit) and ask Copilot "
                "to highlight the duplicate IDs and add a Source_Clean column. Check that no row was deleted - the table must still "
                f"have {E1['rows']} rows.",
                "CHAT ONLY:\n"
                "Check tblApplications for data-quality issues. List any Application_ID that appears more than once,\n"
                "and any Source value that is not exactly one of: LinkedIn, JobStreet, MyCareersFuture, Referral,\n"
                "BrightPath Website, Job Fair, Talent Database. Give the count of each issue. Do not change anything.\n\n"
                "ALLOW EDITING:\n"
                "1) Highlight Application_ID cells that appear more than once in tblApplications with an orange fill.\n"
                "2) Add a column Source_Clean to tblApplications that trims spaces from Source and changes 'Linked In'\n"
                "   to 'LinkedIn'. Do not delete or edit any existing values.",
            ),
            (
                "Ask Copilot a direct question about which sources produce hires. Expand or ask Copilot to summarise any code it "
                "shows. Then ask Copilot (Allow editing) to add the two check formulas to Prompt_Log and compare the results with "
                "its answer. Duplicate rows are Rejected applications, so they do not change the hire counts.",
                "PROMPT (Chat only):\n"
                "Using tblApplications, count applications and hires (Stage = \"Hired\") by Source_Clean.\n"
                "Show a table sorted by hires, highest first, with hire rate = hires / applications.\n\n"
                "CHECK (Allow editing):\n"
                "In empty cells on Prompt_Log, add these check formulas with a label beside each, then tell me the results:\n"
                "=COUNTIFS(tblApplications[Stage],\"Hired\")\n"
                "=COUNTIFS(tblApplications[Source_Clean],\"LinkedIn\",tblApplications[Stage],\"Hired\")",
            ),
            (
                "Switch to Allow editing and build a review view, one change at a time. After each change, look at the sheet and "
                "keep it or select Undo in the Copilot pane. Finally ask Copilot how many rows the filter shows, then ask it to "
                f"clear the filter so all {E1['rows']} rows return.",
                "1) Highlight entire rows of tblApplications where Stage is \"Hired\" with a light green fill.\n"
                "2) Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline\n"
                "   from lowest to highest.\n"
                "3) Filter tblApplications to show only Client_Sector = \"Technology\" and Stage = \"Interview\".\n"
                "4) How many rows of tblApplications are visible with this filter?\n"
                f"5) Clear all filters on tblApplications and confirm all {E1['rows']} rows are visible.",
            ),
            (
                "Ask Copilot to complete the Prompt_Log for every prompt you ran: mode, prompt, what Copilot did, how it was "
                "verified and your decision (Keep, Revise or Undo). AutoSave keeps the cloud copy current.",
                _LOG,
            ),
        ],
        sample_prompts=[
            ("Describe", "Describe tblApplications: what one row represents, each column's data type, the row count and the Applied_Date range. Do not change the workbook."),
            ("Data quality", "List any Application_ID that appears more than once in tblApplications and any Source value that is not in the canonical list on About_This_Data."),
            ("Clean helper column", "Add a column Source_Clean that trims spaces from Source and replaces 'Linked In' with 'LinkedIn'. Do not change the original Source column."),
            ("Question", "Using tblApplications, count applications and hires by Source_Clean and show the hire rate. Sort by hires, highest first."),
            ("Question", "Which Client_Sector received the most applications, and how many? Show the count for every sector."),
            ("Highlight", "Highlight rows where Stage is Hired in light green."),
            ("Sort", "Sort tblApplications by Interview_Score from highest to lowest, then by Days_in_Pipeline from lowest to highest."),
            ("Filter", "Filter to show only Technology applications at the Interview stage."),
            ("Insight", "What's interesting about this pipeline data? Give three observations, each with the numbers behind it."),
        ],
        test=(
            f"tblApplications exists with {E1['rows']} data rows and =COUNTA(UNIQUE(tblApplications[Application_ID])) returns "
            f"{E1['unique']}. The repeated IDs are {E1['dup_ids'][0]} and {E1['dup_ids'][1]} (both Rejected) and are orange. "
            f"Copilot found {E1['variants_linkedin']} 'Linked In' and {E1['variants_jobstreet']} 'JobStreet ' (trailing space) "
            f"values; =COUNTIF(tblApplications[Source_Clean],\"LinkedIn\") returns {E1['linkedin_clean_all']}. Hires total "
            f"{E1['hired']} and {E1['top_hire_source']} has the most hires ({E1['top_hire_source_n']}). The Technology + Interview "
            f"filter shows {E1['tech_interview_all']} rows. Technology received the most applications ({E1['top_sector_n']} unique). "
            f"After clearing the filter all {E1['rows']} rows are visible, and Prompt_Log has at least 5 entries."
        ),
        checkpoint=(
            "Keep this workbook in work-c34. Lab 2 uses a separate requisitions workbook, so a learner who is behind can "
            "start Lab 2 from its own data folder."
        ),
        troubleshooting=[
            (
                "The Copilot button is missing or greyed out.",
                "Save the file as .xlsx in OneDrive or SharePoint, turn AutoSave On, close and reopen it. Check your licence in "
                "your Microsoft 365 account and ask IT if Copilot is disabled for your tenant. Pair with a neighbour meanwhile.",
            ),
            (
                "'Unsupported file state' message.",
                "If the file is checked out in SharePoint, check it in and reopen it, or open it in Excel for the web. Strict "
                "Open XML files are not supported - save as a normal .xlsx.",
            ),
            (
                "Copilot answers about the whole sheet instead of the table.",
                "Name the table (tblApplications) and the exact columns in the prompt, and click inside the table before asking.",
            ),
            (
                "Copilot's count differs from your formula.",
                "Ask Copilot how it counted (rows vs unique IDs, which Stage values). Trust the formula, record the difference "
                "and refine the prompt.",
            ),
        ],
        challenge=(
            "Ask Copilot to add a Pipeline_Status column that groups Stage into Active (Applied, Screening, Interview, Offer), "
            "Hired and Closed (Rejected, Withdrawn). Verify the three counts add up to the row count."
        ),
        reflection="Which data-quality issue would most distort a source-effectiveness report if you had not found it, and why?",
    ),
    # ------------------------------------------------------------------ LAB 2
    dict(
        num=2,
        topic=1,
        title="Generate and Verify Formula Columns and Lookups for Job Requisitions",
        duration=35,
        objective=(
            "LO1: use Copilot to generate lookup and calculated columns from plain-English rules, explain them, "
            "and reconcile them to independent control formulas"
        ),
        goal=(
            "Add recruiter names, days open, fill rate, estimated placement fee and an ageing flag to a requisitions table, "
            "then prove each column is right."
        ),
        workflow=["Lookup", "Days open", "Fill rate & fee", "Ageing flag", "Reconcile"],
        desc=(
            "BrightPath's permanent-recruitment desk tracks 60 client job orders. You will ask Copilot to add an XLOOKUP for "
            "the recruiter's name and four calculated columns written from business rules, ask it to explain each formula, "
            "hand-check one requisition and reconcile column totals to independent formulas on the Control sheet. You will then "
            "ask Copilot which recruiter generated the most estimated fees and verify the answer."
        ),
        build=(
            "BrightPath_Job_Requisitions.xlsx with Recruiter_Name, Days_Open, Fill_Rate, Est_Fee and Ageing_Flag columns in "
            "tblRequisitions, overdue requisitions highlighted, and a Control sheet that reconciles the totals"
        ),
        services="Excel · Copilot (Allow editing, Chat only) · XLOOKUP · structured references · BrightPath_Job_Requisitions.xlsx",
        data_files=["data/BrightPath_Job_Requisitions.xlsx"],
        prerequisites=[
            "Lab 1 completed, or Copilot access confirmed on a OneDrive file with AutoSave On.",
            "This lab folder's data/BrightPath_Job_Requisitions.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Job_Requisitions.xlsx")
                + " Note the fee rule: Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, and the report date in Control!B1.",
                "Tables : tblRequisitions (60 rows) | tblRecruiters (6 rows)\n"
                "Control!B1 Report_Date = 2026-06-30",
            ),
            (
                "Choose Allow editing. Ask Copilot to add the lookup column, then ask it to explain the formula and show one row "
                "next to the matching row in tblRecruiters.",
                "In tblRequisitions, add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters\n"
                "using Recruiter_ID. Use XLOOKUP and return \"Not found\" if there is no match.\n\n"
                "Then: Explain the Recruiter_Name formula, and show one row's Recruiter_ID and Recruiter_Name next to\n"
                "the matching row in tblRecruiters.",
            ),
            (
                "Ask for Days_Open with an explicit rule for open requisitions. Read the explanation, then ask Copilot to show an "
                "open row (blank Close_Date) and a closed row so you can confirm the rule.",
                "Add a column Days_Open to tblRequisitions. If Close_Date is blank, use Control!B1 minus Open_Date;\n"
                "otherwise use Close_Date minus Open_Date. Show whole numbers.\n\n"
                "FOLLOW-UP: Show one open requisition (blank Close_Date) and one closed requisition with Open_Date,\n"
                "Close_Date and Days_Open, and the subtraction behind each.",
            ),
            (
                "Ask for Fill_Rate and Est_Fee, one prompt each. Format Fill_Rate as a percentage and Est_Fee as S$ with no "
                f"decimals. Then ask Copilot to show the calculation for {E2['sample_id']}: Filled {E2['sample_detail'][0]} x S${E2['sample_detail'][1]:,} x 12 x "
                f"{E2['sample_detail'][2]:.0%} = {_money(E2['sample_fee'])}.",
                "Add a column Fill_Rate = Filled divided by Openings, formatted as a percentage.\n\n"
                "Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals.\n"
                "Explain the formula.\n\n"
                f"FOLLOW-UP: Show the Est_Fee calculation for {E2['sample_id']} step by step.",
            ),
            (
                "Ask for the ageing flag and a highlight, then ask Copilot how many requisitions are flagged.",
                "Add a column Ageing_Flag that shows \"Overdue\" when Status is \"Open\" and Days_Open is greater than 45,\n"
                "otherwise leave it blank. Highlight rows where Ageing_Flag is \"Overdue\" with a light amber fill.\n\n"
                "FOLLOW-UP: How many rows have Ageing_Flag = \"Overdue\"?",
            ),
            (
                "Ask Copilot to build independent controls on the Control sheet with the exact formulas below. The independent "
                "formulas recalculate from the raw columns, so they do not rely on Copilot's new columns. Each pair must match.",
                "Enter these labels and formulas on the Control sheet exactly as written. Then tell me each result\n"
                "and whether each pair (B3 and B4, B6 and B7) matches:\n"
                "A3 Total Est_Fee (column)       B3 =SUM(tblRequisitions[Est_Fee])\n"
                "A4 Total fee (independent)      B4 =SUMPRODUCT(tblRequisitions[Filled],tblRequisitions[Monthly_Salary_Budget],tblRequisitions[Fee_Rate])*12\n"
                "A5 Avg days to fill (Filled)    B5 =AVERAGEIFS(tblRequisitions[Days_Open],tblRequisitions[Status],\"Filled\")\n"
                "A6 Overdue (independent)        B6 =COUNTIFS(tblRequisitions[Status],\"Open\",tblRequisitions[Days_Open],\">45\")\n"
                "A7 Overdue flags (column)       B7 =COUNTIF(tblRequisitions[Ageing_Flag],\"Overdue\")",
            ),
            (
                "In Chat only, ask which recruiter generated the most estimated fees. Then ask Copilot (Allow editing) to add the "
                "SUMIFS check to the Control sheet and to log all prompts on Prompt_Log.",
                "CHAT ONLY:\n"
                "Using tblRequisitions, which Recruiter_Name has the highest total Est_Fee? Show the total for every recruiter.\n\n"
                "CHECK (Allow editing):\n"
                "On the Control sheet, add this check for the top recruiter and tell me if it matches your answer:\n"
                "=SUMIFS(tblRequisitions[Est_Fee],tblRequisitions[Recruiter_Name],\"<name>\")\n\n" + _LOG,
            ),
        ],
        sample_prompts=[
            ("Lookup", "Add a column Recruiter_Name that looks up Recruiter_Name from tblRecruiters using Recruiter_ID with XLOOKUP. Return \"Not found\" if missing."),
            ("Duration", "Add a column Days_Open: if Close_Date is blank use Control!B1 minus Open_Date, otherwise Close_Date minus Open_Date."),
            ("Ratio", "Add a column Fill_Rate = Filled / Openings formatted as a percentage."),
            ("Business rule", "Add a column Est_Fee = Filled x Monthly_Salary_Budget x 12 x Fee_Rate, formatted as S$ with no decimals."),
            ("Flag", "Add a column Ageing_Flag that shows \"Overdue\" when Status is Open and Days_Open > 45, otherwise blank."),
            ("Explain", "Explain the formula in the selected cell step by step, and tell me how blanks are handled."),
            ("Summary row", "Add a total row to tblRequisitions that sums Openings, Filled and Est_Fee and averages Days_Open."),
            ("Question", "Which recruiter has the highest total Est_Fee? Show every recruiter's total, highest first."),
        ],
        test=(
            f"Control!B3 and B4 both show {_money(E2['fee_total'])}. B5 shows {E2['avg_days_filled']} days "
            f"({E2['filled_n']} Filled requisitions). B6 and B7 both show {E2['overdue_n']}. {E2['sample_id']} shows Est_Fee "
            f"{_money(E2['sample_fee'])} and Days_Open {E2['sample_days']}. Every Recruiter_Name resolves (no \"Not found\"). "
            f"The top recruiter by estimated fees is {E2['top_rec']} with {_money(E2['top_rec_fee'])}."
        ),
        checkpoint=(
            "Keep the workbook in work-c34. Topic 2 starts from a new metrics workbook in the Lab 3 folder."
        ),
        troubleshooting=[
            (
                "Days_Open is negative or huge for open requisitions.",
                "The formula used today's date or a blank as zero. Ask Copilot to use Control!B1 explicitly, or edit the formula.",
            ),
            (
                "Est_Fee is 12 times too small or too large.",
                "Re-read the rule: fee is on first-year salary (monthly x 12). Ask Copilot to explain the formula and correct it.",
            ),
            (
                "The new column is outside the table.",
                "Click inside tblRequisitions before prompting and name the table in the prompt; or drag the table's resize handle.",
            ),
            (
                "Copilot cannot add a formula column.",
                "Confirm Calculation Options is Automatic and the data is an Excel table; try Excel for the web if desktop fails.",
            ),
        ],
        challenge=(
            "Ask Copilot to add a Fee_Achieved column that is Est_Fee for Filled requisitions only, then build a small "
            "summary of Fee_Achieved by Sector with SUMIFS. Check that the sector totals add up to the grand total."
        ),
        reflection="Which of today's formulas would have looked correct but been wrong without a hand-check, and how did you catch it?",
    ),
]
