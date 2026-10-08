"""Topic 2 labs for C34: Use Copilot to Plot Charts (BrightPath recruitment and placement data)."""

from mock_data import EXPECTED
from data_domain1 import _SETUP_STEP

E3 = EXPECTED[3]
E4 = EXPECTED[4]
_F = E3["funnel"]

DOMAIN2 = [
    # ------------------------------------------------------------------ LAB 3
    dict(
        num=3,
        topic=2,
        title="Create and Refine Recruitment Charts with Copilot",
        duration=45,
        objective=(
            "LO2: use Copilot to build PivotTables and trend, comparison, funnel and combo charts from prompts, refine "
            "them with follow-up prompts and tie each chart to a checked total"
        ),
        goal=(
            "Produce four decision-ready recruitment charts - hiring trend, hires by sector, the H1 2026 funnel and source "
            "effectiveness - each backed by a reconciled summary."
        ),
        workflow=["Trend line", "Sector bars", "Funnel", "Source combo", "Refine & check"],
        desc=(
            "BrightPath's leadership wants to see how hiring is moving and which sectors and sourcing channels perform best. "
            "You will ask Copilot to build PivotTables and charts from 18 months of funnel metrics and a sourcing summary, "
            "choose each chart from the question it answers, refine titles, labels, sort order and axes with follow-up "
            "prompts, and reconcile every chart to an independent SUMIFS check."
        ),
        build=(
            "BrightPath_Recruitment_Metrics.xlsx with PivotTables and four refined charts on the Charts sheet: a monthly hires "
            "line, a sorted 2026 hires-by-sector bar, a 2026 recruitment funnel and a hire-rate vs cost-per-hire combo chart"
        ),
        services="Excel · Copilot (Allow editing) · PivotTables · line, bar, funnel and combo charts · BrightPath_Recruitment_Metrics.xlsx",
        data_files=["data/BrightPath_Recruitment_Metrics.xlsx"],
        prerequisites=[
            "Topic 1 completed: you can open a OneDrive workbook, use Copilot modes and verify with formulas.",
            "This lab folder's data/BrightPath_Recruitment_Metrics.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Recruitment_Metrics.xlsx")
                + " Write the four questions below at the top of the Charts sheet - each chart must answer one of them.",
                "Tables : tblMonthly (108 rows = 18 months x 6 sectors) | tblSources (7 rows)\n\n"
                "Q1 Trend      : How have total hires moved month by month since Jan 2025?\n"
                "Q2 Comparison : Which sector produced the most hires in Jan-Jun 2026?\n"
                "Q3 Funnel     : Where do candidates drop out between application and hire in 2026?\n"
                "Q4 Efficiency : Which sources convert best, and at what cost per hire?",
            ),
            (
                "Choose Allow editing. Ask Copilot for the trend PivotTable and line chart. Check the Pivot grand total against "
                "=SUM(tblMonthly[Hires]) and find the peak month.",
                "Create a PivotTable from tblMonthly on a new sheet named Pivot_Trend with Month_Label in rows and\n"
                "Sum of Hires as values. Then create a line chart of total hires by month from that PivotTable.\n"
                "Show data labels and gridlines. Title: \"Monthly hires, Jan 2025 - Jun 2026\".",
            ),
            (
                "Ask for the 2026 sector comparison. Then refine it with a follow-up prompt. Verify Technology's 2026 hires "
                "with the SUMIFS shown.",
                "Create a PivotTable on a new sheet Pivot_Sector from tblMonthly with Sector in rows and Sum of Hires,\n"
                "filtered to months in 2026 only. Create a bar chart from it.\n\n"
                "FOLLOW-UP: Sort the bars from highest to lowest and change the title to state the leading sector.\n\n"
                "CHECK: =SUMIFS(tblMonthly[Hires],tblMonthly[Sector],\"Technology\",tblMonthly[Month],\">=\"&DATE(2026,1,1))",
            ),
            (
                "Ask Copilot to build a 2026 funnel summary with formulas, then a funnel chart. Confirm the summary uses SUMIFS "
                "(not typed numbers) by clicking a value cell.",
                "On the Charts sheet, build a small table of 2026 totals for Applications, Screened, Interviews, Offers\n"
                "and Hires from tblMonthly using SUMIFS on Month >= 1 Jan 2026. Then create a funnel chart from it\n"
                "with data labels. Add a column showing each stage as a % of Applications.",
            ),
            (
                "Ask Copilot to add rate columns to tblSources and a combo chart with two axes. Check the best hire rate and the "
                "lowest cost per hire.",
                "In tblSources add Hire_Rate = Hires / Applications (percentage, 1 decimal) and\n"
                "Cost_per_Hire = Sourcing_Spend / Hires (S$, no decimals). Create a combo chart: Hire_Rate as\n"
                "clustered columns sorted highest first, Cost_per_Hire as a line on a secondary axis.\n"
                "Label both axes with units.",
            ),
            (
                "Refine all four charts with short follow-up prompts, one change at a time, then move them onto the Charts sheet "
                "in a 2 x 2 grid. Add alt text to each chart.",
                "Make every chart title state the finding in one sentence.\n"
                "Use the same blue colour for all hires series and format numbers with thousands separators.\n"
                "Move all four charts to the Charts sheet in a 2 x 2 grid of equal size.\n"
                "Add alt text to each chart describing what it shows.",
            ),
            (
                "Ask Copilot (Chat only) to review the charts against the four questions, then log your prompts and checks on "
                "Prompt_Log.",
                "Review the four charts on the Charts sheet. For each, tell me which question at the top of the sheet\n"
                "it answers, whether the chart type suits that question, and one improvement. Do not change anything.",
            ),
        ],
        sample_prompts=[
            ("PivotTable", "Create a PivotTable from tblMonthly with Month_Label in rows, Sector in columns and Sum of Hires as values."),
            ("Line chart", "Create a line chart showing total hires by month. Show data labels and gridlines."),
            ("Bar chart", "Create a bar chart of 2026 hires by Sector, sorted from highest to lowest."),
            ("Funnel", "Create a funnel chart of 2026 Applications, Screened, Interviews, Offers and Hires with data labels."),
            ("Combo chart", "Create a combo chart of Hire_Rate as columns and Cost_per_Hire as a line on a secondary axis."),
            ("Refine", "Change the chart title to: Technology led hiring in H1 2026."),
            ("Refine", "Sort the bars from highest to lowest and add data labels formatted with thousands separators."),
            ("Accessibility", "Add alt text to this chart that describes the trend it shows."),
            ("Review", "Is a line chart the right choice for this data? Suggest a better chart if not, and explain why."),
        ],
        test=(
            f"Pivot_Trend grand total = {E3['h1_hires'] + E3['total_2025']:,} hires (= SUM(tblMonthly[Hires])); the peak month is "
            f"{E3['peak_month']} with {E3['peak_month_hires']} hires. The 2026 sector bar chart shows {E3['top_sector']} first with "
            f"{E3['top_sector_hires']} hires and its Pivot total is {E3['h1_hires']}. The funnel shows {_F[0]:,} > {_F[1]:,} > "
            f"{_F[2]:,} > {_F[3]:,} > {_F[4]:,}. In tblSources, {E3['best_rate_source']} has the highest hire rate "
            f"({E3['best_rate']}%) and {E3['lowest_cps_source']} the lowest cost per hire (S${E3['lowest_cps']}). All four charts "
            "sit on the Charts sheet with finding-style titles, labelled axes and alt text."
        ),
        checkpoint="Keep the workbook in work-c34. Lab 4 uses its own placements workbook.",
        troubleshooting=[
            (
                "The 2026 filter includes 2025 months.",
                "Ask Copilot to filter on the Month date field (>= 1 Jan 2026), not on Month_Label text, or add a Year field.",
            ),
            (
                "Copilot inserts a picture of a chart instead of a chart.",
                "That came from a Chat answer and is static. Switch to Allow editing and ask it to create a chart linked to the data.",
            ),
            (
                "The funnel chart type is not offered.",
                "Funnel charts need Excel 2019 or Microsoft 365. Ask for a sorted horizontal bar chart of the stages instead.",
            ),
            (
                "Both series in the combo chart share one axis and the line is flat.",
                "Ask: 'Put Cost_per_Hire on a secondary axis and format it as S$'.",
            ),
        ],
        challenge=(
            "Ask Copilot for a 100% stacked column chart showing each sector's share of hires by quarter for 2025-2026, then "
            "decide whether it communicates better or worse than the bar chart, and why."
        ),
        reflection="Which of your four charts would most change a BrightPath leadership decision, and what check makes you trust it?",
    ),
    # ------------------------------------------------------------------ LAB 4
    dict(
        num=4,
        topic=2,
        title="Visualise Pay and Placement Distributions with Python in Excel",
        duration=35,
        objective=(
            "LO2: use Copilot with Python in Excel to create distribution, relationship and heatmap visuals, review the "
            "generated code and verify plotted statistics with native formulas"
        ),
        goal=(
            "Show how contract bill rates, pay and margins are distributed across job levels and sectors, and explain one "
            "outlier, using Python visuals you have checked."
        ),
        workflow=["Margin column", "Box plot", "Scatter", "Heatmap", "Outlier check"],
        desc=(
            "BrightPath's flexible-staffing desk places contract associates with clients. You will ask Copilot to add a gross "
            "margin column, then create Python in Excel visuals - a box plot of bill rates by job level, a scatter of "
            "experience against pay with a trend line, and a sector-by-level margin heatmap. You will read the =PY code, "
            "re-compute a median and a correlation with Excel formulas, and investigate the highest pay rate."
        ),
        build=(
            "BrightPath_Contract_Placements.xlsx with a Gross_Margin_Pct column and, on Python_Charts, a box plot, a scatter "
            "with trend line, a margin heatmap and an outlier table - each with a native-formula check beside it"
        ),
        services="Excel · Copilot (Allow editing, Chat only) · Python in Excel (=PY) · BrightPath_Contract_Placements.xlsx",
        data_files=["data/BrightPath_Contract_Placements.xlsx"],
        prerequisites=[
            "Lab 3 completed or equivalent charting experience.",
            "Python in Excel available: Formulas > Insert Python is visible (Windows, Mac or web; internet required).",
            "This lab folder's data/BrightPath_Contract_Placements.xlsx.",
        ],
        steps=[
            (
                _SETUP_STEP.format(file="BrightPath_Contract_Placements.xlsx")
                + " Check that the Formulas tab shows Insert Python. If it does not, follow the fallback in Troubleshooting.",
                "Table  : tblPlacements (180 rows)\n"
                "Levels : Junior | Mid | Senior | Lead\n"
                "Output sheet : Python_Charts",
            ),
            (
                "Choose Allow editing and ask for the margin column. Check one row by hand and the average with AVERAGE.",
                "Add a column Gross_Margin_Pct to tblPlacements = (Bill_Rate_Hourly - Pay_Rate_Hourly) / Bill_Rate_Hourly,\n"
                "formatted as a percentage with 1 decimal.\n\n"
                "CHECK: =AVERAGE(tblPlacements[Gross_Margin_Pct])",
            ),
            (
                "Ask Copilot to create the box plot with Python. Click the =PY cell and read the code: which table and columns it "
                "uses and the order of levels. Then verify the Lead median with a native formula in the cell beside it.",
                "Using Python in Excel, create a box plot of Bill_Rate_Hourly by Job_Level from tblPlacements,\n"
                "ordered Junior, Mid, Senior, Lead. Put it in Python_Charts!B2. Title it\n"
                "\"Contract bill rate by job level (S$/hour)\" and label both axes.\n\n"
                "CHECK: =MEDIAN(FILTER(tblPlacements[Bill_Rate_Hourly],tblPlacements[Job_Level]=\"Lead\"))",
            ),
            (
                "Ask for a scatter plot with a trend line and the correlation coefficient. Verify r with CORREL.",
                "Using Python, create a scatter plot of Years_Experience (x) against Pay_Rate_Hourly (y) from\n"
                "tblPlacements, coloured by Job_Level, with a linear trend line. Also return the Pearson correlation\n"
                "coefficient in the cell below the chart.\n\n"
                "CHECK: =CORREL(tblPlacements[Years_Experience],tblPlacements[Pay_Rate_Hourly])",
            ),
            (
                "Ask for a heatmap of average margin. Read the code to confirm it uses the mean of Gross_Margin_Pct per cell.",
                "Using Python, create a heatmap of the average Gross_Margin_Pct by Sector (rows) and Job_Level (columns,\n"
                "ordered Junior to Lead) with the values printed in each cell as percentages.",
            ),
            (
                "Ask Copilot (Chat only) to find outliers with a stated rule, then look at the highest pay rate row. Decide in "
                "Prompt_Log whether it is an error, a genuine specialist rate, or needs checking with the desk - an outlier is not "
                "automatically wrong.",
                "Using the IQR rule (above Q3 + 1.5 x IQR), list the Pay_Rate_Hourly outliers in tblPlacements with\n"
                "Placement_ID, Sector, Job_Level and Pay_Rate_Hourly. State Q1, Q3 and the upper fence. Show the code.",
            ),
            (
                "Ask Copilot to explain what the Python code in the box-plot cell does, then record all prompts and checks in "
                "Prompt_Log.",
                "Explain the Python code in Python_Charts!B2 line by line in plain English.",
            ),
        ],
        sample_prompts=[
            ("Calculated column", "Add Gross_Margin_Pct = (Bill_Rate_Hourly - Pay_Rate_Hourly) / Bill_Rate_Hourly as a percentage."),
            ("Box plot", "Using Python, create a box plot of Bill_Rate_Hourly by Job_Level ordered Junior, Mid, Senior, Lead."),
            ("Histogram", "Using Python, create a histogram of Pay_Rate_Hourly with 15 bins and mark the median."),
            ("Scatter", "Using Python, plot Years_Experience against Pay_Rate_Hourly with a linear trend line and return the correlation."),
            ("Heatmap", "Using Python, create a heatmap of average Gross_Margin_Pct by Sector and Job_Level with values shown."),
            ("Outliers", "Using the IQR rule, list Pay_Rate_Hourly outliers with Placement_ID, Sector and Job_Level, and state the upper fence."),
            ("Explain code", "Explain the Python code in the selected cell line by line."),
            ("Direct answer", "What are the median and mean Bill_Rate_Hourly by Job_Level? Show the code."),
        ],
        test=(
            f"Average Gross_Margin_Pct = {E4['avg_margin']}%. The box plot shows four levels in order and the Lead median bill rate "
            f"check returns S${E4['median_bill']['Lead']:.2f} (Junior S${E4['median_bill']['Junior']:.2f}, Mid "
            f"S${E4['median_bill']['Mid']:.2f}, Senior S${E4['median_bill']['Senior']:.2f}). CORREL returns about {E4['corr']} - a "
            f"moderate positive relationship. The heatmap shows {E4['low_margin_sector']} with the lowest average margin "
            f"(~{E4['sector_margin'][E4['low_margin_sector']]}%). The IQR rule flags {E4['pay_outliers']} pay-rate outliers above "
            f"about S${E4['upper_fence']:.2f}; the highest is {E4['max_pay_id']} at S${E4['max_pay']:.2f}, and your Prompt_Log "
            "records a decision about it."
        ),
        checkpoint="Keep the workbook in work-c34. Topic 3 starts from the dashboard workbook in the Lab 5 folder.",
        troubleshooting=[
            (
                "Insert Python is not visible or =PY returns an error.",
                "Python in Excel needs a supported plan, Windows/Mac/web and internet. Fallback: ask Copilot in Chat only for the "
                "same visual (it is returned as a static image) and use Insert > Charts > Statistical > Box and Whisker for a native box plot.",
            ),
            (
                "The =PY cell shows a Python object instead of a picture.",
                "Select the cell and change its output (the icon beside the formula bar) from Python Object to Excel Value, or "
                "click the card icon to view the image.",
            ),
            (
                "#BUSY! or a slow calculation.",
                "Python runs in the Microsoft cloud. Wait, or select Formulas > Reset (Python) and recalculate.",
            ),
            (
                "Levels appear in alphabetical order (Junior, Lead, Mid, Senior).",
                "Ask Copilot to set the category order explicitly to Junior, Mid, Senior, Lead.",
            ),
        ],
        challenge=(
            "Ask Copilot, using Python, whether contracts that were Extended have a higher average Client_Rating than those "
            "that were not, with counts. Re-check the two averages with AVERAGEIFS."
        ),
        reflection="Which Python visual told you something a column chart of averages would have hidden?",
    ),
]
