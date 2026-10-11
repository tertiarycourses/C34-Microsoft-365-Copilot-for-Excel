# Lab 7 — Surface Attrition Insights with Copilot and Advanced Analysis

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 4:** Use Copilot to Draw Insights and Recommendations  
**Maps to:** LO4: use Copilot direct answers and Python-based advanced analysis to surface attrition patterns, verify them with native formulas and record each as OBSERVED evidence or a HYPOTHESIS with its limitation  
**Duration:** 40 minutes  
**Tools:** Excel · Copilot (Chat only, advanced analysis mode) · Python in Excel · BrightPath_Attrition_Analysis.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-07-Instructions.pdf](Lab-07-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 🌐 **Prompts web page (one-click copy):** [Lab 7 prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/#lab-7)  
> 📊 **Mock data:** [data/BrightPath_Attrition_Analysis.xlsx](data/BrightPath_Attrition_Analysis.xlsx)  

---

## Goal

Identify the work-pattern factors most associated with leaving BrightPath, with verified rates, group sizes and honest limitations in an Insight_Log.

## What You Will Do

BrightPath lost 14% of a 400-person study group in 12 months. You will start with an open 'What's interesting?' prompt, set the baseline attrition rate, ask Copilot for specific group comparisons with counts, enter advanced analysis mode to compare leavers and stayers with Python, re-compute the headline numbers with COUNTIFS and AVERAGEIFS, and write four labelled insights - including why a high rate in a small team is only a signal.

## What You Will Build

BrightPath_Attrition_Analysis.xlsx with verified attrition rates by department and factor, an advanced-analysis sheet with reviewed Python output, and an Insight_Log of four OBSERVED or HYPOTHESIS insights with evidence and limitations

## Prerequisites

- Labs 1-6 completed, or equivalent skill with Copilot prompts, PivotTables and COUNTIFS.
- Python in Excel available for advanced analysis (fallback in Troubleshooting).
- This lab folder's data/BrightPath_Attrition_Analysis.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Attrition_Analysis.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Note the warning on About_This_Data: these are team-level signals, never a basis for decisions about an individual.

```text
Table   : tblAttrition (400 employees active on 2025-07-01)
Outcome : Left_Company = Yes / No (left within 12 months)
Template: Insight_Log sheet
```

### 2. Choose Chat only and ask the open question. Copy Copilot's candidate patterns into Prompt_Log - do not record them as insights yet; each must be verified first.

```text
What's interesting about tblAttrition in relation to Left_Company? Give up to five patterns, each with
the numbers and group sizes behind it. Do not change the workbook.
```

### 3. Set the baseline: ask for the overall rate, then check it with formulas in two empty cells.

```text
What is the overall attrition rate in tblAttrition (Left_Company = "Yes" divided by all employees)?

CHECK:
=COUNTIFS(tblAttrition[Left_Company],"Yes")
=COUNTIFS(tblAttrition[Left_Company],"Yes")/ROWS(tblAttrition[Employee_ID])
```

### 4. Ask for specific comparisons with counts. Verify at least the engagement and overtime splits with COUNTIFS, and note which departments have fewer than 30 people.

```text
Using tblAttrition, show attrition rate AND number of employees for:
1) each Department;
2) Engagement_Score below 3 vs 3 and above;
3) Overtime_Hours_Month above 20 vs 20 or below;
4) Years_Since_Promotion 3 or more vs below 3;
5) Compa_Ratio below 0.9 vs 0.9 and above;
6) each Work_Mode.

CHECK (engagement < 3):
=COUNTIFS(tblAttrition[Engagement_Score],"<3",tblAttrition[Left_Company],"Yes")/COUNTIFS(tblAttrition[Engagement_Score],"<3")
```

### 5. Go deeper: ask Copilot to enter advanced analysis mode (or select 'Get deeper results using advanced analysis mode'). Copilot opens a new sheet with Python. Read the code: which rows, which columns, how blanks are handled. Verify two of the means with AVERAGEIFS, then select Stop advanced analysis.

```text
Enter advanced analysis mode.
Compare employees who left with those who stayed on Overtime_Hours_Month, Engagement_Score,
Years_Since_Promotion, Compa_Ratio and Training_Hours. Show the mean for each group, the group sizes,
and the correlation between Overtime_Hours_Month and Engagement_Score. Explain the method.

CHECK:
=AVERAGEIFS(tblAttrition[Overtime_Hours_Month],tblAttrition[Left_Company],"Yes")
=AVERAGEIFS(tblAttrition[Engagement_Score],tblAttrition[Left_Company],"Yes")
```

### 6. Ask Copilot to help label your findings, then edit them yourself into four rows of the Insight_Log sheet. Each row needs the verified numbers, a label, a possible explanation to test and a limitation.

```text
Here are my verified findings: <paste your checked rates and means>.
For each, write one OBSERVED sentence with the numbers, one HYPOTHESIS that could explain it, and
one limitation (sample size, correlation is not causation, overlapping factors). Keep each under 30 words.
```

### 7. Ask Copilot (Chat only) to challenge your Insight_Log, then log all prompts on Prompt_Log.

```text
Review my Insight_Log. Which statement overstates the evidence or implies causation? How would
you reword it?
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Open question | What's interesting about this data in relation to Left_Company? Give five patterns with numbers and group sizes. |
| 2 | Baseline | What is the overall attrition rate in tblAttrition? Show the numerator and denominator. |
| 3 | Comparison | Show attrition rate and headcount for Engagement_Score below 3 vs 3 and above. |
| 4 | By group | Show attrition rate and number of employees by Department, sorted highest first. |
| 5 | Advanced analysis | Enter advanced analysis mode. Compare leavers and stayers on overtime, engagement, promotion, compa-ratio and training. |
| 6 | Correlation | What is the correlation between Overtime_Hours_Month and Engagement_Score, and what does it mean for interpreting the attrition patterns? |
| 7 | Outliers | Are there any managers (Manager_ID) whose teams have unusually high attrition? Only include teams with at least 10 people. |
| 8 | Challenge | Which of my insights could be a coincidence or caused by another factor? Explain. |

## Test It

Overall attrition = 56 of 400 = 14.0%. Engagement below 3: 29.6% (n = 81) vs 10.0% (n = 319). Overtime above 20 hours: 34.0% (n = 47) vs 11.3% (n = 353). Promotion 3+ years ago: 21.3% vs 13.0%. On-site staff: 19.3%. IT shows the highest department rate (28.0%) but has only 25 people, so it is recorded as a small-sample signal. Leavers average 15.9 overtime hours and engagement 3.10, vs 10.6 and 3.54 for stayers. Overtime and engagement correlate at about -0.25, so the two factors overlap. Insight_Log has four rows, each labelled and with a limitation.

## Checkpoint and Rejoin Point

Keep the workbook in work-c34 and keep your Insight_Log open - Lab 8 turns verified insights like these into recommendations.

## Troubleshooting

| If this happens | Fix |
|---|---|
| Advanced analysis mode is not offered. | It needs Python in Excel. Stay in Chat only: direct answers also use Python and can show code. Or build the comparisons with a PivotTable (Left_Company in columns, Count of Employee_ID, % of row total). |
| Copilot gives a rate without a group size. | Follow up: 'Show the number of employees in each group next to each rate.' Never record a rate without n. |
| Copilot says a factor 'causes' attrition. | Ask it to restate as an association, and record any causal idea as a HYPOTHESIS to test (e.g. exit interviews). |
| Your COUNTIFS rate differs slightly from Copilot's. | Check the boundary (< 3 vs <= 3, > 20 vs >= 20) and whether blanks were excluded. Use your formula. |

## Challenge

In advanced analysis mode, ask Copilot to fit a simple logistic regression of Left_Company on the five factors and explain the output in plain English. Record why this still does not prove cause.

## Reflection

Which insight would you be most careful presenting to managers, and what extra evidence would you want first?

---

[← Lab 6](../lab-06-build-an-interactive-workforce-and-attrition-dashboard/README.md) · [Lab 8 →](../lab-08-turn-engagement-survey-feedback-into-recommendations/README.md)
