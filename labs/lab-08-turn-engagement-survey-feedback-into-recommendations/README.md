# Lab 8 — Turn Engagement Survey Feedback into Recommendations

**Course:** Microsoft 365 Copilot for Excel  
**Course Code:** C34  
**Version:** v2.0 (8 October 2026)  
**Topic 4:** Use Copilot to Draw Insights and Recommendations  
**Maps to:** LO4: use Copilot to quantify survey scores and eNPS, classify comment themes and sentiment with a human spot-check, and draft owned, measurable recommendations and an executive summary from verified findings  
**Duration:** 30 minutes  
**Tools:** Excel · Copilot (Chat only, Allow editing) · COPILOT function (where available) · BrightPath_Engagement_Survey.xlsx

> 📄 **Instruction PDF (with sample prompts):** [Lab-08-Instructions.pdf](Lab-08-Instructions.pdf)  
> 💬 **Copy-ready prompts:** [sample-prompts.txt](sample-prompts.txt)  
> 🌐 **Prompts web page (one-click copy):** [Lab 8 prompts](https://tertiarycourses.github.io/C34-Microsoft-365-Copilot-for-Excel/prompts/#lab-8)  
> 📊 **Mock data:** [data/BrightPath_Engagement_Survey.xlsx](data/BrightPath_Engagement_Survey.xlsx)  

---

## Goal

Produce three evidence-backed recommendations and a 150-word executive summary from BrightPath's June engagement pulse survey, with every number verified.

## What You Will Do

BrightPath ran an anonymous pulse survey of 160 staff. You will have Copilot calculate eNPS and question averages and verify them with COUNTIFS and AVERAGE, compare departments while respecting small groups, add Theme and Sentiment columns from a fixed list and spot-check them, then use the Recommendations template and Copilot to turn verified findings into owned actions and an executive summary you approve.

## What You Will Build

BrightPath_Engagement_Survey.xlsx with verified eNPS and question averages, Theme and Sentiment columns with a recorded spot-check, a Recommendations sheet with three owned, measurable actions, and an Executive_Summary

## Prerequisites

- Lab 7 completed: you can verify an insight and label it OBSERVED or HYPOTHESIS.
- This lab folder's data/BrightPath_Engagement_Survey.xlsx.

> **Data note.** Use only the supplied fictional BrightPath data or information you are authorised to process. Do not paste credentials, confidential business material or real employees' personal data into Copilot.

## Steps

### 1. Go to onedrive.com (or your SharePoint site), open or create the folder work-c34 and select Upload > Files. Upload BrightPath_Engagement_Survey.xlsx from this lab's data folder, then open it in Excel for the web or select Open in Desktop App. Confirm the title bar shows the cloud location and AutoSave is On, and that Formulas > Calculation Options is set to Automatic. Read the About_This_Data sheet before you prompt. Responses are anonymous: do not try to identify anyone, and do not report a group with fewer than 5 responses.

```text
Table : tblSurvey (160 responses; Q1-Q6 scored 1-5; eNPS_Score 0-10; Comment)
eNPS  = % promoters (9-10) - % detractors (0-6)
Templates: Recommendations | Executive_Summary
```

### 2. Choose Chat only and ask for eNPS and the question averages. Then ask Copilot (Allow editing) to add the check formulas on a new Checks sheet.

```text
CHAT ONLY:
Using tblSurvey, calculate the number of promoters (9-10), passives (7-8) and detractors (0-6) on
eNPS_Score, and the eNPS. Then give the average of Q1 to Q6, lowest to highest.

CHECK (Allow editing):
Create a sheet named Checks. Enter these formulas in B1:B4 with a label beside each in A1:A4,
then tell me each result:
=COUNTIFS(tblSurvey[eNPS_Score],">=9")
=COUNTIFS(tblSurvey[eNPS_Score],"<=6")
=(COUNTIFS(tblSurvey[eNPS_Score],">=9")-COUNTIFS(tblSurvey[eNPS_Score],"<=6"))/ROWS(tblSurvey[Response_ID])*100
=AVERAGE(tblSurvey[Q3_Workload_Balance])
```

### 3. In Chat only, ask for eNPS by department with response counts. Mark any department with fewer than 10 responses as 'interpret with caution' and never report one with fewer than 5.

```text
Show eNPS and number of responses for each Department in tblSurvey, sorted lowest eNPS first.
```

### 4. Choose Allow editing and ask Copilot to classify comments into a fixed list. If your Excel has the COPILOT function you may ask Copilot to use the formula option instead. Keep the original Comment column unchanged.

```text
OPTION A (Copilot pane):
Add two columns to tblSurvey. Theme: classify each Comment into exactly one of Workload, Career Growth,
Manager Support, Pay & Benefits, Recognition, Tools & Systems, Flexibility, or Unclear. Sentiment:
Positive, Negative or Mixed. Leave both blank where Comment is blank. Do not change Comment.

OPTION B (COPILOT function, where available):
In tblSurvey, add a Theme column, enter this formula in its first cell and fill it down the column:
=COPILOT("Classify this staff comment into exactly one theme from the list. Reply with the theme only.",
         [@Comment], "List:", "Workload, Career Growth, Manager Support, Pay & Benefits, Recognition,
         Tools & Systems, Flexibility, Unclear")
```

### 5. Spot-check: ask Copilot to show 10 classified comments, read them and count any you disagree with. If more than 1 in 10 is wrong, refine the prompt (for example, add an example per theme) and rerun. Then ask for theme counts, the COUNTIF check, and ask Copilot to record the error rate on Prompt_Log.

```text
CHAT ONLY:
Show me 10 random rows of tblSurvey with Comment, Theme and Sentiment only.

CHAT ONLY:
Count comments by Theme and by Sentiment in tblSurvey, highest first. For the top theme, summarise
what staff are saying in three bullet points without quoting anything that could identify a person.

CHECK (Allow editing):
On the Checks sheet, add this formula and tell me if it matches your Workload count:
=COUNTIF(tblSurvey[Theme],"Workload")

ALLOW EDITING:
Add a row to Prompt_Log: "Theme spot-check: <number> of 10 disagreed".
```

### 6. Choose Allow editing to fill the Recommendations sheet: ask Copilot to draft three rows from your verified findings, then tell Copilot what to change so the owners, targets and dates are realistic. Every Evidence cell must cite numbers you verified.

```text
Using only these verified findings: <paste eNPS, lowest question, top theme, department results>,
draft three recommendations in the Recommendations sheet columns: Finding (OBSERVED), Evidence,
Recommendation (a team, process or policy action - not about any individual), Owner (role),
Priority, Success KPI and target, Review date (by Dec 2026).

FOLLOW-UP: In the Recommendations sheet, change <row and column> to <new owner, target or date>.
```

### 7. Ask Copilot to draft the executive summary on Executive_Summary. Check every number against your Checks sheet, tell Copilot what wording to change, and ask it to add your decision: Share, Revise or Hold. Then ask Copilot to log all prompts. Optionally ask Copilot in Word or PowerPoint to turn the summary into a memo or two slides.

```text
Draft a 150-word executive summary for BrightPath's leadership from the Recommendations sheet and the
verified results. Include eNPS, the lowest-scoring question, the top comment theme and the three
recommendations with owners. Use plain English, no jargon, and do not claim causes.

FOLLOW-UP: Add a final line to Executive_Summary: "Decision: <Share / Revise / Hold> - <one-line reason>".

LOG (Allow editing):
Add one row to Prompt_Log for each prompt I ran in this lab: Mode, Prompt, What Copilot did,
How it was verified, and Decision = Keep. I will tell you any row to change to Revise or Undo.
```

## Sample Prompts

| # | Purpose | Prompt |
|---|---|---|
| 1 | Score | Calculate eNPS from eNPS_Score in tblSurvey: promoters 9-10, passives 7-8, detractors 0-6. Show the counts. |
| 2 | Averages | Give the average of each question Q1 to Q6, lowest to highest, with the number of responses. |
| 3 | Groups | Show eNPS and number of responses by Department, sorted lowest first. |
| 4 | Themes | Classify each Comment into exactly one theme from my list (or Unclear) in a new Theme column. Do not change Comment. |
| 5 | Sentiment | Add a Sentiment column: Positive, Negative or Mixed for each non-blank Comment. |
| 6 | COPILOT function | =COPILOT("Classify the sentiment of this comment as Positive, Negative or Mixed", [@Comment]) |
| 7 | Summarise text | Summarise what staff say about Workload in three bullet points without quoting anything identifying. |
| 8 | Recommend | From these verified findings, draft three recommendations with owner, priority, success KPI and review date. |
| 9 | Executive summary | Draft a 150-word executive summary for leadership. Do not claim causes; cite the numbers. |

## Test It

Promoters 39, passives 61, detractors 60 of 160; eNPS = -13.1. Lowest question Workload Balance (2.71); highest Manager Support (3.76). Human Resources has the lowest department eNPS (-36.4) from only 11 responses, so it is flagged 'interpret with caution'. 134 responses have a comment; after classification Workload is the most frequent theme (about 39 comments - AI labels can vary by a few, which is why the spot-check error rate is recorded). The Recommendations sheet has three rows, each with verified evidence, an owner role, a KPI target and a review date, and Executive_Summary ends with a Share, Revise or Hold decision.

## Checkpoint and Rejoin Point

This is the final lab. Keep all eight workbooks in work-c34 as your personal reference set.

## Troubleshooting

| If this happens | Fix |
|---|---|
| Copilot invents new themes outside the list. | Repeat the exact list, add 'reply with one theme from the list only', and ask it to use Unclear when unsure. |
| The COPILOT function returns #NAME? or is missing. | It is not available on every licence or channel yet. Use Option A in the Copilot pane instead. |
| Copilot's eNPS uses 7-10 as promoters or averages the scores. | Restate the rule: promoters 9-10, detractors 0-6, eNPS = % promoters - % detractors. Trust your COUNTIFS. |
| The summary quotes a comment that could identify someone. | Remove it and ask Copilot to paraphrase themes without quotes. Anonymity comes before vividness. |

## Challenge

Combine Lab 7 and Lab 8: ask Copilot whether the departments with the highest attrition also have the lowest Workload Balance scores, then write one recommendation that uses both sources - labelled as a HYPOTHESIS.

## Reflection

Which recommendation would you be confident defending to the HR Director, and which number makes it defensible?

---

[← Lab 7](../lab-07-surface-attrition-insights-with-copilot-and-advanced-analysis/README.md) · [Labs index →](../README.md)
