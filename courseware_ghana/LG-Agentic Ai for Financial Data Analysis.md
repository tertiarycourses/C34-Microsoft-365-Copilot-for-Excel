# Agentic Ai for Financial Data Analysis — Learner Guide

**Course Code:** C34  |  **Conducted by:** Tertiary Infotech Academy Pte Ltd (UEN 201200696W)  |  **Version v1.6 · 24 November 2026**

## Contents

- [Introduction](#introduction)
- [Course Learning Outcomes](#course-learning-outcomes)
- [Before You Start — Preparation](#before-you-start--preparation)
- [Topic 01 — Business Innovation with Agentic AI and AI Agents](#topic-01--business-innovation-with-agentic-ai-and-ai-agents)
  - [From Rules to AI Agents](#from-rules-to-ai-agents)
  - [Agentic AI and AI Agents](#agentic-ai-and-ai-agents)
  - [Where Agents Add Business Value](#where-agents-add-business-value)
  - [The Tema Distributor Business Model Canvas](#the-tema-distributor-business-model-canvas)
  - [Always-On Agents in Finance Teams](#always-on-agents-in-finance-teams)
  - [Tools, Skills and Context](#tools-skills-and-context)
  - [Lab 1 · Meta Muse — Prioritise Month-End Exceptions with Meta Muse](#lab-1--meta-muse--prioritise-month-end-exceptions-with-meta-muse)
  - [Topic 01 Recap - Mapped to Learning Outcomes](#topic-01-recap---mapped-to-learning-outcomes)
- [Topic 02 — Create an Interactive Dashboard in Excel](#topic-02--create-an-interactive-dashboard-in-excel)
  - [Sales and Financial Measures](#sales-and-financial-measures)
  - [Charts and Interactive Controls](#charts-and-interactive-controls)
  - [Restyle Without Breaking the Dashboard](#restyle-without-breaking-the-dashboard)
  - [Lab 2 · Interactive Excel dashboard — Create and Restyle an Interactive Sales Dashboard](#lab-2--interactive-excel-dashboard--create-and-restyle-an-interactive-sales-dashboard)
  - [Topic 02 Recap - Mapped to Learning Outcomes](#topic-02-recap---mapped-to-learning-outcomes)
- [Topic 03 — Skills](#topic-03--skills)
  - [Convert a Prompt into a Skill](#convert-a-prompt-into-a-skill)
  - [Install, Enable and Invoke the Skill](#install-enable-and-invoke-the-skill)
  - [Lab 3 · Excel custom skills — Install and Reuse the Dashboard Skill](#lab-3--excel-custom-skills--install-and-reuse-the-dashboard-skill)
  - [Topic 03 Recap - Mapped to Learning Outcomes](#topic-03-recap---mapped-to-learning-outcomes)
- [Topic 04 — Analyse a Financial Report with Copilot](#topic-04--analyse-a-financial-report-with-copilot)
  - [Ground Copilot in the Published Report](#ground-copilot-in-the-published-report)
  - [Reconcile and Record the Evidence](#reconcile-and-record-the-evidence)
  - [Lab 4 · Copilot financial review — Review the Bank of Ghana Financial Statements](#lab-4--copilot-financial-review--review-the-bank-of-ghana-financial-statements)
  - [Topic 04 Recap - Mapped to Learning Outcomes](#topic-04-recap---mapped-to-learning-outcomes)
- [AI Security in Finance](#ai-security-in-finance)
- [Next Steps](#next-steps)
- [Glossary](#glossary)


## Introduction

This Datin edition is for nine accounting professionals from West Africa. Work as the finance team of a fictional wholesale distributor in Tema, Ghana, using only mock month-end and sales data. First use Meta Muse to prioritise cut-off evidence, a duplicate invoice and a credit-note request. Then shift from exception triage to management reporting. Use the same distributor’s mock H1 orders to build and restyle a sales dashboard and reconcile its figures in Ghana cedi (GHS).

Lab 3 packages the H1 dashboard style and build steps into a OneDrive skill. Use the distributor’s fresh July–December orders to check whether the same reporting workflow produces accurate new figures. In Topic 4, move from mock sales data to the published Bank of Ghana report and verify a Word review memo.


## Course Learning Outcomes

- LO1: Describe the development of AI, distinguish an AI agent from agentic AI, and explain where tools, skills and approvals fit.
- LO2: Build an interactive sales dashboard in Excel, reconcile its GHS metrics and change its theme while preserving interactive behavior.
- LO3: Turn a styling prompt into an Excel skill, enable it through OneDrive and reuse it on a fresh workbook with a short request.
- LO4: Review a published financial report with Copilot, verify the figures and record the evidence and review limits in Word.


## Before You Start — Preparation

**What you need**

- Use Microsoft 365 Excel with Copilot editing, a permitted work account and OneDrive. Verify that Manage plugins & skills is visible.
- Check Meta Muse access or join the trainer demonstration. Keep the supplied mock material and lab folders together.
- Open prompts/index.html from the Datin folder for the copy-ready prompts.

**Verify your setup**

Open the Lab 2 starter as an editable OneDrive copy with AutoSave on. Check the Sales Data and Data Guide tabs. In Copilot, verify that workbook editing is available.

**Conventions used in every lab**

- GHS means Ghana cedi. Use the Data Guide to reconcile All, Greater Accra/Wholesale and Greater Accra/Wholesale/Rice.
- Use class-only copies and record the before and after values when Copilot modifies a dashboard.


## Topic 01 — Business Innovation with Agentic AI and AI Agents

A brief history of AI, business use cases, tools and skills, and one Meta Muse activity.

**Key concepts**

- AI has several eras — Rules, machine learning, deep learning and large language models coexist in modern systems.
- Agentic AI describes behaviour — A system pursues a goal by planning, acting and checking its results.
- An agent is a deployed worker — The agent carries out a defined job through the permissions and tools it has.
- Tools and skills — A tool performs an operation. A skill packages a repeatable way to work and verify the outcome.


### From Rules to AI Agents

Symbolic systems encoded rules and knowledge. Statistical machine learning learned patterns from examples. Deep learning scaled representation learning, and large language models made fluent language interfaces practical.

These approaches build on and coexist with each other. Businesses still use rules to enforce policy, predictive models to score risk and databases to ground answers. An agent adds an execution layer that can call those existing systems.

**How it works**

- Start with explicit policy rules.
- Use a learned model where a pattern can be inferred from examples.
- Use a language model to interpret instructions and explain results.
- Add tools, a feedback loop and runtime controls to pursue a delegated goal.

**A Short History of AI**

| Aspect | Capability | Business example |
|---|---|---|
| 1950s–1980s | Symbolic AI and expert systems | Apply if–then policy rules |
| 1990s–2000s | Wider use of statistical ML | Classify or score records |
| 2010s | Deep learning at scale | Recognise images and speech |
| 2020s | LLMs, then tool-using agents | Draft, retrieve, act and verify |

**Worked example**

- An overdue-invoice threshold starts as a rule.
- A model estimates late payment risk. A language model explains why an invoice was flagged.
- An agent retrieves the invoice and payment records and prepares a draft follow-up.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Features and next actions depend on context. | A stable, fully specified rule already solves the decision. |

**Authoritative references**

- https://proceedings.neurips.cc/paper_files/paper/2012/hash/c399862d3b9d6b76c8436e924a68c45b-Abstract.html
- https://arxiv.org/abs/2005.14165

---


### Agentic AI and AI Agents

Agentic AI is a broad term for goal-directed behaviour involving planning, action and feedback. An AI agent is a concrete worker configured for a particular goal, with a model, tools, context, state and controls.

A single agent can be invoked once or triggered on a schedule. A workflow can coordinate one or more agents with conventional rules and human approvals. More autonomy increases the need for clear permissions, stop rules and evaluation.

**How it works**

- Receive a goal and boundaries.
- Plan a permitted action using approved context.
- Call a tool and inspect its result.
- Repeat within limits, stop once complete, or escalate.

**Worked example**

- A collections agent reads a synthetic receivables export and drafts reminders for overdue accounts.
- A coordinator verifies any disputed balance and approves external sends.
- A workflow is a sequence of steps, such as matching an invoice, flagging a mismatch and routing it for review.
- The automated workflow matches invoices using preset rules. It routes unusual exceptions to finance when no rule resolves them. The accountant can then use a copilot to inspect the evidence and steer the outcome. Finance keeps approval.
- Robotic process automation (RPA) is one way to automate those steps, often through existing application interfaces. The workflow is the sequence itself. It can also be run by an integration or a person.
- A goal-directed agent can decide which approved records to inspect next, within the limits the finance team sets.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The next step changes with data or exceptions. | The task grants broad permissions with no business owner. |

**Authoritative references**

- https://www.anthropic.com/research/building-effective-agents

---


### Where Agents Add Business Value

Finance use cases include cut-off review, duplicate invoice checks, receivables follow-up, regional revenue reporting and cash reconciliation.

Start with recurring work, known source data and a human owner. Measure time saved and errors alongside the quality of the result. A good pilot limits the agent to a draft or well-scoped action and gives it a clear stop condition.

**How it works**

- Choose a recurring, time-consuming workflow.
- Define the baseline and how to check the result.
- Grant the least access required.
- Give a person responsibility for exceptions and approvals.

**Three Business Use Cases**

| Aspect | Agent action | Human gate |
|---|---|---|
| Revenue cut-off | Match invoice date to delivery evidence | Approve close adjustment |
| Payables | Flag suspected duplicate supplier invoices | Review exception before payment |
| Management reporting | Reconcile regional and product revenue | Approve management pack |

**Worked example**

- For weekly sales reporting, compare a dashboard’s totals with the source and measure analyst review time.
- For invoice matching, flag amount or vendor discrepancies to a reviewer.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Reliable inputs and a repeatable outcome exist. | Source data or the accountable reviewer is undefined. |

**Authoritative references**

- https://www.anthropic.com/research/building-effective-agents

---


### The Tema Distributor Business Model Canvas

The Business Model Canvas uses nine blocks to show who a business serves, what value it provides, how it operates and how it earns revenue. Read the framework, then use the fictional Tema wholesale distributor case to work through the completed sample.

The distributor receives about 80 sales invoices, delivery notes and returns messages a day. A finance exception takes an illustrative 18 minutes to check and route at close. The pilot tests whether an agent can prioritise exceptions and assemble evidence while accountants retain approval.

**How it works**

- Start with the customer and the desired improvement.
- Map permitted finance records, partners, channels and approval owners.
- Estimate full costs and decide what to measure before the pilot.

**Worked example**

- Key Partners — Suppliers, hauliers and the ERP provider furnish matching records.
- Key Activities — Rank close exceptions. Match orders, invoices and delivery records.
- Key Resources — Sales ledger, delivery notes, returns and approval rules.
- Value Propositions — Faster, traceable corrections to short deliveries, invoices and credits.
- Customer Relationships — Finance reviews disputes and approves credits and payments.
- Channels — Customer email and WhatsApp. Internal finance queue.
- Customer Segments — Retailers and wholesalers in Greater Accra, Ashanti and Western.
- Cost Structure — Illustrative year-one budget: GHS 39,000 for licences, integration and training. Measure time saved and review effort.
- Revenue Streams — GHS sales of staples, household and personal care products. Faster dispute resolution supports collections.

**Decision guide**

| Use when | Avoid when |
|---|---|
| You need to decide which finance task to pilot and how it supports the business. | The agent cannot trace a recommendation to approved source records. |

**Authoritative references**

- https://www.anthropic.com/research/building-effective-agents

---


### Always-On Agents in Finance Teams

An always-on agent follows approved triggers between requests. The reference deck profiles Meta Muse, OpenAI Dots, Claude Duet, Grok Bot, Manus Cue and Gemini Spark or Task. The class uses Meta Muse for inbox and calendar tracking. Product access is account- and region-dependent.

The five collaboration patterns describe how people and agents share work. Agreement on a common ticket ID, evidence source and approval owner makes the close brief traceable across hand-offs.

**How it works**

- One accountant and one agent track close deadlines.
- Several payables staff share one invoice-exception agent.
- One controller delegates extraction, matching and review to several agents.
- Multiple finance teams coordinate payables, receivables and general ledger agents.
- A buyer’s payables agent coordinates with a supplier’s receivables agent, with approval at both organisations.

**Worked example**

- For the Tema case, Muse matches demo email and calendar records by ticket ID. The finance team checks cut-off, duplicate payment and credit-note evidence.

**Decision guide**

| Use when | Avoid when |
|---|---|
| A recurring workflow has controlled inputs and a designated reviewer. | Use of the connected inbox or calendar is outside the approved scope. |

**Authoritative references**

- https://www.anthropic.com/research/building-effective-agents
- https://muse.ai/

---


### Tools, Skills and Context

A tool is a callable capability, such as reading a spreadsheet, querying a ledger or creating a chart. APIs and MCP can expose tool capabilities. A skill is a reusable set of instructions, examples and checks for when and how to perform work.

A skill tells an agent how to sequence operations and what success looks like. It uses the access the environment already provides. Memory carries approved state between tasks. Knowledge supplies records and documents to consult. Instructions give the current goal, and runtime controls govern triggers, permissions and limits.

**How it works**

- Identify what the agent must do and which tools permit it.
- Write the sequence and quality checks for a repeatable task as a skill.
- Ground the work in the current data and confirm permissions before any write.

**A Tool Acts and a Skill Guides**

| Aspect | Purpose | Datin example |
|---|---|---|
| Tool | Read, calculate or edit through an allowed capability | Insert an Excel chart |
| Skill | Store a procedure and its verification checks | Build and style a sales dashboard |
| Knowledge | Supply facts or records | Read tblDatinSales |
| Runtime | Enforce access and stop conditions | Require approval before sharing |

**Worked example**

- Excel’s chart operation can plot a linked series.
- A Datin dashboard skill says which KPIs, chart types, colours and filter tests to use.

**Decision guide**

| Use when | Avoid when |
|---|---|
| A work standard should be reused across workbooks. | The user has no permission to access the required data. |

**Authoritative references**

- https://www.anthropic.com/research/building-effective-agents
- https://support.microsoft.com/en-gb/excel/copilot/copilot-in-excel-skills

---


### Lab 1 · Meta Muse — Prioritise Month-End Exceptions with Meta Muse

Learning outcome: LO1 · Explain agent behaviour and human approval.

Goal: Use Meta Muse to track demo email and calendar deadlines and prepare a traceable month-end brief.

Read a demo Gmail inbox and Google Calendar to prepare a finance exception queue for a Tema distributor. Check cut-off, duplicate payment and credit-note evidence.

**What you'll build**

A three-row month-end exception brief showing evidence to check and approval owners.   (Tools: Meta Muse in a browser.)

**Prerequisites**

- Use muse.ai and a trainer-approved demo Gmail account and Google Calendar. The Outlook training accounts are for Excel.
- Before class, the trainer seeds three demo email subjects prefixed DATIN DEMO with ticket IDs S-104, S-109 and S-112. Put the mock amounts and exception notes from the fallback prompt in the messages. Add matching calendar events for S-104 and S-109 within seven days, S-104 first, in Accra time (GMT). Leave S-112 unscheduled.
- If connector access is unavailable, watch the trainer’s connected demonstration or use the inline mock prompt.

**Step-by-step**

1. Open muse.ai and start a new chat. Under Settings > Connectors, connect only the approved demo Gmail and Google Calendar. Use read-only permissions where offered. Run the connector check.

   ```text
   Confirm whether you can read Gmail and Google Calendar in this demo account. Tell me which connectors are available and what approval you need before an action. Do not read anything yet or change settings.
   ```

2. Ask for a read-only email and calendar brief. Confirm the returned sources and dates in the demo account.

   ```text
   Prepare a read-only month-end brief for the fictional Tema distributor. In the approved demo Gmail account find only subjects beginning DATIN DEMO. In the demo Google Calendar check only events beginning DATIN DEMO for the next seven days, in Accra time (GMT). Match the email and event by ticket ID. Rank S-104 (sales cut-off), S-109 (supplier duplicate) and S-112 (credit note) by the recorded deadline. Return the source, deadline, evidence to check and approval owner. Mark missing details. Do not send or move email, change events, post entries or make payments. Tell me whether you can repeat this as a read-only scheduled brief. Ask before creating any schedule.
   ```

3. If the connectors are unavailable, use this mock-notes variant. Check the three ticket IDs and verify which case has no deadline.

   ```text
   You are helping the finance team of a fictional wholesale distributor in Tema, Ghana, at month-end. Work only from these mock notes. Today is Tuesday. S-104: reconcile a short rice delivery against a GHS 4,800 sales invoice before the cut-off on Wednesday at 2 pm. S-109: check a possible duplicate GHS 3,250 supplier invoice before the payment review on Thursday at 10 am. S-112: check a request for a GHS 620 credit note for returned detergent, with no deadline given. Rank the three items by explicit deadline. Return ID, reason, evidence to check and one draft recommendation per item. Mark missing information and name the accounting record or tool needed. Identify who should approve a posting, payment or customer credit. Draft only. Leave accounts, connectors, records and payments untouched.
   ```

4. If scheduled tracking is supported, review the proposed demo-only scope and frequency. The trainer may approve one read-only scheduled brief for the demo, then stop it at the end of class. Compare the result with the actual demo sources.

**Sample prompts**

| Purpose | Prompt |
|---|---|
| Check email and calendar connectors | Confirm whether you can read Gmail and Google Calendar in this demo account. Tell me which connectors are available and what approval you need before an action. Do not read anything yet or change settings. |
| Track month-end email and calendar | Prepare a read-only month-end brief for the fictional Tema distributor. In the approved demo Gmail account find only subjects beginning DATIN DEMO. In the demo Google Calendar check only events beginning DATIN DEMO for the next seven days, in Accra time (GMT). Match the email and event by ticket ID. Rank S-104 (sales cut-off), S-109 (supplier duplicate) and S-112 (credit note) by the recorded deadline. Return the source, deadline, evidence to check and approval owner. Mark missing details. Do not send or move email, change events, post entries or make payments. Tell me whether you can repeat this as a read-only scheduled brief. Ask before creating any schedule. |
| Fallback · Draft from inline mock notes | You are helping the finance team of a fictional wholesale distributor in Tema, Ghana, at month-end. Work only from these mock notes. Today is Tuesday. S-104: reconcile a short rice delivery against a GHS 4,800 sales invoice before the cut-off on Wednesday at 2 pm. S-109: check a possible duplicate GHS 3,250 supplier invoice before the payment review on Thursday at 10 am. S-112: check a request for a GHS 620 credit note for returned detergent, with no deadline given. Rank the three items by explicit deadline. Return ID, reason, evidence to check and one draft recommendation per item. Mark missing information and name the accounting record or tool needed. Identify who should approve a posting, payment or customer credit. Draft only. Leave accounts, connectors, records and payments untouched. |

**Test it**

| Check | Expected result |
|---|---|
| Order and evidence | S-104 ranks before S-109. S-112 has no stated deadline. All three IDs appear. In the connected version, S-104 and S-109 have matched inbox and calendar records. |
| Tool and approval | Check proof of delivery, the AP ledger and the returns record. The finance manager approves postings, payments or credits. |

**Checkpoint and rejoin point**

Keep your close brief. Lab 2 moves from individual exceptions to Tema’s H1 sales and margin reporting. Use the independent mock workbook and reconcile it to its Data Guide.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Muse unavailable | Follow the trainer demonstration. Use the same mock brief to check the reasoning together. |
| Muse invents a due date | Ask it to cite the notes and label the missing date as unknown. |

**Challenge**

Revise the prompt to define what counts as urgent and compare the resulting priorities.

**Reflection**

Which part of this task could run on a schedule, and which decision should a person own?

> **Note:** The lab folder labs/lab-01-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (none). Use class copies of all workpapers and use the specified official report for Topic 4.

---


### Topic 01 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO1 | Use the Meta Muse brief to explain what the agent inferred, what needs a tool and what needs human approval. |

---


## Topic 02 — Create an Interactive Dashboard in Excel

Build with Copilot, reconcile sales and profit in Ghana cedi, then apply the dark theme.

**Key concepts**

- One table, one grain — tblDatinSales has one row per order and an Order_ID key.
- Use the right denominator — Gross margin divides total gross profit by total net sales in the current selection.
- Filters must control the whole view — The KPIs and all charts should respond to the same Region, Channel and Product selections.
- Protect the analytical structure — Check the same filter states before and after the theme activity.


### Sales and Financial Measures

The data represents a fictional Ghana distributor. Sales and finance fields live on the same order row. Net sales is Units × List Price − Discount. Gross profit subtracts cost of goods sold from net sales.

Gross margin is a weighted rate computed from totals. Averaging order-level percentages gives each order equal weight regardless of its value. Gross profit excludes other operating costs and taxes.

**How it works**

- Verify the table name, row grain and unique order IDs.
- Sum Net_Sales_GHS and Gross_Profit_GHS in the selected view.
- Count Order_ID values. Divide total Gross Profit by total Net Sales and format the ratio as a percent.

**Worked example**

- 60 units at GHS 100 minus GHS 300 in discounts gives GHS 5,700 net sales.
- Cost of goods sold is GHS 3,600. Gross profit is GHS 2,100 and gross margin is 36.84%.
- Lab 2 all-data baseline is GHS 1,448,550 net sales, GHS 563,084 gross profit, 38.87% margin, 432 orders.

**Decision guide**

| Use when | Avoid when |
|---|---|
| Finance and sales values use the same grain and currency. | Sales and finance tables have different grains and are joined without an aggregation plan. |

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Charts and Interactive Controls

A dashboard brings the headline metrics and comparisons onto one sheet. Use a monthly line for trends, bars for regional comparisons and doughnuts for product and channel shares of revenue. Region, Channel and Product inputs define the subset shown.

A filter that changes one chart while the KPI still shows all data makes the dashboard misleading. Slicers can work across PivotTables that share a data source. Drop-downs with formula summaries provide another option.

**How it works**

- Start from tblDatinSales and use Month_Start dates in order.
- Connect all KPI and chart summaries to the same filters.
- Start at All, select Greater Accra and Wholesale, then add Rice. Reset all selectors.

**Worked example**

- A wholesale manager selects Greater Accra and Wholesale. The KPI totals and each chart update. Adding Rice isolates one product across the selection.
- The Data Guide tab shows independent totals for that selection.

**Decision guide**

| Use when | Avoid when |
|---|---|
| A manager needs to compare subsets without editing the input table. | Users cannot tell which filters apply. |

**Authoritative references**

- https://support.microsoft.com/en-us/excel/get-started/use-slicers-to-filter-data
- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Restyle Without Breaking the Dashboard

A dark theme changes the workbook presentation. The Datin design uses navy panels, gold accents and high-contrast white labels.

Preserving formulas, chart series and filter logic keeps the visual redesign comparable with the earlier view. Legibility depends on font size, contrast and clear selected states as well as palette.

**How it works**

- Save a pre-theme copy and record selected and unfiltered totals.
- Run Prompt 2 on the dashboard sheet.
- Verify the same KPIs, charts and controls at normal zoom.

**Worked example**

- Use navy for the canvas, gold for the main chart series and white for labels.
- Compare the two control selections before and after. Check white axis labels and visibly selected filters.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The figures and interactions are already correct. | A redesign has changed formulas or the dashboard’s data sources. |

**Authoritative references**

- https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel

---


### Lab 2 · Interactive Excel dashboard — Create and Restyle an Interactive Sales Dashboard

Learning outcome: LO2 · Build and verify an interactive dashboard.

Goal: Build an interactive GHS dashboard, then apply the dark theme while preserving every result and control.

The trainer runs Prompt 1 to build the dashboard. Learners run Prompt 2 as the hands-on activity and compare the numbers and filter behavior before and after.

**What you'll build**

A new Sales Dashboard worksheet showing KPI cards, a monthly trend, regional revenue bars, doughnut charts for product and channel mix, and three working filters.   (Tools: Excel for Microsoft 365 with Copilot.)

**Prerequisites**

- Excel for Microsoft 365 and a work account with Copilot editing enabled.
- Open data/Datin_Sales_Finance_H1_GHS.xlsx as a working copy in OneDrive. Turn AutoSave on.

**Step-by-step**

1. Open the practice workbook in Excel. On the Sales Data sheet, click inside tblDatinSales. Open Copilot, select Plan if available to review the proposed changes, then allow editing to apply them.
2. Watch the trainer run Prompt 1, then run it in your own H1 copy. Check that Copilot creates a new dashboard sheet and links every chart and KPI to the data. Compare the structure with the reference video.

   ```text
   Create a professional, interactive sales dashboard using the data in this workbook. Include key sales metrics, charts showing sales trends and performance, and interactive filters so I can explore the data. Put everything on a new worksheet.
   ```

3. Inspect the dashboard. It should show Net Sales, Gross Profit, Gross Margin and Order Count, plus a monthly trend, regional bars and doughnuts for product and channel mix. If any elements are missing, run the follow-up. Test All, Greater Accra plus Wholesale, and then add Rice. Record each result.

   ```text
   On the dashboard sheet, use only tblDatinSales. Show Net Sales, Gross Profit, Gross Margin (total Gross Profit divided by total Net Sales) and Order Count (count Order_ID). Format money in GHS. Show a monthly net revenue trend, regional revenue bars, and doughnut charts for Product and Channel mix. Connect Region, Channel and Product controls to every KPI and chart. Keep the Sales Data table unchanged and tell me which controls you implemented.
   ```

4. Learner activity. Save a pre-theme copy. Select the dashboard sheet, run Prompt 2 and inspect the result at 100% zoom.

   ```text
   Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully functional. Make sure all labels and charts remain easy to read.
   ```

5. Repeat all three filter states. Check that all four KPI values match the pre-theme checks, every chart responds, and the labels and controls remain readable. Clear the filters and save a working copy.

**Sample prompts**

| Purpose | Prompt |
|---|---|
| Prompt 1 · Create the dashboard | Create a professional, interactive sales dashboard using the data in this workbook. Include key sales metrics, charts showing sales trends and performance, and interactive filters so I can explore the data. Put everything on a new worksheet. |
| Prompt 2 · Customize the theme | Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully functional. Make sure all labels and charts remain easy to read. |
| Follow-up · Complete or repair | On the dashboard sheet, use only tblDatinSales. Show Net Sales, Gross Profit, Gross Margin (total Gross Profit divided by total Net Sales) and Order Count (count Order_ID). Format money in GHS. Show a monthly net revenue trend, regional revenue bars, and doughnut charts for Product and Channel mix. Connect Region, Channel and Product controls to every KPI and chart. Keep the Sales Data table unchanged and tell me which controls you implemented. |

**Test it**

| Check | Expected result |
|---|---|
| All regions and channels | GHS 1,448,550 net sales, GHS 563,084 gross profit, 38.87% margin, 432 orders |
| Greater Accra plus Wholesale | GHS 302,775 net sales, GHS 115,670 gross profit, 38.20% margin, 72 orders |
| Greater Accra plus Wholesale plus Rice | GHS 32,080 net sales, GHS 11,629 gross profit, 36.25% margin, 12 orders |
| Dark theme | Navy panels, gold accents and readable white labels. The original layout, charts and controls still work. |

**Checkpoint and rejoin point**

Use examples/Datin_Sales_Dashboard_Example_GHS.xlsx to check the finished look and test the Region, Channel and Product drop-downs. For the theme activity, use your Prompt 1 dashboard or a trainer-provided pre-theme copy. Keep the theme and verification steps. In Lab 3 you will save them as a skill and test them with Tema’s fresh H2 orders.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Figures do not reconcile | Check the table source, GHS fields and shared filters. Use total profit divided by total sales for margin. |
| A chart ignores filters | For PivotTables sharing a source, check Slicer > Report Connections. For drop-downs, trace each chart to a filtered summary. |
| Copilot cannot edit | Check the account, OneDrive copy, AutoSave and edit mode. Use the working example for the filter demonstration. |

**Challenge**

Explain which product has the largest revenue share under Wholesale and what additional costs would be needed to assess net profit.

**Reflection**

Which dashboard tests prove a cosmetic change preserved the analysis?

> **Note:** The lab folder labs/lab-02-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (Datin_Sales_Finance_H1_GHS.xlsx). Use class copies of all workpapers and use the specified official report for Topic 4.

---


### Topic 02 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO2 | Demonstrate the same GHS totals and working filters before and after the dark theme. |

---


## Topic 03 — Skills

Convert Prompt 2 into a custom skill, install it through OneDrive and use it on fresh data.

**Key concepts**

- A skill packages a repeatable method — SKILL.md states when the skill applies, what to do and how to verify it.
- The folder name matches the skill name — Put SKILL.md inside datin-dark-sales-dashboard in the OneDrive folder opened by Excel.
- Enable it for automatic selection — Refresh and turn the custom skill on, then test with a short prompt.
- Retest on a fresh workbook — A new July–December dataset proves the skill can adapt to fresh source values.


### Convert a Prompt into a Skill

A custom Excel skill stores a name and description in YAML frontmatter at the top of SKILL.md, followed by instructions. The description guides selection. The body states the procedure and the expected result.

Prompt 2 restyles a dashboard that already exists. A short create request on a clean workbook needs a build branch before those styling instructions can run. Put both branches in the skill with a clear trigger for each.

**How it works**

- Capture the dark theme and presentation-preservation rules from Prompt 2.
- Add a create-if-missing branch that builds four KPIs, four charts and three filters.
- Add checks that work with either branch.

**Worked example**

- Restyle mode targets the dashboard that already exists. Build mode reads tblDatinSales, creates a dashboard and styles it.
- On the H2 workbook the baseline is GHS 1,732,810 net sales, GHS 665,335 gross profit, 38.40% margin, 432 orders.

**Decision guide**

| Use when | Avoid when |
|---|---|
| A dashboard standard is reused across reporting periods. | A one-off request is too specific to reuse. |

**Authoritative references**

- https://support.microsoft.com/en-gb/excel/copilot/copilot-in-excel-skills

---


### Install, Enable and Invoke the Skill

Excel discovers custom skills in a specific OneDrive folder. Excel’s Custom skills dialog opens the correct account folder. The skill subfolder name must match the SKILL.md metadata name.

Refresh asks Excel to discover the upload. Turning the skill on lets Copilot select it by default when relevant. An explicit @mention lets the learner invoke it and diagnose cases where the short prompt does not trigger it.

**How it works**

- Open the OneDrive skills folder through Excel.
- Upload the named skill folder.
- Refresh and turn on the skill.
- Try a short prompt in a fresh workbook. If needed, choose the explicit @mention and rerun from a clean copy.

**Worked example**

- Short request: Create an interactive Datin sales dashboard.
- Check the July–December dates and the filtered GHS totals before keeping the workbook.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The work account exposes Excel custom skills and supports editing. | A test would modify the only copy of a production workbook. |

**Authoritative references**

- https://support.microsoft.com/en-gb/excel/copilot/copilot-in-excel-skills

---


### Lab 3 · Excel custom skills — Install and Reuse the Dashboard Skill

Learning outcome: LO3 · Convert a prompt into an enabled skill and reuse it.

Goal: Convert Prompt 2 into a reusable skill and use a short request to create a verified dashboard on fresh data.

Save the theme and its preservation checks in SKILL.md. Add the create-if-missing steps so the same skill can apply the theme during a fresh dashboard build.

**What you'll build**

An enabled skill in your OneDrive skills folder and a new interactive dashboard based on July–December sales.   (Tools: Copilot in Excel and OneDrive.)

**Prerequisites**

- Have Prompt 2 and data/Datin_Sales_Finance_H2_GHS.xlsx open as a fresh OneDrive working copy.
- Use Excel with Manage plugins & skills available on your work account.

**Step-by-step**

1. In Copilot in Excel, use @skill-builder when available to draft a skill from Prompt 2. Review the name, trigger, theme and preservation checks. Check that the draft defines both paths: restyle an existing dashboard or build one if a create request finds none. Choose Save skill once the proposed files are correct.

   ```text
   @skill-builder Turn the theme prompt below into an Excel skill named datin-dark-sales-dashboard. For an existing interactive sales dashboard, apply the theme while preserving its layout, source data, formulas, charts and working filters. If the request is to create a dashboard and none exists, first build one from the workbook sales table with GHS Net Sales, Gross Profit, Gross Margin and Order Count, monthly trend and regional charts, doughnut charts for product and channel mix, and Region, Channel and Product filters. Apply the theme and verify the filters. Show the proposed skill for approval. Theme prompt: Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully functional. Make sure all labels and charts remain easy to read.
   ```

2. For the manual upload route, use the reviewed file or the provided skills/datin-dark-sales-dashboard/SKILL.md. In the Copilot pane choose … > Manage plugins & skills > Custom skills. Select Create OneDrive folder if needed, then Open skills folder. Upload the entire datin-dark-sales-dashboard folder, with SKILL.md inside it. If you saved through @skill-builder, confirm the file is in OneDrive and skip the manual upload.
3. Return to Custom skills and select Refresh. Find datin-dark-sales-dashboard under Manage plugins & skills and turn it on. Open the fresh workbook. Start a new Copilot chat, allow editing and send the short prompt.

   ```text
   Create an interactive Datin sales dashboard.
   ```

4. Check the theme and controls. Record the skill name if Copilot shows which skill it used. For a controlled run, or if the short request misses the theme, start from a fresh workbook copy and select the skill from + > Choose plugins and skills or use this explicit request.

   ```text
   @datin-dark-sales-dashboard Create an interactive Datin sales dashboard.
   ```

5. Check the July–December months and the three filter states below. Verify that the charts and KPI values use GHS, every chart follows the filters, and the source data is intact. Restore All and save the result.

**Sample prompts**

| Purpose | Prompt |
|---|---|
| Convert Prompt 2 | @skill-builder Turn the theme prompt below into an Excel skill named datin-dark-sales-dashboard. For an existing interactive sales dashboard, apply the theme while preserving its layout, source data, formulas, charts and working filters. If the request is to create a dashboard and none exists, first build one from the workbook sales table with GHS Net Sales, Gross Profit, Gross Margin and Order Count, monthly trend and regional charts, doughnut charts for product and channel mix, and Region, Channel and Product filters. Apply the theme and verify the filters. Show the proposed skill for approval. Theme prompt: Redesign the dashboard with a modern dark theme using navy blue backgrounds, gold accents, and high-contrast white text. Keep the existing layout, charts, metrics, and interactive filters fully functional. Make sure all labels and charts remain easy to read. |
| Run the short prompt | Create an interactive Datin sales dashboard. |
| Invoke the skill explicitly | @datin-dark-sales-dashboard Create an interactive Datin sales dashboard. |

**Test it**

| Check | Expected result |
|---|---|
| All regions and channels | GHS 1,732,810 net sales, GHS 665,335 gross profit, 38.40% margin, 432 orders |
| Greater Accra plus Wholesale | GHS 361,981 net sales, GHS 136,952 gross profit, 37.83% margin, 72 orders |
| Greater Accra plus Wholesale plus Rice | GHS 42,525 net sales, GHS 15,390 gross profit, 36.19% margin, 12 orders |
| Skill discovery | datin-dark-sales-dashboard appears after Refresh and is switched on. |
| Fresh dashboard | July–December 2026 appears in the trend. All four charts react to filters and the theme is applied. |

**Checkpoint and rejoin point**

Keep the skill folder and the completed H2 working copy. Use the fresh starter again if you want to repeat the short-prompt test.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Skill does not appear | Use Excel’s Open skills folder destination, match the folder and metadata names, check the .md extension and select Refresh. |
| Skill builder or skill menu is missing | Use the manual upload if Custom skills is available. Otherwise review the supplied file while the trainer demonstrates on an enabled account. |
| Short prompt makes a light dashboard | Choose the skill explicitly on a fresh workbook. Verify the new result and note that automatic selection may vary. |

**Challenge**

After verifying the dashboard, restyle it through the skill and check that all three filter states still return the same numbers.

**Reflection**

What must a skill specify so a short request behaves consistently on a new workbook?

> **Note:** The lab folder labs/lab-03-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (Datin_Sales_Finance_H2_GHS.xlsx), plus skills/datin-dark-sales-dashboard/SKILL.md (ready-to-upload skill). Use class copies of all workpapers and use the specified official report for Topic 4.

---


### Topic 03 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO3 | Show the enabled skill and reconcile a dashboard created from fresh data. |

---


## Topic 04 — Analyse a Financial Report with Copilot

Review the Bank of Ghana’s 2025 financial statements and document your checks in Word.

**Key concepts**

- Report scope — Keep Bank and Group, 2025 and 2024, and original GH¢’000 units separate.
- Arithmetic checks — Recompute comprehensive income, equity and the cash bridge.
- Attribution — Reconcile total Group profit or loss to its parent and non-controlling interest components.
- Evidence and review — Trace each claim to the report, record open items and approve the Word memo.


### Ground Copilot in the Published Report

Attach the official Bank of Ghana PDF in Microsoft 365 Copilot Chat. Ask for exact row labels, signed amounts, source pages and note numbers.

The PDF contains four adjacent year and entity columns. A plausible answer can still cite the wrong year, subtotal or entity.

**How it works**

- Open the report and identify the 2025 Bank and 2025 Group columns.
- Run the extraction prompt, then check each copied figure against the PDF page.
- Record printed and viewer pages when you share a citation.

**Worked example**

- The 2025 Bank operating cash flow appears on printed page 23, which is page 24 in the PDF viewer.
- Note 2(a)(i) begins on printed page 25. The auditor’s emphasis of matter refers to that accounting basis.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The official source is available and the attachment is authorised. | A cited figure cannot be located in the report or the response combines Bank and Group. |

**Authoritative references**

- https://www.bog.gov.gh/wp-content/uploads/2026/05/2025-Financial-Statements-1.pdf
- https://support.microsoft.com/en-au/microsoft-365-copilot/add-content-to-microsoft-copilot-chat-prompts
- https://support.microsoft.com/en-us/microsoft-365-copilot/file-formats-supported-by-microsoft-copilot

---


### Reconcile and Record the Evidence

Reperformance means doing the calculation independently using the published signed amounts.

Recomputation checks the maths. The accounting labels and auditor’s report explain the scope of a figure or conclusion.

**How it works**

- Recompute Bank total comprehensive income, closing equity and closing cash.
- Trace the directors’ Group loss summary to the attribution lines on the profit or loss statement.
- Review the auditor’s stated basis and key audit matter. Record verified findings and unresolved points in Word.

**Worked example**

- Assets less liabilities equals negative equity for the Bank. Compare with the Bank equity roll-forward.
- Total Group loss includes parent and non-controlling interest. Cite printed pages 5 and 15 for the label difference.

**Decision guide**

| Use when | Avoid when |
|---|---|
| The figure, formula and relevant source can be retained in the workpaper. | The task requires independent audit evidence or assurance beyond the published report. |

**Authoritative references**

- https://www.bog.gov.gh/wp-content/uploads/2026/05/2025-Financial-Statements-1.pdf

---


### Lab 4 · Copilot financial review — Review the Bank of Ghana Financial Statements

Learning outcome: LO4 · Verify Copilot findings and document the evidence.

Goal: Use the published 2025 statements to verify figures, investigate a scope difference and document evidence.

Attach the published report to Copilot Chat. Recheck its calculations against the signed statements and document your findings in Word.

**What you'll build**

A Word workpaper with a Bank-only reconciliation, a Group attribution check and an evidence-based review memo.   (Tools: Microsoft 365 Copilot Chat and Word.)

**Prerequisites**

- Download the official 2025 PDF, open the Word workpaper, and sign in to Microsoft 365 Copilot Chat.
- Use the PDF supplied for this lab. Printed page numbers are one less than the PDF viewer page positions in the cited sections.

**Step-by-step**

1. Attach the published PDF in Copilot Chat using Add content. Run the first prompt. Check the year, Bank versus Group, signed values and the GH¢’000 heading on printed pages 15–17 and 23.

   ```text
   Use only the attached Bank of Ghana Report and Financial Statements for the year ended 31 December 2025. First identify the four year and entity columns. Extract 2025 and 2024 Bank-only loss for the year, cost of open market operations and other comprehensive income. Then extract 2025 Bank-only total assets, total liabilities, total equity and cash flows. Return one row per item with exact line label, signed values, entity, year, units, printed page and note if shown. Keep figures in GH¢’000. Report a missing or unreadable figure without guessing.
   ```

2. Run the reconciliation prompt. Open the page behind each figure and recompute the arithmetic independently. Compare printed pages 5 and 15 before assigning meaning to the two Group losses.

   ```text
   Use the attached PDF as evidence. Show the formulas and printed source pages for four checks in GH¢’000. Recalculate the 2025 Bank loss plus other comprehensive income, assets less liabilities, and the opening-to-closing cash bridge including the reported exchange-rate effect. Compare Group loss in the directors’ financial results summary with total Group loss and its attribution on the 2025 statement of profit or loss. Quantify the difference and test whether non-controlling interest explains it. Keep Bank and Group separate. Flag anything that remains unresolved.
   ```

3. Open the Word workpaper and record the checked figures, evidence and open items. Briefly review the auditor’s opinion and key audit matter on printed pages 10–12. Run the memo prompt, then review and paste the accepted draft into Word.

   ```text
   Draft a short financial-statement review memo for Word using only the attached Bank of Ghana PDF and the figures I have verified. Summarise KPMG’s opinion in relation to Note 2(a)(i), its emphasis of matter and the investment securities impairment key audit matter. Add three verified findings with signed amounts in GH¢’000, printed pages and a follow-up for each open point. Separate reported facts, calculations and interpretations. State what an analytical review of public material cannot establish. Do not issue an audit opinion.
   ```


**Sample prompts**

| Purpose | Prompt |
|---|---|
| Extract with scope and sources | Use only the attached Bank of Ghana Report and Financial Statements for the year ended 31 December 2025. First identify the four year and entity columns. Extract 2025 and 2024 Bank-only loss for the year, cost of open market operations and other comprehensive income. Then extract 2025 Bank-only total assets, total liabilities, total equity and cash flows. Return one row per item with exact line label, signed values, entity, year, units, printed page and note if shown. Keep figures in GH¢’000. Report a missing or unreadable figure without guessing. |
| Reconcile and investigate | Use the attached PDF as evidence. Show the formulas and printed source pages for four checks in GH¢’000. Recalculate the 2025 Bank loss plus other comprehensive income, assets less liabilities, and the opening-to-closing cash bridge including the reported exchange-rate effect. Compare Group loss in the directors’ financial results summary with total Group loss and its attribution on the 2025 statement of profit or loss. Quantify the difference and test whether non-controlling interest explains it. Keep Bank and Group separate. Flag anything that remains unresolved. |
| Draft the review memo | Draft a short financial-statement review memo for Word using only the attached Bank of Ghana PDF and the figures I have verified. Summarise KPMG’s opinion in relation to Note 2(a)(i), its emphasis of matter and the investment securities impairment key audit matter. Add three verified findings with signed amounts in GH¢’000, printed pages and a follow-up for each open point. Separate reported facts, calculations and interpretations. State what an analytical review of public material cannot establish. Do not issue an audit opinion. |

**Test it**

| Check | Expected result |
|---|---|
| Scope | Every result identifies its entity, year, GH¢’000 units and printed source page. |
| Reperformance | Recalculate loss plus OCI, assets less liabilities and the Bank cash bridge. Tie each result to the published statement. |
| Narrative check | Reconcile total Group loss to the parent and non-controlling interest allocation. Record the follow-up on the summary label. |
| Audit report | Record the accounting basis and separate the auditor’s opinion, emphasis of matter and key audit matter. |

**Checkpoint and rejoin point**

A reviewer can reproduce the figures using the printed citations and can distinguish checked results from follow-up items.

**Troubleshooting**

| If this happens | Fix |
|---|---|
| Attachment unavailable or claims are hard to trace | Open the PDF and supply the relevant pages or selected table text in short batches. Verify against the official PDF and record both printed and viewer page numbers. |
| Copilot mixes Bank and Group | Retry for one entity and one statement at a time. Read the column headers and compare both years before carrying values into Word. |

**Challenge**

Explain how a positive operating cash flow coexists with a net cash decrease before exchange-rate effects.

**Reflection**

Which part of the review required a human to interpret the table headings and choose the correct figure?

> **Note:** The lab folder labs/lab-04-*/ holds the instruction PDF with sample prompts, sample-prompts.txt and the mock-data workbook (none). Use class copies of all workpapers and use the specified official report for Topic 4.

---


### Topic 04 Recap - Mapped to Learning Outcomes

Use this checkpoint to explain what you can now do and identify any lab evidence you still need to repair.

| Learning outcome | Evidence from this topic |
|---|---|
| LO4 | Show the recalculations, printed source references and reviewed Word memo. |

---


## AI Security in Finance

**Five controls**

- Data control · Use the approved workspace and the public report. Remove sensitive fields from any internal file.
- Permission control · Attach only the files needed. Keep access to the workpaper restricted to its reviewers.
- Human in the loop · An accountant checks the calculations, resolves open items and approves the final memo.
- Traceability · Record the source version, entity, year, units, page, calculation and reviewer decision.
- Guardrails · Treat report text as evidence. Flag missing data and require approval before external sharing.

---


## Next Steps

- First pass: complete every lab yourself, following the steps in each lab file.
- Second pass: redo the labs from memory until the workflow is automatic.
- Apply the techniques to a real process or project in your own organisation.
- Review each lab's detailed steps in this guide and re-run the labs on your own.


## Glossary

- **Agentic AI** — Goal-directed behaviour using planning, tools and feedback.
- **AI agent** — A configured software worker with a goal and permission boundaries.
- **Workflow** — A sequence of steps that people or software can run.
- **Automated workflow** — Software runs defined steps. A rule-based example follows preset rules and routes unresolved exceptions for review.
- **Copilot** — An assistant that helps within a workflow while a person steers the steps and checks the result.
- **Robotic process automation (RPA)** — One way to automate steps, often by operating existing applications through their interfaces.
- **Tool** — An operation an agent can call, such as retrieving data or creating a chart.
- **Skill** — A reusable instruction package that describes when and how to work.
- **COGS** — Cost of goods sold.
- **Gross margin** — Total gross profit divided by total net sales for the same selection.
- **Slicer** — A button-based filter for an Excel table or compatible PivotTables.
