# Lab 1 — Prioritise Month-End Exceptions with Meta Muse

[Instructions PDF](Lab-01-Instructions.pdf) · [Copy prompts](../../prompts/index.html#lab-1)

## Goal

Use Meta Muse to track demo email and calendar deadlines and prepare a traceable month-end brief.

## Prerequisites

- Use muse.ai and a trainer-approved demo Gmail account and Google Calendar. The Outlook training accounts are for Excel.
- Before class, the trainer seeds three demo email subjects prefixed DATIN DEMO with ticket IDs S-104, S-109 and S-112. Put the mock amounts and exception notes from the fallback prompt in the messages. Add matching calendar events for S-104 and S-109 within seven days, S-104 first, in Accra time (GMT). Leave S-112 unscheduled.
- If connector access is unavailable, watch the trainer’s connected demonstration or use the inline mock prompt.

## Steps

### 1. Open muse.ai and start a new chat. Under Settings > Connectors, connect only the approved demo Gmail and Google Calendar. Use read-only permissions where offered. Run the connector check.

```text
Confirm whether you can read Gmail and Google Calendar in this demo account. Tell me which connectors are available and what approval you need before an action. Do not read anything yet or change settings.
```

### 2. Ask for a read-only email and calendar brief. Confirm the returned sources and dates in the demo account.

```text
Prepare a read-only month-end brief for the fictional Tema distributor. In the approved demo Gmail account find only subjects beginning DATIN DEMO. In the demo Google Calendar check only events beginning DATIN DEMO for the next seven days, in Accra time (GMT). Match the email and event by ticket ID. Rank S-104 (sales cut-off), S-109 (supplier duplicate) and S-112 (credit note) by the recorded deadline. Return the source, deadline, evidence to check and approval owner. Mark missing details. Do not send or move email, change events, post entries or make payments. Tell me whether you can repeat this as a read-only scheduled brief. Ask before creating any schedule.
```

### 3. If the connectors are unavailable, use this mock-notes variant. Check the three ticket IDs and verify which case has no deadline.

```text
You are helping the finance team of a fictional wholesale distributor in Tema, Ghana, at month-end. Work only from these mock notes. Today is Tuesday. S-104: reconcile a short rice delivery against a GHS 4,800 sales invoice before the cut-off on Wednesday at 2 pm. S-109: check a possible duplicate GHS 3,250 supplier invoice before the payment review on Thursday at 10 am. S-112: check a request for a GHS 620 credit note for returned detergent, with no deadline given. Rank the three items by explicit deadline. Return ID, reason, evidence to check and one draft recommendation per item. Mark missing information and name the accounting record or tool needed. Identify who should approve a posting, payment or customer credit. Draft only. Leave accounts, connectors, records and payments untouched.
```

### 4. If scheduled tracking is supported, review the proposed demo-only scope and frequency. The trainer may approve one read-only scheduled brief for the demo, then stop it at the end of class. Compare the result with the actual demo sources.

## Test it

- **Order and evidence** · S-104 ranks before S-109. S-112 has no stated deadline. All three IDs appear. In the connected version, S-104 and S-109 have matched inbox and calendar records.
- **Tool and approval** · Check proof of delivery, the AP ledger and the returns record. The finance manager approves postings, payments or credits.

## Checkpoint

Keep your close brief. Lab 2 moves from individual exceptions to Tema’s H1 sales and margin reporting. Use the independent mock workbook and reconcile it to its Data Guide.

## Files

- [Open Muse](https://muse.ai/)
