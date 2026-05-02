Prompt for Claude Code:

Build a complete n8n workflow automation system for a Nomura Fixed Income client engagement pipeline. Here are all the details:

OVERVIEW:
An email is received with a raw CRM Excel file attached. The workflow automatically preprocesses it with Python, formats it in Excel, generates AI insights using the OpenAI API, refreshes a Power BI dashboard, and replies to the original email with the clean file + dashboard link + AI summary. Everything runs without human intervention.

---

TECH STACK:
- n8n (self-hosted or cloud) for workflow orchestration
- Python (pandas, openpyxl) for preprocessing and Excel formatting
- OpenAI API (gpt-4o) for generating the management summary
- Power BI REST API for dataset refresh
- Gmail or Outlook via n8n nodes for email trigger and reply

---

THE RAW DATA (what comes in via email attachment):
An Excel file called Nomura_Raw_Export.xlsx with these columns:
  client_name       — text, has trailing/leading spaces, mixed case, duplicates
  analyst_nm        — text, sometimes lowercase
  REGION            — text, values: EMEA / APAC / AMER, sometimes lowercase
  meetings_cnt      — integer, has missing values (blank cells)
  events_attended   — integer, has missing values
  reports_read      — integer, has missing values
  revenue_usd       — number, sometimes stored as text with commas e.g. "75,000", has blanks
  mifid_expiry_dt   — date, mixed formats: YYYY-MM-DD and DD-MM-YYYY
  client_tier       — text, values: Platinum / Gold / Standard, inconsistent casing
  last_contact_dt   — date string, mostly consistent
  notes_col         — free text, ignore for analysis
  internal_flag     — text, DUPLICATE or REMOVE flags on junk rows

Total rows vary every month — could be 80 rows, could be 200 rows. The Python script must handle any number dynamically, never hardcode row counts.

---

PYTHON PREPROCESSING SCRIPT (n8n Execute Code node):
Write a complete Python script that does ALL of the following in order:

1. Load the Excel file from the path passed by n8n
2. Remove junk rows where client_name contains: TEST, internal, dummy, System (case-insensitive)
3. Remove duplicate rows based on client_name + analyst_nm combination, keep first
4. Strip whitespace and apply title case to: client_name, analyst_nm
5. Force uppercase on REGION column
6. Apply title case to client_tier
7. Fix revenue_usd: remove commas, convert to numeric, fill blanks with 0
8. Fix meetings_cnt, events_attended, reports_read: convert to numeric, fill blanks with 0, cast to int
9. Fix mifid_expiry_dt: parse with pd.to_datetime(dayfirst=False, errors='coerce')
10. Add MiFID_Status column: 'EXPIRED' if expiry < today, else 'Active'
11. Add Rev_Per_Meeting column: revenue_usd / meetings_cnt, 0 if meetings = 0, rounded to 2dp
12. Add Engagement_Score column: meetings_cnt + events_attended + (reports_read * 0.5)
13. Add Flag column: 'Review' if meetings_cnt >= 10 AND revenue_usd < 100000, else 'OK'
14. Run validation checks and return results:
    - total_rows_raw (before cleaning)
    - total_rows_clean (after cleaning)
    - rows_removed (difference)
    - null_count (total remaining nulls)
    - duplicate_count (remaining duplicates)
    - expired_mifid_count (count of EXPIRED rows)
    - flagged_clients_count (count of Review rows)
    - error_flag: True if null_count > 0 or duplicate_count > 0, else False
15. Save clean file as Nomura_Clean_[timestamp].xlsx
16. Return all validation results as a JSON object for n8n to use downstream

---

EXCEL FORMATTING (second Python script via n8n Execute Code node):
Using openpyxl, format the clean Excel file:

1. Add a title banner row at the top: "NOMURA FI CLIENT ENGAGEMENT REPORT | Processed: [timestamp] | [total_rows_clean] records"
   - Dark navy background (#1F3864), white bold Arial font, merged across all columns
2. Style the header row: mid blue background (#2E75B6), white bold text, centered, thin borders
3. Alternate row colours: white and light blue (#EBF3FB)
4. Add thin borders to all data cells
5. Conditional formatting:
   - MiFID_Status = EXPIRED → red fill (#FFCCCC), red bold text
   - Flag = Review → amber fill (#FFF3CD), amber bold text
   - Rev_Per_Meeting < 5000 → light red fill on that cell only
6. Auto-fit column widths (set sensible widths per column)
7. Freeze panes at row 3 (below title and header)
8. Add a second sheet called "Analyst_Summary" with:
   - Pivot-style table: analyst name, client count, total revenue, total meetings, avg rev per meeting
   - Sorted by avg rev per meeting ascending (worst efficiency first)
   - Same formatting style as main sheet
9. Add a third sheet called "Flagged_Clients" containing only rows where Flag = Review
   - With a red title banner saying "CLIENTS REQUIRING SENIOR MANAGEMENT REVIEW"
10. Save as Nomura_Formatted_[timestamp].xlsx

---

CHATGPT INSIGHTS (n8n OpenAI node):
Use model: gpt-4o
Temperature: 0.3 (keep it factual and consistent)

Build the prompt dynamically using the validation results and computed stats from Python. The prompt must include ALL of the following computed values which n8n passes as variables:

  {{total_rows_clean}}      — total clean client records this month
  {{rows_removed}}          — how many rows were removed during cleaning
  {{expired_mifid_count}}   — number of clients with expired MiFID agreements
  {{flagged_clients_count}} — number of clients flagged for review
  {{top_analyst_name}}      — analyst with highest total revenue
  {{top_analyst_revenue}}   — their total revenue figure
  {{worst_analyst_name}}    — analyst with lowest avg rev per meeting
  {{worst_analyst_rpm}}     — their avg rev per meeting figure
  {{flagged_clients_list}}  — JSON array of flagged client objects with: client_name, analyst_nm, meetings_cnt, revenue_usd, MiFID_Status
  {{expired_mifid_list}}    — JSON array of expired MiFID clients with: client_name, analyst_nm, mifid_expiry_dt
  {{top_3_clients}}         — top 3 clients by revenue: name + revenue
  {{bottom_3_rpm_clients}}  — bottom 3 clients by rev per meeting: name + RPM value
  {{region_summary}}        — JSON: total revenue per region (EMEA, APAC, AMER)
  {{tier_summary}}          — JSON: client count and total revenue per tier (Platinum, Gold, Standard)
  {{processing_date}}       — today's date formatted as DD Month YYYY

The system prompt for ChatGPT should be:
"You are a senior Fixed Income research analyst assistant at Nomura. You write concise, professional management briefing notes for the Chief Administrative Officer of Global Markets. Your tone is direct, data-driven, and actionable. Never use bullet points — write in short, crisp paragraphs. Flag urgent items clearly."

The user prompt should be:
"Write a complete management briefing note for the Fixed Income client engagement report processed on {{processing_date}}.

Dataset: {{total_rows_clean}} client records processed this month ({{rows_removed}} removed during cleaning).

ANALYST PERFORMANCE:
Best performing analyst by total revenue: {{top_analyst_name}} with ${{top_analyst_revenue}}.
Least efficient analyst by revenue per meeting: {{worst_analyst_name}} averaging ${{worst_analyst_rpm}} per meeting.

REGIONAL BREAKDOWN:
{{region_summary}}

TIER BREAKDOWN:
{{tier_summary}}

TOP CLIENTS BY REVENUE:
{{top_3_clients}}

CLIENTS REQUIRING REVIEW (high meetings, low revenue):
{{flagged_clients_list}}
There are {{flagged_clients_count}} clients flagged this month.

COMPLIANCE — MiFID EXPIRIES:
{{expired_mifid_list}}
There are {{expired_mifid_count}} clients with expired MiFID research agreements.

LOWEST EFFICIENCY CLIENTS (revenue per meeting):
{{bottom_3_rpm_clients}}

Write the briefing note with these sections:
1. Executive summary (2-3 sentences covering overall portfolio health this month)
2. Analyst performance commentary (who is efficient, who needs review, be specific with numbers)
3. Client risk flags (name each flagged client, their analyst, why they are flagged, recommended action)
4. Compliance actions required (list each expired MiFID client, when it expired, urgency level)
5. Regional and tier insights (where is revenue concentrated, any imbalances)
6. Recommended actions for this week (3-4 specific, numbered actions senior management should take)

End with a one-line sign-off: 'Prepared by Nomura CAMS Automated Analytics | {{processing_date}}'"

---

N8N WORKFLOW STRUCTURE:
Build the complete n8n workflow JSON with these nodes in order:

Node 1 — Gmail Trigger
  - Trigger: new email received with attachment
  - Filter: only process emails where subject contains "CRM Export" or attachment is .xlsx
  - Extract: sender email address (store for reply), attachment binary data, email subject

Node 2 — Save attachment
  - Write the attachment to local disk as /tmp/Nomura_Raw_[timestamp].xlsx
  - Pass the file path to next node

Node 3 — Execute Python: Preprocessing
  - Run the full preprocessing script above
  - Pass output JSON (validation results + computed stats) to next nodes
  - Pass clean file path to next node

Node 4 — IF node: Error check
  - Condition: error_flag === true (nulls or duplicates remain)
  - TRUE branch → Node 4a (error email)
  - FALSE branch → Node 5 (continue)

Node 4a — Gmail: Send error email
  - Reply to original sender
  - Subject: "RE: {{original_subject}} — Processing Error"
  - Body: "Your CRM export could not be fully processed. Issues found: [null_count] null values and [duplicate_count] duplicate records remain after cleaning. Please review the attached error log and resubmit."
  - Attach: error log file

Node 5 — Execute Python: Excel formatting
  - Run the formatting script above on the clean file
  - Output: formatted file path

Node 6 — OpenAI: Generate insights
  - Use the dynamic prompt above
  - All {{variables}} populated from Node 3 output
  - Store GPT response as ai_summary variable

Node 7 — HTTP Request: Power BI refresh
  - POST to Power BI REST API: https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/refreshes
  - Auth: Bearer token (OAuth2)
  - Body: {"notifyOption": "MailOnCompletion"}
  - Store response status

Node 8 — Gmail: Send success reply
  - Reply to original sender email address from Node 1
  - Subject: "RE: {{original_subject}} — Clean Report Ready"
  - Body (HTML formatted):
      Header: "NOMURA FI CLIENT ENGAGEMENT — AUTOMATED REPORT"
      Section 1: Processing summary (rows processed, rows removed, date)
      Section 2: The full ChatGPT briefing note (ai_summary variable)
      Section 3: "Your formatted Excel report is attached. The Power BI dashboard has been refreshed and is available at: [dashboard_link]"
      Footer: "This report was generated automatically by the Nomura CAMS Analytics Pipeline"
  - Attach: Nomura_Formatted_[timestamp].xlsx
  - Format: HTML email, professional styling

---

ERROR HANDLING:
- Every Python script must have try/except blocks that catch errors and return them as JSON
- If Power BI refresh fails, still send the email but note "Dashboard refresh pending" instead of the link
- If OpenAI API fails, still send the email but replace ai_summary with "AI insights unavailable — please review attached report manually"
- Log all errors with timestamps to /tmp/nomura_pipeline_errors.log

---

FILE NAMING CONVENTION:
All output files use timestamp format: YYYYMMDD_HHMMSS
Example: Nomura_Clean_20260418_143022.xlsx

---

DELIVERABLES — produce all of the following:
1. The complete Python preprocessing script (standalone .py file)
2. The complete Python Excel formatting script (standalone .py file)
3. The complete n8n workflow as an exportable JSON file (can be imported directly into n8n)
4. A README.md explaining how to set up and run the workflow, including:
   - Required Python packages (requirements.txt)
   - n8n node configuration steps
   - How to get Power BI API credentials
   - How to set up Gmail OAuth in n8n
   - How to set the OpenAI API key in n8n credentials
5. A sample of what the final email output looks like (as an HTML file)







