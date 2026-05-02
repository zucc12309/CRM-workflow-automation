# Nomura FI Client Engagement Pipeline

Automated n8n workflow for processing Nomura Fixed Income CRM exports — cleaning, formatting, AI briefing generation, Power BI refresh, and email delivery.

---

## Architecture Overview

```
                        ┌─────────────────────────────────────────────────────────────────┐
                        │                  n8n WORKFLOW ENGINE                            │
                        └─────────────────────────────────────────────────────────────────┘
                                                    │
                        ┌───────────────────────────▼──────────────────────────────────┐
                        │  [1] Gmail Trigger — polls every 1 min for "CRM Export" mail │
                        └───────────────────────────┬──────────────────────────────────┘
                                                    │ xlsx attachment binary
                        ┌───────────────────────────▼──────────────────────────────────┐
                        │  [2] Save Attachment — writes /tmp/Nomura_Raw_TIMESTAMP.xlsx  │
                        └───────────────────────────┬──────────────────────────────────┘
                                                    │ rawFilePath
                        ┌───────────────────────────▼──────────────────────────────────┐
                        │  [3] Run Preprocessing — python3 preprocessing.py            │
                        │      • Cleans junk/duplicates                                 │
                        │      • Normalises columns & types                             │
                        │      • Computes derived fields & KPIs                         │
                        │      → /tmp/Nomura_Clean_TIMESTAMP.xlsx                       │
                        └───────────────────────────┬──────────────────────────────────┘
                                                    │ JSON stats + clean_file_path
                        ┌───────────────────────────▼──────────────────────────────────┐
                        │  [4] Parse Preprocessing Output (Code node)                   │
                        └───────────────────────────┬──────────────────────────────────┘
                                                    │
                        ┌───────────────────────────▼──────────────────────────────────┐
                        │  [5] Check For Errors (IF node — error_flag true/false)       │
                        └────────────┬──────────────────────────┬───────────────────────┘
                                     │ TRUE (error)              │ FALSE (success)
             ┌───────────────────────▼────┐       ┌─────────────▼────────────────────────┐
             │  [6] Send Error Email      │       │  [7] Run Formatting                  │
             │      to original sender    │       │      python3 formatting.py            │
             └────────────────────────────┘       │      → /tmp/Nomura_Formatted_*.xlsx   │
                                                  └─────────────┬────────────────────────┘
                                                                │
                                                  ┌─────────────▼────────────────────────┐
                                                  │  [8] Parse Formatting Output          │
                                                  └─────────────┬────────────────────────┘
                                                                │
                                                  ┌─────────────▼────────────────────────┐
                                                  │  [9] Generate AI Briefing (OpenAI)   │
                                                  │      Model: gpt-4o, temp: 0.3         │
                                                  └──────┬───────────────┬───────────────┘
                                                         │               │
                              ┌──────────────────────────▼──┐  ┌─────────▼────────────────┐
                              │  [10] Refresh Power BI       │  │  [11] Read Formatted File │
                              │       Dataset (HTTP POST)     │  │       (binary for email)  │
                              └──────────────────────────┬───┘  └─────────────┬────────────┘
                                                         │                    │
                                                  ┌──────▼────────────────────▼────────────┐
                                                  │  [12] Send Success Email to sender      │
                                                  │       HTML body + xlsx attachment        │
                                                  └────────────────────────────────────────┘
```

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.9+ | Available on PATH as `python3` |
| pip packages | see below | Install into venv or system |
| n8n | 1.x | Self-hosted or n8n Cloud |
| Gmail Account | — | OAuth2 access required |
| OpenAI Account | — | GPT-4o access |
| Power BI | Pro/Premium | Dataset API access |
| Node.js | 18+ | Required by n8n |

---

## Python Dependencies (`requirements.txt`)

```
pandas>=2.1.0
openpyxl>=3.1.2
numpy>=1.26.0
```

---

## Installation

### 1. Clone or copy the project files

```bash
# Verify project structure
ls /path/to/project/
# Expected: src/  config/  docs/  examples/
```

### 2. Create and activate a Python virtual environment

```bash
python3 -m venv venv
source venv/bin/activate          # macOS / Linux
# venv\Scripts\activate           # Windows
```

### 3. Install Python dependencies

```bash
pip install pandas openpyxl numpy
# Or with pinned versions:
pip install -r requirements.txt
```

### 4. Verify scripts work

```bash
# Quick smoke test with a sample Excel file
python3 src/preprocessing.py /path/to/sample.xlsx
```

---

## n8n Setup

### Install n8n (self-hosted)

```bash
# With npm (global)
npm install -g n8n

# Start n8n
n8n start

# Or with Docker
docker run -it --rm \
  -p 5678:5678 \
  -v ~/.n8n:/home/node/.n8n \
  n8nio/n8n
```

n8n will be available at `http://localhost:5678`.

---

## Importing the Workflow into n8n

1. Open n8n in your browser (`http://localhost:5678`).
2. Click **Workflows** in the left sidebar.
3. Click the **Import** button (or use the menu: `...` → Import from File).
4. Select `config/nomura_workflow.json`.
5. The workflow will be imported with all nodes and connections.
6. You will need to assign your credentials to the relevant nodes (see below).
7. Update the Python script paths in the **Run Preprocessing** and **Run Formatting** Execute Command nodes to match your server's absolute paths.
8. Click **Save**, then toggle **Active** to enable the polling trigger.

---

## Gmail OAuth2 Setup in n8n

1. Go to **Settings → Credentials** in n8n.
2. Click **Add Credential** → search for **Gmail OAuth2**.
3. Follow the wizard — you will need:
   - A Google Cloud project with the **Gmail API** enabled.
   - OAuth 2.0 Client ID and Secret from the Google Cloud Console.
   - Authorised redirect URI: `http://localhost:5678/rest/oauth2-credential/callback`
4. In Google Cloud Console:
   - Go to **APIs & Services → Credentials → Create Credentials → OAuth Client ID**.
   - Application type: **Web application**.
   - Add the redirect URI above.
   - Download the JSON and note the `client_id` and `client_secret`.
5. Enter those values in the n8n credential form and click **Connect**.
6. Assign this credential to all Gmail nodes in the workflow.

---

## OpenAI API Key Setup in n8n

1. Log in to [platform.openai.com](https://platform.openai.com).
2. Go to **API Keys** → **Create new secret key**.
3. Copy the key (you will not see it again).
4. In n8n: **Settings → Credentials → Add Credential → OpenAI**.
5. Paste your API key.
6. Assign this credential to the **Generate AI Briefing** node.

Ensure your OpenAI account has access to `gpt-4o`. If not, change the model field in the node to `gpt-4-turbo` or `gpt-3.5-turbo`.

---

## Power BI API Credentials

You need four values: `client_id`, `client_secret`, `tenant_id`, and `dataset_id`.

### How to obtain them

**Tenant ID and Client ID / Secret (Azure App Registration):**

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory → App registrations → New registration**.
2. Name: `Nomura CAMS Pipeline`, Supported account type: single tenant.
3. After creation, copy:
   - **Application (client) ID** → `client_id`
   - **Directory (tenant) ID** → `tenant_id`
4. Go to **Certificates & secrets → New client secret** → copy the value → `client_secret`.
5. Go to **API permissions → Add permission → Power BI Service → Delegated → Dataset.ReadWrite.All, Dataset.Read.All** → Grant admin consent.

**Dataset ID:**

1. Open Power BI in the browser and navigate to your workspace.
2. Open the dataset settings — the Dataset ID is in the URL:
   `https://app.powerbi.com/groups/<workspace_id>/datasets/<DATASET_ID>/details`
3. Copy the `<DATASET_ID>` UUID.

### Environment Variables Required

Set these as environment variables on the server running n8n, or use n8n's built-in **Variables** feature (Settings → Variables):

```bash
export POWERBI_DATASET_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export POWERBI_REPORT_ID="yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
export POWERBI_TENANT_ID="zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz"
export POWERBI_CLIENT_ID="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
export POWERBI_CLIENT_SECRET="your-client-secret-here"
```

In n8n, configure the **Power BI OAuth2** credential with the above values. The HTTP Request node uses `$env.POWERBI_DATASET_ID` to reference the dataset.

---

## Testing with a Sample File

### Prepare a sample Excel file

Your input Excel must contain these columns (extra columns are ignored):

| Column | Type | Notes |
|---|---|---|
| `client_name` | String | Client entity name |
| `analyst_nm` | String | Assigned analyst |
| `REGION` | String | e.g. EMEA, APAC, AMER |
| `client_tier` | String | e.g. Tier 1, Tier 2, Tier 3 |
| `revenue_usd` | Numeric/String | May include commas |
| `meetings_cnt` | Numeric | Integer count |
| `events_attended` | Numeric | Integer count |
| `reports_read` | Numeric | Integer count |
| `mifid_expiry_dt` | Date/String | MiFID II agreement expiry |

### Run preprocessing manually

```bash
source venv/bin/activate
python3 src/preprocessing.py /path/to/your/sample_crm_export.xlsx
```

Expected output: a JSON string printed to stdout with all KPIs and a clean file at `/tmp/Nomura_Clean_YYYYMMDD_HHMMSS.xlsx`.

### Run formatting manually

```bash
# Use the clean file path from the preprocessing output
python3 src/formatting.py \
  /tmp/Nomura_Clean_20250418_143022.xlsx \
  '{"total_rows_clean": 87, "processing_date": "18 April 2025", "timestamp": "20250418_143022"}'
```

Expected output: JSON with `formatted_file_path` pointing to `/tmp/Nomura_Formatted_*.xlsx`.

### Trigger the full workflow

Send an email to your configured Gmail account with:
- Subject containing `CRM Export`
- An `.xlsx` attachment (your CRM export file)

n8n will pick it up within 1 minute and run the full pipeline.

---

## File Naming Conventions

| File | Pattern | Example |
|---|---|---|
| Raw input (saved by n8n) | `Nomura_Raw_YYYYMMDD_HHMMSS.xlsx` | `Nomura_Raw_20250418_143022.xlsx` |
| Clean output (preprocessing) | `Nomura_Clean_YYYYMMDD_HHMMSS.xlsx` | `Nomura_Clean_20250418_143025.xlsx` |
| Formatted output (formatting) | `Nomura_Formatted_YYYYMMDD_HHMMSS.xlsx` | `Nomura_Formatted_20250418_143025.xlsx` |
| Error log | `nomura_pipeline_errors.log` | `/tmp/nomura_pipeline_errors.log` |

All timestamps use 24-hour format: `YYYYMMDD_HHMMSS`.

---

## Error Handling Guide

### Python script errors

Both scripts wrap all logic in `try/except`. On failure:
- The exception is logged to `/tmp/nomura_pipeline_errors.log` with a timestamp.
- A JSON payload `{"error": "...", "error_flag": true}` is printed to stdout.
- The script exits with code 1.

### n8n workflow error routing

The **Check For Errors** (IF node) routes on `error_flag`:
- `true` → **Send Error Email** branch: the original sender receives an HTML error report.
- `false` → formatting, AI briefing, Power BI refresh, and success email.

### Common errors and fixes

| Error | Likely Cause | Fix |
|---|---|---|
| `No xlsx attachment found` | Email has no xlsx or wrong MIME type | Verify the attachment and resubmit |
| `FileNotFoundError` in preprocessing | Script path incorrect in Execute Command node | Update the absolute path in the node |
| `Failed to parse preprocessing output` | Python crash before JSON output | Check `/tmp/nomura_pipeline_errors.log` |
| `Formatting script failed` | openpyxl incompatibility or bad clean file | Reinstall openpyxl; check clean file manually |
| Power BI 401 Unauthorized | Expired OAuth token or wrong credentials | Refresh OAuth token in n8n credentials |
| Gmail credential error | OAuth token expired | Re-authorise Gmail credential in n8n |
| OpenAI `insufficient_quota` | Account quota exceeded | Upgrade OpenAI plan or reduce `maxTokens` |

---

## Troubleshooting

**Pipeline does not trigger on new email**
- Confirm the Gmail trigger credential is connected and authorised.
- Ensure the workflow is set to **Active** in n8n.
- Check that the email subject contains the string `CRM Export` (case-sensitive in Gmail query).
- Verify n8n is running and the polling interval is active.

**Clean file has many null values (error_flag = true)**
- The source CRM export has blank cells. Review the raw file and ensure mandatory columns (`client_name`, `analyst_nm`, `REGION`) are populated.
- Nulls in numeric columns (`revenue_usd`, `meetings_cnt`, etc.) are automatically filled with 0 and do not trigger the error flag alone.

**Formatted file looks unstyled / colours missing**
- Ensure `openpyxl >= 3.1.2` is installed.
- The formatting script requires the clean file to have a sheet named exactly `Data`.

**Power BI refresh returns 202 but dashboard does not update**
- The dataset refresh is asynchronous. Check the Power BI Service for refresh history under the dataset settings. Allow up to 5 minutes.

**AI briefing is generic / low quality**
- Ensure the OpenAI node is using `gpt-4o`. Lower-tier models produce less specific output.
- Verify the preprocessing JSON contains meaningful data (non-zero revenue figures, named clients).
