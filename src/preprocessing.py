#!/usr/bin/env python3
"""
Nomura FI Client Engagement Pipeline — Preprocessing Script
Cleans and validates raw CRM Excel exports.
Usage: python preprocessing.py <input_file_path>
"""

import sys
import json
import logging
import traceback
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
LOG_PATH = "/Users/priyanshupatel/nomura_tmp/errors.log"
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.ERROR,
    format="%(asctime)s [PREPROCESSING] %(levelname)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# Non-capturing junk pattern — avoids pandas "match groups" warning
JUNK_RE = re.compile(r"(?i)(?:TEST|internal|dummy|System)\b")

# Currency symbols to strip from revenue values
CURRENCY_RE = re.compile(r"[$£€\s]")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_str(val):
    """Return stripped string, collapsing internal whitespace. Returns NaN for nulls."""
    if pd.isna(val):
        return np.nan
    cleaned = re.sub(r"\s+", " ", str(val).strip())
    return cleaned if cleaned else np.nan


def _fill_text(series, default):
    """Replace blank/null text cells with a stable placeholder."""
    return series.apply(_safe_str).fillna(default)


def _clean_revenue(series):
    """Strip currency symbols/commas, coerce to float, clamp negatives to 0."""
    def _parse(v):
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return 0.0
        s = CURRENCY_RE.sub("", str(v)).replace(",", "").strip()
        try:
            return max(0.0, float(s)) if s else 0.0
        except (ValueError, TypeError):
            return 0.0
    return series.apply(_parse)


def _clean_count(series):
    """Coerce to int, fill nulls with 0, clamp negatives to 0."""
    return (
        pd.to_numeric(series, errors="coerce")
        .fillna(0)
        .clip(lower=0)
        .astype(int)
    )


def compute_top_analyst(df):
    """Return analyst with highest total revenue."""
    grouped = df.groupby("analyst_nm")["revenue_usd"].sum()
    if grouped.empty:
        return None, 0.0
    top = grouped.idxmax()
    return top, float(grouped[top])


def compute_worst_analyst(df):
    """Return analyst with lowest average Rev_Per_Meeting (exclude zero-RPM rows)."""
    non_zero = df[df["Rev_Per_Meeting"] > 0].copy()
    if non_zero.empty:
        return None, 0.0
    grouped = non_zero.groupby("analyst_nm")["Rev_Per_Meeting"].mean()
    worst = grouped.idxmin()
    return worst, round(float(grouped[worst]), 2)


def build_region_summary(df):
    """Client count, total revenue, avg meetings per region (excludes UNKNOWN bucket)."""
    sub = df[df["REGION"] != "UNKNOWN"]
    if sub.empty:
        return {}
    result = {}
    for region, grp in sub.groupby("REGION"):
        result[region] = {
            "client_count":    int(len(grp)),
            "total_revenue":   round(float(grp["revenue_usd"].sum()), 2),
            "avg_meetings":    round(float(grp["meetings_cnt"].mean()), 1),
        }
    return result


def build_tier_summary(df):
    """Client count and total revenue per tier."""
    result = {}
    for tier, grp in df.groupby("client_tier"):
        result[tier] = {
            "client_count": int(len(grp)),
            "total_revenue": round(float(grp["revenue_usd"].sum()), 2),
        }
    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    import base64
    import io

    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    output_path = f"/Users/priyanshupatel/nomura_tmp/Nomura_Clean_{timestamp}.xlsx"

    required_cols = [
        "client_name",
        "analyst_nm",
        "REGION",
        "client_tier",
        "revenue_usd",
        "meetings_cnt",
        "events_attended",
        "reports_read",
        "mifid_expiry_dt",
    ]

    # ------------------------------------------------------------------
    # 1. Load the Excel file — from file path OR base64 via stdin
    # ------------------------------------------------------------------
    if len(sys.argv) >= 2 and sys.argv[1] != "-":
        raw_bytes = None
        file_source = sys.argv[1]
    else:
        raw_bytes = base64.b64decode(sys.stdin.buffer.read())
        file_source = io.BytesIO(raw_bytes)

    expected_cols = {"client_name", "analyst_nm", "REGION", "meetings_cnt", "revenue_usd"}

    def _normalize_col_name(col):
        return re.sub(r"[^a-z0-9]+", "_", str(col).strip().lower()).strip("_")

    def _canonicalize_columns(df):
        alias_map = {
            "client_name":    {"client_name", "client name", "client"},
            "analyst_nm":     {"analyst_nm", "analyst name", "analyst"},
            "REGION":         {"region"},
            "client_tier":    {"client_tier", "client tier", "tier"},
            "revenue_usd":    {"revenue_usd", "revenue usd", "revenue"},
            "meetings_cnt":   {"meetings_cnt", "meetings cnt", "meetings count", "meetings"},
            "events_attended":{"events_attended", "events attended", "events"},
            "reports_read":   {"reports_read", "reports read", "reports"},
            "mifid_expiry_dt":{"mifid_expiry_dt", "mifid expiry dt", "mifid expiry date", "mifid expiry"},
        }
        existing = {_normalize_col_name(c): c for c in df.columns}
        rename_map = {}
        for canonical, aliases in alias_map.items():
            for alias in aliases:
                key = _normalize_col_name(alias)
                if key in existing and existing[key] not in rename_map.values():
                    rename_map[existing[key]] = canonical
                    break
        return df.rename(columns=rename_map) if rename_map else df

    def _make_source():
        return file_source if raw_bytes is None else io.BytesIO(raw_bytes)

    def _read_sheet(sheet_name, header_row):
        df = pd.read_excel(_make_source(), engine="openpyxl",
                           sheet_name=sheet_name, header=header_row)
        return _canonicalize_columns(df)

    def _sheet_names():
        try:
            src = file_source if raw_bytes is None else io.BytesIO(raw_bytes)
            return pd.ExcelFile(src, engine="openpyxl").sheet_names
        except Exception:
            return []

    sheet_names = _sheet_names()
    preferred = [s for s in ["Raw_Data", "Data", "Sheet1"] if s in sheet_names]
    candidate_sheets  = preferred + [s for s in sheet_names if s not in preferred]
    candidate_headers = [0, 1, 2, 3, 4, 5]

    df_raw = None
    attempted = []
    for sheet in (candidate_sheets or [0]):
        for hdr in candidate_headers:
            try:
                df_try = _read_sheet(sheet, hdr)
            except Exception as exc:
                attempted.append(f"sheet={sheet!r} header={hdr}: {exc}")
                continue
            if expected_cols.issubset(set(df_try.columns)):
                df_raw = df_try
                break
            attempted.append(
                f"sheet={sheet!r} header={hdr}: cols={list(df_try.columns)[:10]}"
            )
        if df_raw is not None:
            break

    if df_raw is None:
        raise ValueError(
            "Missing required columns after checking all sheets and header rows. "
            f"Sheets: {sheet_names}. Attempts: {'; '.join(attempted[:8])}"
        )

    # Keep only the business columns used downstream
    df_raw = df_raw[required_cols].copy()
    total_rows_raw = len(df_raw)

    # ------------------------------------------------------------------
    # error_flag computed against RAW data before any cleaning.
    # Only flag as a hard error when > 50% of records are unusable,
    # not for individual null cells (normal in real CRM exports).
    # ------------------------------------------------------------------
    _raw_null_count = int(df_raw[[
        "client_name", "analyst_nm", "REGION", "client_tier",
        "revenue_usd", "meetings_cnt", "events_attended", "reports_read",
    ]].isnull().sum().sum())
    _raw_dup_count = int(
        df_raw.duplicated(subset=["client_name", "analyst_nm"]).sum()
    )
    # Critical: client_name or analyst_nm entirely blank = unprocessable
    _critical_blank = int(
        df_raw[["client_name", "analyst_nm"]].isnull().all(axis=1).sum()
    )
    error_flag = _critical_blank > (total_rows_raw * 0.5)

    df = df_raw.copy()

    # ------------------------------------------------------------------
    # 2. Remove junk rows (client_name AND analyst_nm)
    # ------------------------------------------------------------------
    df["client_name"] = _fill_text(df["client_name"], "Unknown Client")
    df["analyst_nm"]  = _fill_text(df["analyst_nm"],  "Unknown Analyst")
    mask_junk = (
        df["client_name"].astype(str).apply(lambda v: bool(JUNK_RE.search(v))) |
        df["analyst_nm"].astype(str).apply(lambda v: bool(JUNK_RE.search(v)))
    )
    df = df[~mask_junk].reset_index(drop=True)

    # ------------------------------------------------------------------
    # 3. Normalise & title-case text BEFORE dedup so casing variants collapse
    # ------------------------------------------------------------------
    df["analyst_nm"]  = df["analyst_nm"]  # already filled above
    df["REGION"]      = _fill_text(df["REGION"],       "UNKNOWN")
    df["client_tier"] = _fill_text(df["client_tier"],  "Unspecified")

    for col in ["client_name", "analyst_nm"]:
        df[col] = df[col].apply(
            lambda v: v.title() if isinstance(v, str) else v
        )

    # ------------------------------------------------------------------
    # 4. Remove duplicates (client_name + analyst_nm), keep first occurrence
    # ------------------------------------------------------------------
    df = df.drop_duplicates(
        subset=["client_name", "analyst_nm"], keep="first"
    ).reset_index(drop=True)

    # ------------------------------------------------------------------
    # 5. Uppercase REGION, title-case client_tier
    # ------------------------------------------------------------------
    df["REGION"] = df["REGION"].apply(
        lambda v: v.upper() if isinstance(v, str) else v
    )
    df["client_tier"] = df["client_tier"].apply(
        lambda v: v.title() if isinstance(v, str) else v
    )

    # ------------------------------------------------------------------
    # 6. Clean revenue_usd — strip currency symbols, clamp negatives
    # ------------------------------------------------------------------
    df["revenue_usd"] = _clean_revenue(df["revenue_usd"])

    # ------------------------------------------------------------------
    # 7. Clean count columns — coerce, fill nulls, clamp negatives
    # ------------------------------------------------------------------
    for col in ["meetings_cnt", "events_attended", "reports_read"]:
        df[col] = _clean_count(df[col])

    # ------------------------------------------------------------------
    # 8. Parse mifid_expiry_dt (dayfirst=True for EU/UK CRM date formats)
    # ------------------------------------------------------------------
    df["mifid_expiry_dt"] = pd.to_datetime(
        df["mifid_expiry_dt"], dayfirst=True, errors="coerce"
    )

    # ------------------------------------------------------------------
    # 9. MiFID_Status
    #    EXPIRED  — date exists and is in the past
    #    Active   — date exists and is today or future
    #    Unknown  — no date recorded (compliance risk, NOT assumed active)
    # ------------------------------------------------------------------
    today = pd.Timestamp(now.date())
    df["MiFID_Status"] = df["mifid_expiry_dt"].apply(
        lambda d: "EXPIRED" if pd.notna(d) and d < today
        else "Active"  if pd.notna(d)
        else "Unknown"
    )

    # ------------------------------------------------------------------
    # 10. Rev_Per_Meeting
    # ------------------------------------------------------------------
    df["Rev_Per_Meeting"] = df.apply(
        lambda r: round(r["revenue_usd"] / r["meetings_cnt"], 2)
        if r["meetings_cnt"] > 0 else 0.0,
        axis=1,
    )

    # ------------------------------------------------------------------
    # 11. Engagement_Score = meetings + events + (reports × 0.5)
    # ------------------------------------------------------------------
    df["Engagement_Score"] = (
        df["meetings_cnt"]
        + df["events_attended"]
        + (df["reports_read"] * 0.5)
    )

    # ------------------------------------------------------------------
    # 12. Flag — high meetings, low revenue
    # ------------------------------------------------------------------
    df["Flag"] = df.apply(
        lambda r: "Review"
        if r["meetings_cnt"] >= 10 and r["revenue_usd"] < 100_000
        else "OK",
        axis=1,
    )

    # ------------------------------------------------------------------
    # 13. JDF-aligned helper columns
    #     Research_Access_Status: Revoked=EXPIRED, Pending=Unknown, Granted=Active
    #     Client_Focus_Category:  Focus List = flagged for senior review
    #     Readership_Rate:        reports per meeting
    # ------------------------------------------------------------------
    _mifid_to_access = {
        "EXPIRED": "Revoked",
        "Unknown": "Pending Review",
        "Active":  "Granted",
    }
    df["Research_Access_Status"] = df["MiFID_Status"].map(_mifid_to_access).fillna("Pending Review")
    df["Client_Focus_Category"]  = df["Flag"].apply(
        lambda s: "Focus List" if s == "Review" else "Standard"
    )
    df["Readership_Rate"] = df.apply(
        lambda r: round(r["reports_read"] / r["meetings_cnt"], 2)
        if r["meetings_cnt"] > 0 else 0.0,
        axis=1,
    )

    # ------------------------------------------------------------------
    # 14. Summary statistics
    # ------------------------------------------------------------------
    total_rows_clean  = len(df)
    rows_removed      = total_rows_raw - total_rows_clean

    # Post-clean null count (should be 0 after fills — confirms pipeline integrity)
    null_count = int(df[[
        "client_name", "analyst_nm", "REGION", "client_tier",
        "revenue_usd", "meetings_cnt", "events_attended", "reports_read",
    ]].isnull().sum().sum())

    duplicate_count              = int(df.duplicated(subset=["client_name", "analyst_nm"]).sum())
    missing_mifid_count          = int(df["mifid_expiry_dt"].isna().sum())
    unknown_mifid_count          = int((df["MiFID_Status"] == "Unknown").sum())
    expired_mifid_count          = int((df["MiFID_Status"] == "EXPIRED").sum())
    flagged_clients_count        = int((df["Flag"] == "Review").sum())
    research_access_granted_count= int((df["Research_Access_Status"] == "Granted").sum())
    research_access_revoked_count= int((df["Research_Access_Status"] == "Revoked").sum())
    research_access_pending_count= int((df["Research_Access_Status"] == "Pending Review").sum())
    focus_list_count             = int((df["Client_Focus_Category"] == "Focus List").sum())
    avg_readership_rate          = round(float(df["Readership_Rate"].mean()), 2) if len(df) else 0.0

    top_analyst_name,  top_analyst_revenue = compute_top_analyst(df)
    worst_analyst_name, worst_analyst_rpm  = compute_worst_analyst(df)

    flagged_clients_list = (
        df[df["Flag"] == "Review"][[
            "client_name", "analyst_nm", "meetings_cnt", "revenue_usd", "MiFID_Status"
        ]]
        .assign(
            meetings_cnt=lambda x: x["meetings_cnt"].astype(int),
            revenue_usd =lambda x: x["revenue_usd"].round(2),
        )
        .to_dict(orient="records")
    )

    expired_mifid_list = (
        df[df["MiFID_Status"] == "EXPIRED"][[
            "client_name", "analyst_nm", "mifid_expiry_dt"
        ]]
        .assign(
            mifid_expiry_dt=lambda x: x["mifid_expiry_dt"].dt.strftime("%Y-%m-%d")
        )
        .to_dict(orient="records")
    )

    top_3_clients = (
        df.nlargest(3, "revenue_usd")[["client_name", "revenue_usd"]]
        .round({"revenue_usd": 2})
        .to_dict(orient="records")
    )

    bottom_3_rpm_clients = (
        df[df["Rev_Per_Meeting"] > 0]
        .nsmallest(3, "Rev_Per_Meeting")[["client_name", "Rev_Per_Meeting"]]
        .to_dict(orient="records")
    )

    region_summary = build_region_summary(df)
    tier_summary   = build_tier_summary(df)
    processing_date = now.strftime("%d %B %Y")

    # ------------------------------------------------------------------
    # 15. Save clean file
    #     mifid_expiry_dt serialised as ISO string; NaT → empty string
    # ------------------------------------------------------------------
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df_to_save = df.copy()
    df_to_save["mifid_expiry_dt"] = (
        df_to_save["mifid_expiry_dt"]
        .dt.strftime("%Y-%m-%d")
        .fillna("")
    )
    df_to_save.to_excel(output_path, index=False, engine="openpyxl", sheet_name="Data")

    # ------------------------------------------------------------------
    # 16. Output result JSON
    # ------------------------------------------------------------------
    result = {
        "clean_file_path":               output_path,
        "total_rows_raw":                total_rows_raw,
        "total_rows_clean":              total_rows_clean,
        "rows_removed":                  rows_removed,
        "null_count":                    null_count,
        "duplicate_count":               duplicate_count,
        "missing_mifid_count":           missing_mifid_count,
        "unknown_mifid_count":           unknown_mifid_count,
        "expired_mifid_count":           expired_mifid_count,
        "flagged_clients_count":         flagged_clients_count,
        "research_access_granted_count": research_access_granted_count,
        "research_access_revoked_count": research_access_revoked_count,
        "research_access_pending_count": research_access_pending_count,
        "focus_list_count":              focus_list_count,
        "avg_readership_rate":           avg_readership_rate,
        "error_flag":                    error_flag,
        "top_analyst_name":              top_analyst_name,
        "top_analyst_revenue":           round(top_analyst_revenue, 2),
        "worst_analyst_name":            worst_analyst_name,
        "worst_analyst_rpm":             worst_analyst_rpm,
        "flagged_clients_list":          flagged_clients_list,
        "expired_mifid_list":            expired_mifid_list,
        "top_3_clients":                 top_3_clients,
        "bottom_3_rpm_clients":          bottom_3_rpm_clients,
        "region_summary":                region_summary,
        "tier_summary":                  tier_summary,
        "processing_date":               processing_date,
        "timestamp":                     timestamp,
    }

    print(json.dumps(result))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        error_detail = traceback.format_exc()
        logging.error("Preprocessing failed:\n%s", error_detail)
        print(json.dumps({
            "error":        str(exc),
            "error_flag":   True,
            "error_detail": error_detail,
        }))
        sys.exit(1)
