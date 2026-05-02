"""
Nomura FI – Power BI Template (.pbit) generator.
Usage:  python3 powerbi_export.py <clean_excel_path> [output_dir]

Produces:
  Nomura_Dashboard.pbit   – open in Power BI Desktop, refresh → all data loads
  Nomura_PowerBI_Data.xlsx – the data source the template points to

The template ships with pre-built DAX measures and all columns correctly
typed so the user only needs to drag fields onto the blank canvas.
"""

import json
import sys
import uuid
import zipfile
from pathlib import Path

import pandas as pd


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def load(path: str) -> pd.DataFrame:
    xf = pd.ExcelFile(path, engine="openpyxl")
    sheet = "Data" if "Data" in xf.sheet_names else xf.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    num_cols = ["revenue_usd", "meetings_cnt", "events_attended", "reports_read",
                "Rev_Per_Meeting", "Engagement_Score", "Readership_Rate"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    str_cols = ["client_name", "analyst_nm", "REGION", "client_tier",
                "MiFID_Status", "Flag", "Research_Access_Status",
                "Client_Focus_Category", "mifid_expiry_dt"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df


def export_xlsx(df: pd.DataFrame, out_path: Path) -> None:
    df.to_excel(out_path, index=False, engine="openpyxl")


# ---------------------------------------------------------------------------
# PBIT component builders
# ---------------------------------------------------------------------------

CONTENT_TYPES = """<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Default Extension="xml" ContentType="application/xml" />
  <Override PartName="/Version" ContentType="application/octet-stream" />
  <Override PartName="/SecurityBindings" ContentType="application/octet-stream" />
  <Override PartName="/DataModelSchema" ContentType="application/json" />
  <Override PartName="/Settings" ContentType="application/json" />
  <Override PartName="/Report/Layout" ContentType="application/json" />
</Types>"""

RELS = """<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://schemas.microsoft.com/powerbi/2016/report" Target="/Report/Layout" Id="rId1"/>
  <Relationship Type="http://schemas.microsoft.com/powerbi/2016/model" Target="/DataModelSchema" Id="rId2"/>
  <Relationship Type="http://schemas.microsoft.com/powerbi/2016/settings" Target="/Settings" Id="rId3"/>
</Relationships>"""

VERSION = "1.0".encode("utf-16-le")   # Power BI reads Version as UTF-16 LE

SETTINGS = json.dumps({
    "Version": 1,
    "ExplicitMeasures": True,
    "RelationshipImportSettings": {
        "RelationshipImportMode": "ActiveOnly"
    }
})

METADATA = json.dumps({
    "Version": 1,
    "ReportId": str(uuid.uuid4()),
    "ContentExpirationEnabled": False,
    "IsOwnedByMe": True
})

DIAGRAM_LAYOUT = json.dumps({
    "version": 1,
    "diagrams": [
        {
            "ordinal": 0,
            "scrollPosition": {"x": 0, "y": 0},
            "nodes": [
                {
                    "nodeIndex": "Clients",
                    "size": {"width": 238, "height": 418},
                    "position": {"x": 100, "y": 100}
                }
            ]
        }
    ]
})


def m_expression(xlsx_path: str) -> list[str]:
    """Power Query M expression to load and type the Excel file."""
    escaped = xlsx_path.replace("\\", "\\\\").replace('"', '\\"')
    type_list = (
        '{{\"client_name\", type text}}, {{\"analyst_nm\", type text}}, '
        '{{\"REGION\", type text}}, {{\"client_tier\", type text}}, '
        '{{\"revenue_usd\", type number}}, {{\"meetings_cnt\", Int64.Type}}, '
        '{{\"events_attended\", Int64.Type}}, {{\"reports_read\", Int64.Type}}, '
        '{{\"mifid_expiry_dt\", type text}}, {{\"MiFID_Status\", type text}}, '
        '{{\"Rev_Per_Meeting\", type number}}, {{\"Engagement_Score\", type number}}, '
        '{{\"Flag\", type text}}, {{\"Research_Access_Status\", type text}}, '
        '{{\"Client_Focus_Category\", type text}}, {{\"Readership_Rate\", type number}}'
    )
    return [
        "let",
        f'    Source = Excel.Workbook(File.Contents("{escaped}"), null, true),',
        '    Data_Sheet = Source{[Item="Data",Kind="Sheet"]}[Data],',
        '    #"Promoted Headers" = Table.PromoteHeaders(Data_Sheet, [PromoteAllScalars=true]),',
        f'    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {{{type_list}}})',
        "in",
        '    #"Changed Type"',
    ]


def make_column(name: str, data_type: str, fmt: str = "", summarize: str = "none") -> dict:
    col = {
        "name": name,
        "dataType": data_type,
        "lineageTag": str(uuid.uuid4()),
        "sourceColumn": name,
        "summarizeBy": summarize,
        "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}],
    }
    if fmt:
        col["formatString"] = fmt
    return col


def make_measure(name: str, expression: str, fmt: str = "") -> dict:
    m = {
        "name": name,
        "expression": expression,
        "lineageTag": str(uuid.uuid4()),
        "annotations": [{"name": "PBI_FormatHint", "value": '{"isGeneralNumber":true}'}],
    }
    if fmt:
        m["formatString"] = fmt
        m["annotations"] = []
    return m


def data_model_schema(csv_path: str) -> str:
    columns = [
        make_column("client_name",          "string"),
        make_column("analyst_nm",           "string"),
        make_column("REGION",               "string"),
        make_column("client_tier",          "string"),
        make_column("revenue_usd",          "decimal", "\\$#,0.00", "sum"),
        make_column("meetings_cnt",         "int64",   "#,0",       "sum"),
        make_column("events_attended",      "int64",   "#,0",       "sum"),
        make_column("reports_read",         "int64",   "#,0",       "sum"),
        make_column("mifid_expiry_dt",      "string"),
        make_column("MiFID_Status",         "string"),
        make_column("Rev_Per_Meeting",      "decimal", "\\$#,0.00", "average"),
        make_column("Engagement_Score",     "decimal", "0.00",      "average"),
        make_column("Flag",                 "string"),
        make_column("Research_Access_Status", "string"),
        make_column("Client_Focus_Category",  "string"),
        make_column("Readership_Rate",      "decimal", "0.00%",     "average"),
    ]

    measures = [
        make_measure("Total Revenue",
                     'SUM(Clients[revenue_usd])',
                     "\\$#,0.00"),
        make_measure("Total Clients",
                     'COUNTROWS(Clients)',
                     "#,0"),
        make_measure("Total Meetings",
                     'SUM(Clients[meetings_cnt])',
                     "#,0"),
        make_measure("MiFID Expired",
                     'CALCULATE(COUNTROWS(Clients), Clients[MiFID_Status]="EXPIRED")',
                     "#,0"),
        make_measure("Flagged Clients",
                     'CALCULATE(COUNTROWS(Clients), Clients[Flag]="Review")',
                     "#,0"),
        make_measure("Avg Rev Per Meeting",
                     'AVERAGEX(FILTER(Clients, Clients[meetings_cnt]>0), Clients[Rev_Per_Meeting])',
                     "\\$#,0.00"),
        make_measure("Avg Engagement Score",
                     'AVERAGE(Clients[Engagement_Score])',
                     "0.0"),
        make_measure("Revenue per Client",
                     'DIVIDE([Total Revenue], [Total Clients], 0)',
                     "\\$#,0.00"),
        make_measure("MiFID Expired %",
                     'DIVIDE([MiFID Expired], [Total Clients], 0)',
                     "0.0%"),
        make_measure("Flagged %",
                     'DIVIDE([Flagged Clients], [Total Clients], 0)',
                     "0.0%"),
    ]

    schema = {
        "name": "Model",
        "compatibilityLevel": 1480,
        "model": {
            "culture": "en-US",
            "dataAccessOptions": {
                "legacyRedirects": True,
                "returnErrorValuesAsNull": True,
            },
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "sourceQueryCulture": "en-US",
            "tables": [
                {
                    "name": "Clients",
                    "lineageTag": str(uuid.uuid4()),
                    "columns": columns,
                    "measures": measures,
                    "partitions": [
                        {
                            "name": "Clients",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": m_expression(csv_path),
                            },
                        }
                    ],
                    "annotations": [
                        {"name": "PBI_ResultType", "value": "Table"},
                    ],
                }
            ],
            "annotations": [
                {"name": "PBIDesktopVersion", "value": "2.138.0.0"},
                {"name": "PBI_QueryOrder", "value": json.dumps(["Clients"])},
            ],
        },
    }
    return json.dumps(schema, ensure_ascii=False)


def report_layout() -> str:
    """Blank report with one page — user adds visuals in Power BI Desktop."""
    layout = {
        "id": 0,
        "resourcePackages": [],
        "sections": [
            {
                "id": 0,
                "name": "ReportSection",
                "displayName": "Overview",
                "filters": "[]",
                "ordinal": 0,
                "visualContainers": [],
                "config": json.dumps({
                    "relationships": [],
                    "selectedFilters": [],
                    "filterConfig": {"type": 1},
                }),
            }
        ],
        "config": "{}",
        "layoutOptimization": 0,
        "publicCustomVisuals": [],
    }
    return json.dumps(layout, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 powerbi_export.py <clean_excel_path> [output_dir]",
              file=sys.stderr)
        sys.exit(1)

    excel_path = sys.argv[1]
    output_dir = Path(sys.argv[2]) if len(sys.argv) >= 3 else Path(excel_path).parent

    df = load(excel_path)

    # Export XLSX — this is the data source Power BI will read from
    xlsx_path = output_dir / "Nomura_PowerBI_Data.xlsx"
    export_xlsx(df, xlsx_path)

    # Power BI reads all JSON parts as UTF-16 LE; XML/rels stay as UTF-8.
    def u16(s: str) -> bytes:
        return s.encode("utf-16-le")

    # Build .pbit (ZIP archive)
    pbit_path = output_dir / "Nomura_Dashboard.pbit"
    with zipfile.ZipFile(pbit_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels",         RELS)
        zf.writestr("Version",             VERSION)
        zf.writestr("SecurityBindings",    b"")
        zf.writestr("Settings",            u16(SETTINGS))
        zf.writestr("DataModelSchema",     u16(data_model_schema(str(xlsx_path))))
        zf.writestr("Report/Layout",       u16(report_layout()))

    print(f"PBIT : {pbit_path}")
    print(f"XLSX : {xlsx_path}")
    print()
    print("Open Nomura_Dashboard.pbit in Power BI Desktop.")
    print("Click Refresh — it will load data from Nomura_PowerBI_Data.csv.")
    print()
    print("Pre-built measures ready to use:")
    print("  Total Revenue, Total Clients, Total Meetings, MiFID Expired,")
    print("  Flagged Clients, Flagged %, MiFID Expired %, Revenue per Client,")
    print("  Avg Rev Per Meeting, Avg Engagement Score")


if __name__ == "__main__":
    main()
