"""
create_template.py — One-time script to generate urban_renewal_template.xlsx

Run from the 2026.11.02/ directory:
    python -m excel_interface.create_template

Produces:
    2026.11.02/urban_renewal_template.xlsx  (unified 8-sheet workbook)

Sheet layout (matches DATA_MAPPING exactly)
-------------------------------------------
1. README          — operating rules (plain text, no named table)
2. SITE_DATA       — tbl_site_data   (vertical: Field_ID | Label | Value | Status)
3. PARCELS         — tbl_parcels     (Parcel_ID | Block_Number | Parcel_Number | Area_sqm | Source_ID)
4. PLANS           — tbl_plans       (Plan_ID | Plan_Number | Plan_Status | Land_Use | Source_ID)
5. BUILDINGS       — tbl_buildings   (Building_ID | Parcel_ID | Floors_Above_Ground | Dwelling_Units | Year_Built | Source_ID)
6. CONSTRAINTS     — tbl_constraints (Constraint_ID | Parcel_ID | Constraint_Type | Description | Status | Source_ID)
7. SOURCES_LOG     — tbl_sources     (Source_ID | Source_Name | Source_Type | Document_Date | Reliability | Conflict_Notes)
8. DATA_MAPPING    — tbl_data_mapping (full field registry)

Design rules
------------
- Named Excel tables (openpyxl.worksheet.table.Table) are used on all data sheets.
- Column names match DATA_MAPPING Column_Name values exactly.
- Sheet names match DATA_MAPPING Sheet_Name values exactly (PLANS, BUILDINGS, CONSTRAINTS).
- No formulas, no conditional formatting, no colours used for identification.
- The script is idempotent: running it again overwrites the template file.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Output path (relative to this file's parent directory — 2026.11.02/)
OUTPUT_PATH = Path(__file__).parent.parent / "urban_renewal_template.xlsx"

# ── Colour palette (used only for header row readability — NOT for identification) ──
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
CELL_FONT   = Font(name="Calibri", size=10)
ALIGN_WRAP  = Alignment(wrap_text=True, vertical="top")
ALIGN_WRAP_RTL = Alignment(wrap_text=True, vertical="top")

# ── Sheet / table definitions ──────────────────────────────────────────────────

README_RULES = [
    ("Rule", "Description"),
    ("Field identification", "Always use Field_ID → Sheet_Name → Table_Name → Column_Name. Never use cell position or colour."),
    ("Read DATA_MAPPING first", "Load DATA_MAPPING before any read or write operation."),
    ("Write permission", "Only write to fields where Write_Permission = yes."),
    ("Duplicate prevention", "Check for existing record by unique ID before appending a new row."),
    ("No ID changes", "Never modify Parcel_ID, Plan_ID, Building_ID, Constraint_ID, or Source_ID once created."),
    ("Source traceability", "Every value written must be linked to a Source_ID in SOURCES_LOG."),
    ("No invented data", "Leave fields empty and mark status as 'Missing Information' — do not invent values."),
    ("Conflict flagging", "When new value differs from existing non-empty value, set status to 'Conflicting Information'."),
    ("Separate collected vs verified", "Use DataStatus: Collected / Requires Verification / Verified."),
    ("Original backup", "Save a backup before modifying any existing file."),
    ("Extensibility", "New columns/fields require updating DATA_MAPPING only — do not change existing column names."),
    ("Error reporting", "Produce a missing-fields report and an error report after every write session."),
]

SITE_DATA_ROWS = [
    # (Field_ID, Label, Value, Status)
    ("SITE_001", "שם עיר",          "", "Missing Information"),
    ("SITE_002", "שכונה",           "", "Missing Information"),
    ("SITE_003", "כתובת / ציר",     "", "Missing Information"),
    ("SITE_004", "גוש",             "", "Missing Information"),
    ("SITE_005", "שטח מתחם כולל",   "", "Missing Information"),
]

PARCELS_COLS        = ["Parcel_ID", "Block_Number", "Parcel_Number", "Area_sqm", "Source_ID"]
PLANS_COLS          = ["Plan_ID", "Plan_Number", "Plan_Status", "Land_Use", "Source_ID"]
BUILDINGS_COLS      = ["Building_ID", "Parcel_ID", "Floors_Above_Ground", "Dwelling_Units", "Year_Built", "Source_ID"]
CONSTRAINTS_COLS    = ["Constraint_ID", "Parcel_ID", "Constraint_Type", "Description", "Status", "Source_ID"]
SOURCES_LOG_COLS    = ["Source_ID", "Source_Name", "Source_Type", "Document_Date", "Reliability", "Conflict_Notes"]

DATA_MAPPING_COLS = [
    "Field_ID", "Field_Name", "Sheet_Name", "Table_Name", "Column_Name",
    "Data_Type", "Unit", "Required", "Write_Permission", "Preferred_Source",
    "Validation_Rule", "Notes",
]

DATA_MAPPING_ROWS = [
    # (Field_ID, Field_Name, Sheet_Name, Table_Name, Column_Name, Data_Type, Unit, Required, Write_Permission, Preferred_Source, Validation_Rule, Notes)
    ("SITE_001",  "שם עיר",                  "SITE_DATA",   "tbl_site_data",    "Value",               "text",    "",       "yes", "yes", "GIS עירוני",           "לא ריק",              "מיפוי ערכי טופס"),
    ("SITE_002",  "שכונה",                   "SITE_DATA",   "tbl_site_data",    "Value",               "text",    "",       "no",  "yes", "GIS עירוני",           "טקסט",                ""),
    ("SITE_003",  "כתובת / ציר",             "SITE_DATA",   "tbl_site_data",    "Value",               "text",    "",       "no",  "yes", "GIS / כתובת",          "טקסט",                ""),
    ("SITE_005",  "שטח מתחם כולל",           "SITE_DATA",   "tbl_site_data",    "Value",               "number",  "sqm",    "no",  "yes", "GIS / מדידה",          "גדול מאפס",           ""),
    ("PARCEL_001","מזהה חלקה",               "PARCELS",     "tbl_parcels",      "Parcel_ID",           "text",    "",       "yes", "no",  "מערכת פנימית",         "ייחודי ולא ריק",      "נוצר על ידי המערכת"),
    ("PARCEL_002","מספר גוש",                "PARCELS",     "tbl_parcels",      "Block_Number",        "text",    "",       "yes", "yes", "מפ\"י / GIS",          "טקסט ספרות",          ""),
    ("PARCEL_003","מספר חלקה",               "PARCELS",     "tbl_parcels",      "Parcel_Number",       "text",    "",       "yes", "yes", "מפ\"י / GIS",          "טקסט ספרות",          ""),
    ("PARCEL_004","שטח חלקה",                "PARCELS",     "tbl_parcels",      "Area_sqm",            "number",  "sqm",    "no",  "yes", "מפ\"י / GIS",          "גדול מאפס",           ""),
    ("PLAN_001",  "מספר תכנית",              "PLANS",       "tbl_plans",        "Plan_Number",         "text",    "",       "yes", "yes", "תכנון זמין / XPlan",   "לא ריק",              ""),
    ("PLAN_002",  "סטטוס תכנית",             "PLANS",       "tbl_plans",        "Plan_Status",         "text",    "",       "no",  "yes", "תכנון זמין / מסמכי תכנית", "ערך מתוך רשימה",  ""),
    ("PLAN_003",  "ייעוד קרקע",              "PLANS",       "tbl_plans",        "Land_Use",            "text",    "",       "no",  "yes", "תקנון ותשריט",         "טקסט + מקור",         ""),
    ("BUILD_001", "מספר קומות מעל הקרקע",    "BUILDINGS",   "tbl_buildings",    "Floors_Above_Ground", "integer", "floors", "no",  "yes", "תיק בניין / מדידה",   "מספר שלם >= 0",       ""),
    ("BUILD_002", "מספר יחידות דיור",         "BUILDINGS",   "tbl_buildings",    "Dwelling_Units",      "integer", "units",  "no",  "yes", "היתר / תיק בניין",    "מספר שלם >= 0",       ""),
    ("BUILD_003", "שנת בנייה",               "BUILDINGS",   "tbl_buildings",    "Year_Built",          "integer", "year",   "no",  "yes", "תיק בניין / GIS",      "טווח 1800–2100",      ""),
    ("CONST_001", "סוג מגבלה",               "CONSTRAINTS", "tbl_constraints",  "Constraint_Type",     "text",    "",       "no",  "yes", "תכנית / GIS / רשות",   "לא ריק אם קיים",     ""),
    ("CONST_002", "תיאור מגבלה",             "CONSTRAINTS", "tbl_constraints",  "Description",         "text",    "",       "no",  "yes", "מסמך מקור",            "טקסט + Source_ID",    ""),
    ("SRC_001",   "שם מקור מידע",            "SOURCES_LOG", "tbl_sources",      "Source_Name",         "text",    "",       "yes", "yes", "מערכת המקור",          "לא ריק",              ""),
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def _add_named_table(ws, ref: str, table_name: str) -> None:
    """Add an Excel named table to worksheet *ws* covering *ref* (e.g. 'A1:E1')."""
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tbl = Table(displayName=table_name, ref=ref)
    tbl.tableStyleInfo = style
    ws.add_table(tbl)


def _style_header_row(ws, n_cols: int, row: int = 1) -> None:
    """Apply header styling to row *row* across *n_cols* columns."""
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = ALIGN_WRAP


def _set_col_widths(ws, widths: list[int]) -> None:
    """Set column widths by position (1-based)."""
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── Sheet builders ─────────────────────────────────────────────────────────────

def _build_readme(wb: openpyxl.Workbook) -> None:
    ws = wb.create_sheet("README")
    ws.sheet_view.rightToLeft = False

    ws["A1"] = "Urban Renewal AI — Excel Data Interface Operating Rules"
    ws["A1"].font = Font(bold=True, size=14, name="Calibri")
    ws["A2"] = "Read this sheet before using the workbook. Field identification must always go through Field_ID → DATA_MAPPING."
    ws["A2"].font = Font(italic=True, size=10, name="Calibri", color="595959")

    ws.append([])  # blank row

    # Header
    ws.append(["Rule", "Description"])
    _style_header_row(ws, 2, row=ws.max_row)

    for rule, desc in README_RULES:
        ws.append([rule, desc])
        for col in [1, 2]:
            ws.cell(row=ws.max_row, column=col).alignment = ALIGN_WRAP
            ws.cell(row=ws.max_row, column=col).font = CELL_FONT

    _set_col_widths(ws, [30, 90])
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A5"


def _build_site_data(wb: openpyxl.Workbook) -> None:
    ws = wb.create_sheet("SITE_DATA")
    ws.sheet_view.rightToLeft = True  # Hebrew RTL

    headers = ["Field_ID", "Label", "Value", "Status"]
    ws.append(headers)
    _style_header_row(ws, len(headers), row=1)

    for fid, label, value, status in SITE_DATA_ROWS:
        ws.append([fid, label, value, status])
        for col in range(1, len(headers) + 1):
            ws.cell(row=ws.max_row, column=col).font = CELL_FONT
            ws.cell(row=ws.max_row, column=col).alignment = ALIGN_WRAP_RTL

    # Named table covers header + data rows
    last_row = len(SITE_DATA_ROWS) + 1
    _add_named_table(ws, f"A1:D{last_row}", "tbl_site_data")
    _set_col_widths(ws, [14, 28, 40, 24])
    ws.freeze_panes = "C2"


def _build_horizontal_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    table_name: str,
    columns: list[str],
    col_widths: list[int],
) -> None:
    """Generic builder for horizontal-table sheets (PARCELS, PLANS, BUILDINGS, CONSTRAINTS, SOURCES_LOG)."""
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.rightToLeft = False

    ws.append(columns)
    _style_header_row(ws, len(columns), row=1)

    # Named table — header row only (no data rows yet, but Excel requires at least the header)
    _add_named_table(ws, f"A1:{get_column_letter(len(columns))}1", table_name)
    _set_col_widths(ws, col_widths)
    ws.freeze_panes = "B2"


def _build_data_mapping(wb: openpyxl.Workbook) -> None:
    ws = wb.create_sheet("DATA_MAPPING")
    ws.sheet_view.rightToLeft = False

    ws.append(DATA_MAPPING_COLS)
    _style_header_row(ws, len(DATA_MAPPING_COLS), row=1)

    for row_data in DATA_MAPPING_ROWS:
        ws.append(list(row_data))
        for col in range(1, len(DATA_MAPPING_COLS) + 1):
            ws.cell(row=ws.max_row, column=col).font = CELL_FONT
            ws.cell(row=ws.max_row, column=col).alignment = ALIGN_WRAP

    last_row = len(DATA_MAPPING_ROWS) + 1
    _add_named_table(ws, f"A1:{get_column_letter(len(DATA_MAPPING_COLS))}{last_row}", "tbl_data_mapping")
    _set_col_widths(ws, [14, 30, 16, 20, 24, 10, 8, 10, 18, 26, 24, 30])
    ws.freeze_panes = "B2"


# ── Main ───────────────────────────────────────────────────────────────────────

def create_template(output_path: Path = OUTPUT_PATH) -> Path:
    """Create the unified urban_renewal_template.xlsx workbook.

    Parameters
    ----------
    output_path:
        Destination path.  Defaults to ``OUTPUT_PATH``.

    Returns
    -------
    Path — the path where the workbook was saved.
    """
    wb = openpyxl.Workbook()

    # Remove default empty sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # 1. README
    _build_readme(wb)

    # 2. SITE_DATA (vertical layout)
    _build_site_data(wb)

    # 3. PARCELS
    _build_horizontal_sheet(
        wb, "PARCELS", "tbl_parcels", PARCELS_COLS,
        col_widths=[18, 16, 16, 14, 14],
    )

    # 4. PLANS
    _build_horizontal_sheet(
        wb, "PLANS", "tbl_plans", PLANS_COLS,
        col_widths=[18, 20, 20, 30, 14],
    )

    # 5. BUILDINGS
    _build_horizontal_sheet(
        wb, "BUILDINGS", "tbl_buildings", BUILDINGS_COLS,
        col_widths=[18, 18, 22, 18, 14, 14],
    )

    # 6. CONSTRAINTS
    _build_horizontal_sheet(
        wb, "CONSTRAINTS", "tbl_constraints", CONSTRAINTS_COLS,
        col_widths=[18, 18, 22, 40, 24, 14],
    )

    # 7. SOURCES_LOG
    _build_horizontal_sheet(
        wb, "SOURCES_LOG", "tbl_sources", SOURCES_LOG_COLS,
        col_widths=[14, 28, 16, 16, 16, 40],
    )

    # 8. DATA_MAPPING
    _build_data_mapping(wb)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    print(f"Template created: {output_path}")
    return output_path


if __name__ == "__main__":
    path = create_template()
    sys.exit(0)
