"""
reader.py — ExcelReader: loads a workbook and exposes read-only access.

Responsibilities
----------------
- Open the workbook with openpyxl (data_only=True so cached formula values are read).
- Always read DATA_MAPPING first; every other operation depends on it.
- Provide access to rows by table/sheet name, never by cell address.
- Detect SITE_DATA vertical layout (Field_ID | Label | Value | Status) vs. horizontal table layouts.
- Identify required fields that are empty.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl

from .models import FieldLocation

# Sheet name that uses a vertical Field_ID | Label | Value | Status layout
VERTICAL_SHEET = "SITE_DATA"

# Name of the mapping sheet — must be read before anything else
DATA_MAPPING_SHEET = "DATA_MAPPING"

# Column indices (1-based) in the DATA_MAPPING table
_DM_COL = {
    "Field_ID":         1,
    "Field_Name":       2,
    "Sheet_Name":       3,
    "Table_Name":       4,
    "Column_Name":      5,
    "Data_Type":        6,
    "Unit":             7,
    "Required":         8,
    "Write_Permission": 9,
    "Preferred_Source": 10,
    "Validation_Rule":  11,
    "Notes":            12,
}


def _cell_str(cell) -> str:
    """Return cell value as stripped string, or '' if None."""
    v = cell.value
    return str(v).strip() if v is not None else ""


class ExcelReader:
    """Read-only access to a unified urban-renewal Excel workbook.

    Always call ``load_data_mapping()`` before using any other method.

    Parameters
    ----------
    path:
        Path to the ``.xlsx`` workbook to open.
    """

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(f"Workbook not found: {self.path}")
        self.workbook = openpyxl.load_workbook(str(self.path), data_only=False)
        # Populated by load_data_mapping()
        self.field_locations: Dict[str, FieldLocation] = {}

    # ------------------------------------------------------------------
    # DATA_MAPPING
    # ------------------------------------------------------------------

    def load_data_mapping(self) -> Dict[str, FieldLocation]:
        """Read the DATA_MAPPING sheet and return a dict of FieldLocation objects.

        Must be called before any write or validation operation.

        Returns
        -------
        dict[field_id, FieldLocation]

        Raises
        ------
        KeyError
            If the workbook does not contain a DATA_MAPPING sheet.
        """
        if DATA_MAPPING_SHEET not in self.workbook.sheetnames:
            raise KeyError(
                f"Sheet '{DATA_MAPPING_SHEET}' not found in workbook. "
                "This workbook may not be a valid urban-renewal template."
            )

        ws = self.workbook[DATA_MAPPING_SHEET]
        result: Dict[str, FieldLocation] = {}

        # Row 1 is the header; data starts at row 2
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Skip completely blank rows
            if not any(v for v in row if v is not None):
                continue

            def _get(idx: int) -> str:
                """Safe 0-based column getter."""
                try:
                    v = row[idx - 1]
                    return str(v).strip() if v is not None else ""
                except IndexError:
                    return ""

            field_id   = _get(_DM_COL["Field_ID"])
            if not field_id:
                continue

            loc = FieldLocation(
                field_id         = field_id,
                sheet_name       = _get(_DM_COL["Sheet_Name"]),
                table_name       = _get(_DM_COL["Table_Name"]),
                column_name      = _get(_DM_COL["Column_Name"]),
                data_type        = _get(_DM_COL["Data_Type"]).lower() or "text",
                unit             = _get(_DM_COL["Unit"]),
                required         = _get(_DM_COL["Required"]).lower() == "yes",
                write_permission = _get(_DM_COL["Write_Permission"]).lower() == "yes",
                source_hint      = _get(_DM_COL["Preferred_Source"]),
            )
            result[field_id] = loc

        self.field_locations = result
        return result

    # ------------------------------------------------------------------
    # Sheet / table access
    # ------------------------------------------------------------------

    def read_sheet(self, sheet_name: str) -> List[Dict[str, Any]]:
        """Return all data rows of *sheet_name* as a list of dicts.

        For SITE_DATA (vertical layout) each row dict is
        ``{"Field_ID": ..., "Label": ..., "Value": ..., "Status": ...}``.
        For all other sheets each dict is keyed by the column header in row 1.

        Empty rows (all values None) are skipped.

        Parameters
        ----------
        sheet_name:
            One of: SITE_DATA, PARCELS, PLANS, BUILDINGS, CONSTRAINTS,
            SOURCES_LOG, DATA_MAPPING.

        Returns
        -------
        list[dict[column_name, value]]

        Raises
        ------
        KeyError
            If the sheet does not exist.
        """
        if sheet_name not in self.workbook.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found in workbook.")

        ws = self.workbook[sheet_name]

        # ── Read column headers from row 1 ──────────────────────────
        headers: List[str] = []
        for cell in ws[1]:
            v = cell.value
            headers.append(str(v).strip() if v is not None else "")

        # Strip trailing empty headers
        while headers and not headers[-1]:
            headers.pop()

        if not headers:
            return []

        # ── Read data rows (starting from row 2) ────────────────────
        rows: List[Dict[str, Any]] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Skip completely blank rows
            if not any(v for v in row if v is not None):
                continue
            row_dict: Dict[str, Any] = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                val = row[i] if i < len(row) else None
                row_dict[header] = val
            rows.append(row_dict)

        return rows

    def get_record_by_id(
        self,
        sheet_name: str,
        id_column: str,
        id_value: str,
    ) -> Optional[Dict[str, Any]]:
        """Find and return the first row where *id_column* == *id_value*.

        Parameters
        ----------
        sheet_name:
            Sheet to search (e.g. "PARCELS").
        id_column:
            Column header used as the unique key (e.g. "Parcel_ID").
        id_value:
            Value to look up (e.g. "P_6660_001").

        Returns
        -------
        Row dict if found, ``None`` otherwise.
        """
        for row in self.read_sheet(sheet_name):
            cell_val = row.get(id_column)
            if cell_val is not None and str(cell_val).strip() == str(id_value).strip():
                return row
        return None

    def get_site_value(self, field_id: str) -> Optional[Any]:
        """Return the Value cell for *field_id* in SITE_DATA, or None if not found.

        Convenience method for the vertical layout sheet.
        """
        for row in self.read_sheet(VERTICAL_SHEET):
            if str(row.get("Field_ID", "")).strip() == field_id:
                return row.get("Value")
        return None

    # ------------------------------------------------------------------
    # Missing-field detection
    # ------------------------------------------------------------------

    def identify_missing_fields(self) -> List[str]:
        """Return Field_IDs of required+writable fields that currently have no value.

        Uses the field_locations dict populated by ``load_data_mapping()``.
        Only fields with ``required=True`` and ``write_permission=True`` are checked.

        Returns
        -------
        list[field_id]

        Raises
        ------
        RuntimeError
            If ``load_data_mapping()`` has not been called first.
        """
        if not self.field_locations:
            raise RuntimeError(
                "field_locations is empty. Call load_data_mapping() before identify_missing_fields()."
            )

        missing: List[str] = []

        for field_id, loc in self.field_locations.items():
            if not loc.required or not loc.write_permission:
                continue

            if loc.sheet_name == VERTICAL_SHEET:
                # Vertical layout: look up by Field_ID in SITE_DATA
                val = self.get_site_value(field_id)
                if val is None or str(val).strip() == "":
                    missing.append(field_id)
            else:
                # Horizontal layout: field is "missing" if NO row has a value
                # in that column (i.e., the sheet has no data at all).
                # For required IDs (e.g. PLAN_001), flag if the table is empty.
                try:
                    rows = self.read_sheet(loc.sheet_name)
                except KeyError:
                    missing.append(field_id)
                    continue
                if not rows:
                    missing.append(field_id)
                # If there are rows, each individual record manages its own completeness.
                # Record-level validation happens in DataValidator.

        return missing

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def sheet_names(self) -> List[str]:
        """Return all sheet names present in the workbook."""
        return self.workbook.sheetnames

    def column_headers(self, sheet_name: str) -> List[str]:
        """Return the column headers (row 1) for *sheet_name*."""
        if sheet_name not in self.workbook.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        ws = self.workbook[sheet_name]
        headers = []
        for cell in ws[1]:
            v = cell.value
            if v is not None:
                headers.append(str(v).strip())
        return headers
