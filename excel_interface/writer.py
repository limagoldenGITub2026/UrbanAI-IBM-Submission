"""
writer.py — ExcelWriter: applies WriteOperations to the workbook.

Responsibilities
----------------
- Write individual fields in the SITE_DATA vertical layout (Field_ID | Label | Value | Status).
- Append new rows to horizontal-table sheets (PARCELS, PLANS, BUILDINGS, CONSTRAINTS).
- Update existing rows by unique ID — never overwrite without an ID check.
- Never modify existing record IDs (Parcel_ID, Plan_ID, etc.).
- Save output to a new ``_updated.xlsx`` file; keep a ``_backup_YYYYMMDD.xlsx`` copy.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl

from .reader import ExcelReader

# Sheet that uses vertical Field_ID | Label | Value | Status layout
VERTICAL_SHEET = "SITE_DATA"

# Status value written to SITE_DATA when a value is set
STATUS_COLLECTED = "Collected"


class ExcelWriter:
    """Write data to a workbook in-memory; save to disk on request.

    Parameters
    ----------
    workbook:
        An open ``openpyxl.Workbook`` object (already loaded by ExcelReader).
    reader:
        The ``ExcelReader`` instance associated with the same workbook,
        used for column-header lookups and ID checks before appends.
    original_path:
        Path of the source file — used to produce backup and output paths.
    """

    def __init__(
        self,
        workbook: openpyxl.Workbook,
        reader: ExcelReader,
        original_path: str | Path,
    ) -> None:
        self.workbook = workbook
        self.reader = reader
        self.original_path = Path(original_path)
        self._backup_created: bool = False

    # ------------------------------------------------------------------
    # SITE_DATA (vertical layout)
    # ------------------------------------------------------------------

    def write_field(self, field_id: str, value: Any) -> None:
        """Update a single field in the SITE_DATA vertical table.

        Finds the row whose Field_ID column equals *field_id*, sets its
        Value column, and updates Status to "Collected".
        Only works for SITE_DATA; raises ValueError otherwise.

        Parameters
        ----------
        field_id:
            The Field_ID to update (e.g. "SITE_001").
        value:
            The new value to write.

        Raises
        ------
        ValueError
            If *field_id* is not found in the SITE_DATA sheet.
        """
        ws = self.workbook[VERTICAL_SHEET]

        # Discover column positions from header row (row 1)
        headers = self._get_headers(VERTICAL_SHEET)
        try:
            fid_col   = headers.index("Field_ID") + 1   # 1-based
            value_col = headers.index("Value") + 1
            status_col = headers.index("Status") + 1
        except ValueError as exc:
            raise ValueError(
                f"SITE_DATA is missing required column: {exc}. "
                "Expected columns: Field_ID, Label, Value, Status."
            ) from exc

        for row in ws.iter_rows(min_row=2):
            cell_fid = row[fid_col - 1]
            if cell_fid.value is not None and str(cell_fid.value).strip() == field_id:
                row[value_col - 1].value = value
                row[status_col - 1].value = STATUS_COLLECTED
                return

        raise ValueError(
            f"Field_ID '{field_id}' not found in {VERTICAL_SHEET}. "
            "Add it to the sheet or update the template."
        )

    # ------------------------------------------------------------------
    # Horizontal-table sheets
    # ------------------------------------------------------------------

    def append_row(
        self,
        sheet_name: str,
        row_data: Dict[str, Any],
    ) -> None:
        """Append *row_data* as a new row in *sheet_name*'s table.

        Column order follows the table headers already in the sheet.
        Columns not present in *row_data* are left as None.

        Parameters
        ----------
        sheet_name:
            Target sheet (e.g. "PARCELS").
        row_data:
            Dict keyed by column header (e.g. ``{"Block_Number": "6660"}``).

        Notes
        -----
        Duplicate-checking is the caller's responsibility.
        ``ExcelDataInterface.write()`` always checks before calling this.
        """
        ws = self.workbook[sheet_name]
        headers = self._get_headers(sheet_name)

        new_row: List[Any] = [row_data.get(h) for h in headers]
        ws.append(new_row)

        # Extend every Table's ref to include the new last row,
        # otherwise Excel reports the table as corrupt on open.
        last_row = ws.max_row
        for tbl in ws.tables.values():
            first_col = tbl.ref.split(":")[0]           # e.g. "A1"
            last_col_letter = tbl.ref.split(":")[1]     # e.g. "E2"
            # keep only the column letter(s), strip digits
            import re as _re
            last_col = _re.sub(r"\d+", "", last_col_letter)
            tbl.ref = f"{first_col}:{last_col}{last_row}"

    def update_row(
        self,
        sheet_name: str,
        id_column: str,
        id_value: str,
        updates: Dict[str, Any],
    ) -> bool:
        """Update specific columns of an existing row.

        Finds the row where *id_column* == *id_value* and overwrites only the
        columns listed in *updates*.  The ID column itself is never modified.

        Parameters
        ----------
        sheet_name:
            Sheet to search (e.g. "PARCELS").
        id_column:
            Column header used as the unique key (e.g. "Parcel_ID").
        id_value:
            Identifier of the row to update (e.g. "P_6660_001").
        updates:
            Dict of column_name → new value.
            The *id_column* key is silently ignored if present.

        Returns
        -------
        True if the row was found and updated, False if not found.
        """
        ws = self.workbook[sheet_name]
        headers = self._get_headers(sheet_name)

        if id_column not in headers:
            return False

        id_col_idx = headers.index(id_column)   # 0-based

        for row in ws.iter_rows(min_row=2):
            cell_id = row[id_col_idx]
            if cell_id.value is not None and str(cell_id.value).strip() == str(id_value).strip():
                for col_name, new_val in updates.items():
                    if col_name == id_column:
                        continue   # never modify the ID column
                    if col_name in headers:
                        col_idx = headers.index(col_name)  # 0-based
                        row[col_idx].value = new_val
                return True

        return False

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def save(self, output_path: Optional[str | Path] = None) -> Path:
        """Save the workbook and create a date-stamped backup of the original.

        - The original file is copied to ``<stem>_backup_YYYYMMDD.xlsx``
          next to the original (only on the first save call per session).
        - The updated workbook is saved to *output_path*.  If omitted,
          defaults to ``<stem>_updated.xlsx`` next to the original.

        Parameters
        ----------
        output_path:
            Explicit destination path.  If ``None``, an ``_updated`` suffix
            is appended to the original stem.

        Returns
        -------
        Path
            The path where the updated workbook was saved.
        """
        import shutil

        # ── Backup original (once per session) ──────────────────────
        if not self._backup_created and self.original_path.exists():
            stamp = date.today().strftime("%Y%m%d")
            backup_path = self.original_path.with_name(
                f"{self.original_path.stem}_backup_{stamp}{self.original_path.suffix}"
            )
            if not backup_path.exists():
                shutil.copy2(str(self.original_path), str(backup_path))
            self._backup_created = True

        # ── Determine output path ────────────────────────────────────
        if output_path is None:
            dest = self.original_path.with_name(
                f"{self.original_path.stem}_updated{self.original_path.suffix}"
            )
        else:
            dest = Path(output_path)

        dest.parent.mkdir(parents=True, exist_ok=True)
        self.workbook.save(str(dest))
        return dest

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _get_headers(self, sheet_name: str) -> List[str]:
        """Return the column headers (row 1) for *sheet_name* as a list.

        The list is 0-indexed and may contain empty strings for blank cells.
        """
        ws = self.workbook[sheet_name]
        return [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in ws[1]
        ]
