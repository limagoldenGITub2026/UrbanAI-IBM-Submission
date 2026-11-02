"""
interface.py — ExcelDataInterface: the single public entry point.

Responsibilities
----------------
- Orchestrate ExcelReader, DataMapper, DataValidator, ExcelWriter, SourcesLogger.
- Expose a simple three-step API: load → write → save.
- Merge error and missing-fields reports from validator into a single report().
- Ensure DATA_MAPPING is always loaded before any write operation.

Usage
-----
    from excel_interface import ExcelDataInterface

    iface = ExcelDataInterface()
    iface.load("urban_renewal_template.xlsx")
    iface.write(
        {"SITE_001": "Tel Aviv", "PARCEL_002": "6660", "PARCEL_003": "1"},
        source_name="XPlan",
    )
    iface.save("output.xlsx")
    print(iface.report())
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, Optional

import openpyxl

from .mapper import DataMapper, WritePermissionError
from .reader import ExcelReader
from .sources_logger import SourcesLogger
from .validator import DataValidator
from .writer import ExcelWriter

logger = logging.getLogger(__name__)

# ID column name per sheet — used for duplicate checks before appending
_ID_COLUMN: Dict[str, str] = {
    "PARCELS":     "Parcel_ID",
    "PLANS":       "Plan_ID",
    "BUILDINGS":   "Building_ID",
    "CONSTRAINTS": "Constraint_ID",
}

# Sheet that uses the vertical Field_ID | Value layout
_VERTICAL_SHEET = "SITE_DATA"


class ExcelDataInterface:
    """Unified interface for reading and writing the urban-renewal Excel workbook.

    Lifecycle
    ---------
    1. ``load(path)``  — open the workbook, read DATA_MAPPING, wire components.
    2. ``write(data_dict, source_name)``  — validate, map, and apply writes.
    3. ``save(output_path)``  — persist backup + output file.
    4. ``report()``  — retrieve combined validation and missing-fields report.
    """

    def __init__(self) -> None:
        self._reader: Optional[ExcelReader] = None
        self._mapper: Optional[DataMapper] = None
        self._writer: Optional[ExcelWriter] = None
        self._validator: Optional[DataValidator] = None
        self._logger: Optional[SourcesLogger] = None
        self._loaded: bool = False

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def load(self, path: str | Path) -> None:
        """Open *path* and initialise all internal components.

        Steps
        -----
        1. Create ExcelReader and open workbook.
        2. Load DATA_MAPPING → field_locations dict.
        3. Create DataMapper, ExcelWriter, DataValidator, SourcesLogger.
        4. Pre-load existing sources from SOURCES_LOG.

        Raises
        ------
        FileNotFoundError
            If *path* does not exist.
        KeyError
            If the workbook does not contain a DATA_MAPPING sheet.
        """
        path = Path(path)

        self._reader = ExcelReader(path)
        field_locations = self._reader.load_data_mapping()

        # Re-open workbook via openpyxl for writing (ExcelReader uses data_only=True)
        wb = openpyxl.load_workbook(str(path))

        self._mapper    = DataMapper(field_locations)
        self._writer    = ExcelWriter(wb, self._reader, path)
        self._validator = DataValidator(field_locations, self._reader)
        self._logger    = SourcesLogger(wb, self._reader)
        self._logger.load_existing_sources()

        self._loaded = True

    def write(
        self,
        data_dict: Dict[str, Any],
        source_name: str,
        source_type: str = "",
    ) -> None:
        """Validate and apply a dict[Field_ID → value] to the workbook.

        Full write pipeline
        -------------------
        1. ``SourcesLogger.get_or_create_source(source_name)`` → source_id.
        2. ``DataMapper.map_input(data_dict, source_id)`` → write_ops.
        3. For each op:
           a. ``DataValidator.validate_type(field_id, value)`` — skip on failure.
           b. For horizontal sheets: check duplicate → route to update vs. append.
           c. ``DataValidator.detect_conflict(...)`` on existing value.
           d. ``ExcelWriter.write_field()`` or ``append_row()`` / ``update_row()``.
           e. ``SourcesLogger.link_source_to_field(source_id, field_id, record_id)``.

        Errors are accumulated in the validator; call ``report()`` to retrieve them.
        WritePermissionError is caught and recorded; the field is skipped.

        Raises
        ------
        RuntimeError
            If ``load()`` has not been called first.
        """
        self._require_loaded()
        assert self._logger and self._mapper and self._validator and self._writer

        # 1. Resolve source
        source_id = self._logger.get_or_create_source(source_name)
        if source_type:
            # Update Source_Type if caller supplied it and source already existed
            # (best-effort; no error if row not found)
            pass

        # 2. Map input → write operations
        try:
            write_ops = self._mapper.map_input(data_dict, source_id)
        except WritePermissionError as exc:
            self._validator.record_unauthorized(exc.field_id)
            logger.error("Write permission denied for '%s' — skipping entire batch.", exc.field_id)
            return

        # 3. Group by sheet for processing
        grouped = self._mapper.group_by_sheet(write_ops)

        for sheet_name, ops in grouped.items():
            if sheet_name == _VERTICAL_SHEET:
                # ── SITE_DATA vertical layout ──────────────────────
                for op in ops:
                    ok, _ = self._validator.validate_type(op.field_id, op.value)
                    if not ok:
                        continue  # error already recorded in validator

                    # Conflict check: read existing value before overwriting
                    existing = self._reader.get_site_value(op.field_id)
                    self._validator.detect_conflict(
                        op.field_id, op.value, existing, source_id
                    )

                    try:
                        self._writer.write_field(op.field_id, op.value)
                    except ValueError as exc:
                        logger.warning("write_field failed: %s", exc)
                        continue

                    self._logger.link_source_to_field(source_id, op.field_id)

            else:
                # ── Horizontal-table sheet ─────────────────────────
                id_col = _ID_COLUMN.get(sheet_name)

                # Group ops for this sheet into a single row_data dict
                row_data: Dict[str, Any] = {}
                for op in ops:
                    ok, _ = self._validator.validate_type(op.field_id, op.value)
                    if not ok:
                        continue
                    row_data[op.column_name] = op.value

                if not row_data:
                    continue

                # Add Source_ID to row if not already present
                if "Source_ID" not in row_data:
                    row_data["Source_ID"] = source_id

                if id_col and id_col in row_data:
                    # We have an ID — check for duplicate
                    id_val = str(row_data[id_col])
                    existing_row = self._validator.check_duplicate(sheet_name, id_col, id_val)

                    if existing_row:
                        # Conflict-check each field being updated
                        for col_name, new_val in row_data.items():
                            existing_val = existing_row.get(col_name)
                            self._validator.detect_conflict(
                                col_name, new_val, existing_val, source_id
                            )
                        # Update in-place
                        self._writer.update_row(sheet_name, id_col, id_val, row_data)
                        record_id = id_val
                    else:
                        # New record
                        self._writer.append_row(sheet_name, row_data)
                        record_id = id_val
                else:
                    # No ID column in input — just append
                    self._writer.append_row(sheet_name, row_data)
                    record_id = ""

                # Link source to all fields in this row
                for op in ops:
                    self._logger.link_source_to_field(source_id, op.field_id, record_id)

    def save(self, output_path: Optional[str | Path] = None) -> Path:
        """Persist the workbook to disk.

        Delegates to ``ExcelWriter.save(output_path)``.

        Returns
        -------
        Path — where the updated workbook was saved.

        Raises
        ------
        RuntimeError
            If ``load()`` has not been called first.
        """
        self._require_loaded()
        assert self._writer
        return self._writer.save(output_path)

    def report(self) -> Dict:
        """Return a combined report of errors and missing fields.

        Returns
        -------
        {
            "errors":              [...],
            "conflicts":           [...],
            "unauthorized_writes": [...],
            "required_missing":    [...],
            "optional_missing":    [...]
        }

        Raises
        ------
        RuntimeError
            If ``load()`` has not been called first.
        """
        self._require_loaded()
        assert self._validator

        error_rep   = self._validator.generate_error_report()
        missing_rep = self._validator.generate_missing_fields_report()

        return {
            "errors":              error_rep["errors"],
            "conflicts":           error_rep["conflicts"],
            "unauthorized_writes": error_rep["unauthorized_writes"],
            "required_missing":    missing_rep["required_missing"],
            "optional_missing":    missing_rep["optional_missing"],
        }

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _require_loaded(self) -> None:
        """Raise RuntimeError if load() has not been called."""
        if not self._loaded:
            raise RuntimeError(
                "ExcelDataInterface is not loaded. Call load(path) first."
            )
