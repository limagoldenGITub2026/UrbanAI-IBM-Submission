"""
sources_logger.py — SourcesLogger: audit trail for all written values.

Responsibilities
----------------
- Every value written to the workbook is linked to a Source_ID.
- Deduplicate sources by name: same source name → same Source_ID.
- Auto-generate Source_IDs as SRC_001, SRC_002, …
- Append new source rows to the SOURCES_LOG sheet (tbl_sources).
- Link a source to a specific field and record for full traceability.
"""

from __future__ import annotations

import logging
from datetime import date
from typing import Dict, List, Optional, Tuple

import openpyxl

from .reader import ExcelReader

logger = logging.getLogger(__name__)

SOURCES_SHEET = "SOURCES_LOG"
SOURCE_ID_PREFIX = "SRC_"

# SOURCES_LOG columns (must match the template exactly)
SOURCES_COLS = [
    "Source_ID",
    "Source_Name",
    "Source_Type",
    "Document_Date",
    "Reliability",
    "Conflict_Notes",
]


class SourcesLogger:
    """Manage the SOURCES_LOG sheet and provide audit-trail linking.

    Parameters
    ----------
    workbook:
        Open openpyxl workbook (same instance used by ExcelWriter).
    reader:
        ExcelReader for the same workbook — used to pre-load existing sources.
    """

    def __init__(
        self,
        workbook: openpyxl.Workbook,
        reader: ExcelReader,
    ) -> None:
        self.workbook = workbook
        self.reader = reader
        # Cache: lower-cased source_name → Source_ID
        self._source_cache: Dict[str, str] = {}
        # In-memory audit links: (source_id, field_id, record_id)
        self._links: List[Tuple[str, str, str]] = []

    # ------------------------------------------------------------------
    # Initialisation
    # ------------------------------------------------------------------

    def load_existing_sources(self) -> None:
        """Pre-load the SOURCES_LOG sheet into the in-memory cache.

        Call once after construction so that ``get_or_create_source``
        returns consistent IDs for sources logged in a previous session.
        """
        try:
            rows = self.reader.read_sheet(SOURCES_SHEET)
        except KeyError:
            logger.warning("Sheet '%s' not found — starting with empty source cache.", SOURCES_SHEET)
            return

        for row in rows:
            sid  = str(row.get("Source_ID", "") or "").strip()
            name = str(row.get("Source_Name", "") or "").strip()
            if sid and name:
                self._source_cache[name.lower()] = sid

    # ------------------------------------------------------------------
    # ID generation
    # ------------------------------------------------------------------

    def get_next_source_id(self) -> str:
        """Return the next available Source_ID (e.g. "SRC_004").

        Reads the maximum SRC_NNN from the cache and increments it.

        Returns
        -------
        str — e.g. "SRC_001" on an empty sheet.
        """
        if not self._source_cache:
            return f"{SOURCE_ID_PREFIX}001"

        max_n = 0
        for sid in self._source_cache.values():
            if sid.startswith(SOURCE_ID_PREFIX):
                try:
                    n = int(sid[len(SOURCE_ID_PREFIX):])
                    if n > max_n:
                        max_n = n
                except ValueError:
                    pass
        return f"{SOURCE_ID_PREFIX}{max_n + 1:03d}"

    # ------------------------------------------------------------------
    # Source management
    # ------------------------------------------------------------------

    def log_source(
        self,
        source_name: str,
        source_type: str = "",
        document_date: Optional[date] = None,
        reliability: str = "",
    ) -> str:
        """Append a new source row to SOURCES_LOG and return its Source_ID.

        Does NOT deduplicate — use ``get_or_create_source`` for deduplication.

        Parameters
        ----------
        source_name:
            Display name of the data source (e.g. "XPlan").
        source_type:
            Category string (e.g. "API", "Document", "GIS", "Manual").
        document_date:
            Date the source document was issued or retrieved.
        reliability:
            Free-text note (e.g. "Official", "Verified", "Pending").

        Returns
        -------
        str — The newly created Source_ID.
        """
        source_id = self.get_next_source_id()
        date_str = document_date.isoformat() if document_date else ""

        ws = self.workbook[SOURCES_SHEET]
        ws.append([source_id, source_name, source_type, date_str, reliability, ""])

        # Update cache
        self._source_cache[source_name.lower()] = source_id
        return source_id

    def get_or_create_source(self, source_name: str) -> str:
        """Return an existing Source_ID for *source_name*, or create a new one.

        Deduplication is by exact source name (case-insensitive).

        Parameters
        ----------
        source_name:
            Name of the data source.

        Returns
        -------
        str — Source_ID (existing or newly created).
        """
        key = source_name.lower().strip()
        if key in self._source_cache:
            return self._source_cache[key]

        return self.log_source(source_name)

    # ------------------------------------------------------------------
    # Audit linking
    # ------------------------------------------------------------------

    def link_source_to_field(
        self,
        source_id: str,
        field_id: str,
        record_id: str = "",
    ) -> None:
        """Record that *source_id* provided the value for *field_id* / *record_id*.

        Stored in-memory; used when writing conflict notes to SOURCES_LOG.

        Parameters
        ----------
        source_id:  Source_ID from SOURCES_LOG.
        field_id:   Field_ID that was written.
        record_id:  Optional unique record identifier (e.g. "P_6660_001").
        """
        self._links.append((source_id, field_id, record_id))

    def write_conflict_note(
        self,
        source_id: str,
        note: str,
    ) -> bool:
        """Append *note* to the Conflict_Notes cell of the row for *source_id*.

        Parameters
        ----------
        source_id:  Source_ID whose row should be updated.
        note:       Conflict description to append.

        Returns
        -------
        True if the row was found and updated.
        """
        ws = self.workbook[SOURCES_SHEET]
        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in ws[1]
        ]
        if "Source_ID" not in headers or "Conflict_Notes" not in headers:
            return False

        sid_col  = headers.index("Source_ID")
        note_col = headers.index("Conflict_Notes")

        for row in ws.iter_rows(min_row=2):
            if str(row[sid_col].value or "").strip() == source_id:
                existing = str(row[note_col].value or "").strip()
                row[note_col].value = f"{existing}; {note}".strip("; ") if existing else note
                return True
        return False
