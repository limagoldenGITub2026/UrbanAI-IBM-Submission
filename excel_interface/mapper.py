"""
mapper.py — DataMapper: translates Field_ID-keyed input dicts into WriteOperations.

Responsibilities
----------------
- Accept a dict[Field_ID → value] from the AI agent.
- Look up each Field_ID in the FieldLocation registry loaded from DATA_MAPPING.
- Enforce write_permission: raise WritePermissionError for protected fields.
- Warn (but do not raise) for unknown Field_IDs.
- Return a list of WriteOperation objects ready for ExcelWriter.
- Group operations by sheet for batched processing.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from typing import Any, Dict, List

from .models import FieldLocation, WriteOperation

logger = logging.getLogger(__name__)


class WritePermissionError(PermissionError):
    """Raised when a write is attempted on a field with write_permission=False."""

    def __init__(self, field_id: str) -> None:
        super().__init__(
            f"Field '{field_id}' has write_permission=False and cannot be written. "
            "Only system-generated IDs (e.g. Parcel_ID) are protected."
        )
        self.field_id = field_id


class DataMapper:
    """Translate an input dict[Field_ID → value] into WriteOperations.

    Parameters
    ----------
    field_locations:
        The registry produced by ``ExcelReader.load_data_mapping()``.
        Must be populated before calling ``map_input()``.
    """

    def __init__(self, field_locations: Dict[str, FieldLocation]) -> None:
        self.field_locations = field_locations

    def map_input(
        self,
        data_dict: Dict[str, Any],
        source_id: str = "",
    ) -> List[WriteOperation]:
        """Translate an input dict into a list of WriteOperations.

        Rules
        -----
        - Unknown Field_IDs: emit a warning, skip silently.
        - ``write_permission=False``: raise ``WritePermissionError``.
        - All others: produce a ``WriteOperation``.

        Parameters
        ----------
        data_dict:
            Input from the AI agent, e.g. ``{"SITE_001": "Tel Aviv"}``.
        source_id:
            Source_ID to attach to all operations (from SourcesLogger).

        Returns
        -------
        list[WriteOperation]

        Raises
        ------
        WritePermissionError
            If any Field_ID has ``write_permission=False``.
        """
        operations: List[WriteOperation] = []

        for field_id, value in data_dict.items():
            loc = self.field_locations.get(field_id)

            if loc is None:
                logger.warning(
                    "DataMapper: unknown Field_ID '%s' — skipping (not in DATA_MAPPING).",
                    field_id,
                )
                continue

            if not loc.write_permission:
                raise WritePermissionError(field_id)

            operations.append(
                WriteOperation(
                    field_id    = field_id,
                    sheet_name  = loc.sheet_name,
                    table_name  = loc.table_name,
                    column_name = loc.column_name,
                    value       = value,
                    source_id   = source_id,
                )
            )

        return operations

    def group_by_sheet(
        self,
        operations: List[WriteOperation],
    ) -> Dict[str, List[WriteOperation]]:
        """Group a list of WriteOperations by their target sheet name.

        Parameters
        ----------
        operations:
            Output of ``map_input()``.

        Returns
        -------
        dict[sheet_name → list[WriteOperation]]
        """
        grouped: Dict[str, List[WriteOperation]] = defaultdict(list)
        for op in operations:
            grouped[op.sheet_name].append(op)
        return dict(grouped)
