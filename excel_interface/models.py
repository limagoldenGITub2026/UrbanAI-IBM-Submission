"""
models.py — Shared dataclasses and enums for excel_interface
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional


class DataStatus(str, Enum):
    """Status of a data field value.

    Used in every record and in SOURCES_LOG to indicate the provenance
    and verification state of each collected datum.
    """

    MISSING = "Missing Information"
    COLLECTED = "Collected"
    REQUIRES_VERIFICATION = "Requires Verification"
    VERIFIED = "Verified"
    CONFLICTING = "Conflicting Information"
    NOT_APPLICABLE = "Not Applicable"


@dataclass
class FieldLocation:
    """Describes where a Field_ID lives in the workbook and its write rules.

    Loaded from the DATA_MAPPING sheet at the start of every session.
    Never derived from cell position or colour.

    Attributes:
        field_id:         Unique identifier (e.g. "SITE_001", "PARCEL_002").
        sheet_name:       Excel sheet name (e.g. "SITE_DATA", "PARCELS").
        table_name:       Named Excel table inside the sheet (e.g. "tbl_site_data").
        column_name:      Column header in the table (e.g. "Value", "Block_Number").
        data_type:        Declared type: "text" | "integer" | "number" | "date" | "boolean".
        unit:             Optional measurement unit (e.g. "sqm", "floors"). May be empty.
        required:         Whether the field must have a value (True = mandatory).
        write_permission: False means the field is system-managed; writes are blocked.
        source_hint:      Suggested data source from DATA_MAPPING (informational only).
    """

    field_id: str
    sheet_name: str
    table_name: str
    column_name: str
    data_type: str                  # "text" | "integer" | "number" | "date" | "boolean"
    unit: str = ""
    required: bool = False
    write_permission: bool = True
    source_hint: str = ""


@dataclass
class WriteOperation:
    """A single resolved write operation, ready for ExcelWriter to apply.

    Produced by DataMapper after translating an input dict[Field_ID → value]
    through DATA_MAPPING.

    Attributes:
        field_id:    The Field_ID that originated this operation.
        sheet_name:  Target sheet.
        table_name:  Target named table.
        column_name: Target column header.
        value:       Value to write (already type-checked by DataValidator).
        source_id:   Source_ID from SOURCES_LOG for audit trail.
        record_id:   Optional existing record identifier (for updates).
    """

    field_id: str
    sheet_name: str
    table_name: str
    column_name: str
    value: Any
    source_id: str
    record_id: Optional[str] = None


@dataclass
class ValidationError:
    """A single validation failure recorded by DataValidator.

    Attributes:
        field_id:  The Field_ID that failed.
        value:     The value that was rejected.
        reason:    Human-readable error description.
    """

    field_id: str
    value: Any
    reason: str


@dataclass
class ConflictRecord:
    """A data conflict detected between an incoming value and an existing one.

    Attributes:
        field_id:       The Field_ID in question.
        new_value:      The incoming value.
        existing_value: The value already stored in the workbook.
        source_id:      Source_ID of the incoming value.
    """

    field_id: str
    new_value: Any
    existing_value: Any
    source_id: str
