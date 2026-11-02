"""
validator.py — DataValidator: type-checking, duplicate detection, conflict detection.

Responsibilities
----------------
- Validate incoming values against the declared type in DATA_MAPPING.
- Apply domain rules: area > 0, floors >= 0, units >= 0, year 1800–2100.
- Detect duplicate records by unique ID before a row is appended.
- Detect conflicts (new value differs from an existing non-empty value).
- Accumulate all errors and conflicts; never raise on a validation failure —
  log it and allow the caller to decide whether to proceed.
- Produce structured error and missing-fields reports.
"""

from __future__ import annotations

import logging
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

from .models import ConflictRecord, FieldLocation, ValidationError
from .reader import ExcelReader

logger = logging.getLogger(__name__)

# Year range for Year_Built validation
YEAR_MIN = 1800
YEAR_MAX = 2100

# Domain rules: field_id → (min_value, max_value, allow_zero)
# None means no limit on that side
_DOMAIN_RULES: Dict[str, Tuple[Optional[float], Optional[float], bool]] = {
    "PARCEL_004": (0, None, False),    # area > 0
    "BUILD_001":  (0, None, True),     # floors >= 0
    "BUILD_002":  (0, None, True),     # units >= 0
    "BUILD_003":  (YEAR_MIN, YEAR_MAX, False),  # year in range
}


class DataValidator:
    """Validate values and detect conflicts against the workbook state.

    Parameters
    ----------
    field_locations:
        The registry produced by ``ExcelReader.load_data_mapping()``.
    reader:
        Used for duplicate checks and reading existing values.
    """

    def __init__(
        self,
        field_locations: Dict[str, FieldLocation],
        reader: ExcelReader,
    ) -> None:
        self.field_locations = field_locations
        self.reader = reader
        self._errors: List[ValidationError] = []
        self._conflicts: List[ConflictRecord] = []
        self._unauthorized: List[str] = []  # field_ids with write_permission=False

    # ------------------------------------------------------------------
    # Type validation
    # ------------------------------------------------------------------

    def validate_type(
        self,
        field_id: str,
        value: Any,
    ) -> Tuple[bool, str]:
        """Validate *value* against the declared type of *field_id*.

        Validation rules
        ----------------
        - ``text``    : any non-None value accepted.
        - ``integer`` : must be castable to int.
        - ``number``  : must be castable to float.
        - ``date``    : must be a ``datetime.date`` or ISO-8601 string.
        - ``boolean`` : must be bool or "true"/"false" (case-insensitive).

        Domain rules
        ------------
        PARCEL_004 (area_sqm): value > 0
        BUILD_001 (floors):    value >= 0
        BUILD_002 (units):     value >= 0
        BUILD_003 (year):      1800 <= value <= 2100

        Returns
        -------
        (is_valid, error_message)  — error_message is "" when valid.
        """
        loc = self.field_locations.get(field_id)
        if loc is None:
            return False, f"Unknown Field_ID '{field_id}'"

        dtype = loc.data_type

        # ── Type check ───────────────────────────────────────────────
        ok, msg = self._check_type(dtype, value)
        if not ok:
            self._record_error(field_id, value, msg)
            return False, msg

        # ── Domain rules ─────────────────────────────────────────────
        rule = _DOMAIN_RULES.get(field_id)
        if rule is not None:
            ok, msg = self._check_domain(field_id, value, rule)
            if not ok:
                self._record_error(field_id, value, msg)
                return False, msg

        return True, ""

    def _check_type(self, dtype: str, value: Any) -> Tuple[bool, str]:
        """Return (ok, msg) for basic type check."""
        if dtype == "text":
            return True, ""

        if dtype == "integer":
            try:
                int(value)
                return True, ""
            except (TypeError, ValueError):
                return False, f"Expected integer, got {type(value).__name__}: {value!r}"

        if dtype == "number":
            try:
                float(value)
                return True, ""
            except (TypeError, ValueError):
                return False, f"Expected number, got {type(value).__name__}: {value!r}"

        if dtype == "date":
            if isinstance(value, (date, datetime)):
                return True, ""
            if isinstance(value, str):
                try:
                    datetime.fromisoformat(value)
                    return True, ""
                except ValueError:
                    pass
            return False, f"Expected date or ISO-8601 string, got {type(value).__name__}: {value!r}"

        if dtype == "boolean":
            if isinstance(value, bool):
                return True, ""
            if isinstance(value, str) and value.lower() in ("true", "false"):
                return True, ""
            return False, f"Expected boolean, got {type(value).__name__}: {value!r}"

        # Unknown type — pass through
        return True, ""

    def _check_domain(
        self,
        field_id: str,
        value: Any,
        rule: Tuple[Optional[float], Optional[float], bool],
    ) -> Tuple[bool, str]:
        """Apply domain min/max rules to a numeric value."""
        min_val, max_val, allow_zero = rule
        try:
            num = float(value)
        except (TypeError, ValueError):
            return False, f"Cannot apply domain rule to non-numeric value: {value!r}"

        if not allow_zero and num == 0:
            return False, f"{field_id} must be > 0, got {num}"
        if min_val is not None and num < min_val:
            return False, f"{field_id} must be >= {min_val}, got {num}"
        if max_val is not None and num > max_val:
            return False, f"{field_id} must be <= {max_val}, got {num}"

        return True, ""

    # ------------------------------------------------------------------
    # Duplicate detection
    # ------------------------------------------------------------------

    def check_duplicate(
        self,
        sheet_name: str,
        id_column: str,
        id_value: str,
    ) -> Optional[Dict[str, Any]]:
        """Return the existing row if a record with *id_value* already exists.

        Parameters
        ----------
        sheet_name, id_column, id_value:
            As for ``ExcelReader.get_record_by_id``.

        Returns
        -------
        Existing row dict if found, ``None`` if no duplicate.
        """
        return self.reader.get_record_by_id(sheet_name, id_column, id_value)

    # ------------------------------------------------------------------
    # Conflict detection
    # ------------------------------------------------------------------

    def detect_conflict(
        self,
        field_id: str,
        new_value: Any,
        existing_value: Any,
        source_id: str = "",
    ) -> bool:
        """Return True and record a ConflictRecord if values differ.

        A conflict is flagged only when *both* new_value and existing_value
        are non-empty/non-None and they differ.

        Returns
        -------
        True if a conflict was detected.
        """
        if new_value is None or str(new_value).strip() == "":
            return False
        if existing_value is None or str(existing_value).strip() == "":
            return False
        if str(new_value).strip() == str(existing_value).strip():
            return False

        self._conflicts.append(
            ConflictRecord(
                field_id       = field_id,
                new_value      = new_value,
                existing_value = existing_value,
                source_id      = source_id,
            )
        )
        logger.warning(
            "Conflict detected for '%s': new=%r, existing=%r",
            field_id, new_value, existing_value,
        )
        return True

    # ------------------------------------------------------------------
    # Reports
    # ------------------------------------------------------------------

    def generate_error_report(self) -> Dict:
        """Return a structured report of all accumulated validation failures.

        Returns
        -------
        {
            "errors":              [{"field_id": ..., "value": ..., "reason": ...}],
            "conflicts":           [{"field_id": ..., "new": ..., "existing": ..., "source_id": ...}],
            "unauthorized_writes": [field_id, ...]
        }
        """
        return {
            "errors": [
                {"field_id": e.field_id, "value": e.value, "reason": e.reason}
                for e in self._errors
            ],
            "conflicts": [
                {
                    "field_id":  c.field_id,
                    "new":       c.new_value,
                    "existing":  c.existing_value,
                    "source_id": c.source_id,
                }
                for c in self._conflicts
            ],
            "unauthorized_writes": list(self._unauthorized),
        }

    def generate_missing_fields_report(self) -> Dict:
        """Return a report of required and optional fields that are empty.

        Returns
        -------
        {
            "required_missing": [field_id, ...],
            "optional_missing": [field_id, ...]
        }
        """
        missing_required: List[str] = []
        missing_optional: List[str] = []

        for field_id, loc in self.field_locations.items():
            if not loc.write_permission:
                continue

            value = None
            if loc.sheet_name == "SITE_DATA":
                value = self.reader.get_site_value(field_id)
            else:
                try:
                    rows = self.reader.read_sheet(loc.sheet_name)
                except KeyError:
                    rows = []
                # A field is "missing" at sheet level if there are no rows at all
                if not rows:
                    if loc.required:
                        missing_required.append(field_id)
                    else:
                        missing_optional.append(field_id)
                    continue
                # If rows exist, check each row's column value
                has_any_value = any(
                    row.get(loc.column_name) not in (None, "")
                    for row in rows
                )
                if not has_any_value:
                    if loc.required:
                        missing_required.append(field_id)
                    else:
                        missing_optional.append(field_id)
                continue

            if value is None or str(value).strip() == "":
                if loc.required:
                    missing_required.append(field_id)
                else:
                    missing_optional.append(field_id)

        return {
            "required_missing": missing_required,
            "optional_missing": missing_optional,
        }

    def record_unauthorized(self, field_id: str) -> None:
        """Record a write attempt on a protected field (write_permission=False)."""
        self._unauthorized.append(field_id)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _record_error(self, field_id: str, value: Any, reason: str) -> None:
        """Accumulate a ValidationError without raising."""
        self._errors.append(ValidationError(field_id=field_id, value=value, reason=reason))
