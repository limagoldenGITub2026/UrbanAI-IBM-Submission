"""
base.py — Shared data contracts for all connectors.

Every connector returns a CollectionResult with a standardised structure
so the CollectionEngine and Streamlit UI can handle all connectors uniformly.
"""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class CollectionStatus(str, Enum):
    """Outcome of a single connector's fetch attempt."""
    OK               = "ok"               # data collected successfully
    PARTIAL          = "partial"          # some fields collected, some missing
    UNAVAILABLE      = "unavailable"      # API unreachable / network error
    MANUAL_REQUIRED  = "manual_required"  # source requires auth / payment


@dataclass
class CollectionResult:
    """Standardised result returned by every connector.

    Attributes
    ----------
    source_name:
        Human-readable name of the data source (e.g. "Govmap WFS").
    field_data:
        Dict[Field_ID → value] ready to pass to ExcelDataInterface.write().
    status:
        Outcome of the fetch attempt.
    notes:
        Human-readable explanation — shown in the UI per connector row.
    fields_collected:
        List of Field_IDs successfully populated (subset of field_data keys).
    """
    source_name:      str
    field_data:       Dict[str, Any]
    status:           CollectionStatus
    notes:            str                        = ""
    fields_collected: List[str]                  = field(default_factory=list)

    @property
    def is_ok(self) -> bool:
        return self.status in (CollectionStatus.OK, CollectionStatus.PARTIAL)

    @property
    def status_emoji(self) -> str:
        return {
            CollectionStatus.OK:              "✅",
            CollectionStatus.PARTIAL:         "⚠️",
            CollectionStatus.UNAVAILABLE:     "❌",
            CollectionStatus.MANUAL_REQUIRED: "🔒",
        }[self.status]


class BaseConnector(ABC):
    """Abstract base class for all data-source connectors.

    Subclasses must:
    - Set the ``name`` class attribute.
    - Set ``required_fields`` with the Field_IDs this connector can populate.
    - Implement ``fetch()``.

    All HTTP calls should use a 10-second timeout.
    On network errors: catch the exception, return status=UNAVAILABLE.
    On partial data:   return status=PARTIAL with explanation in notes.
    """

    name: str = "BaseConnector"
    required_fields: List[str] = []   # Field_IDs this connector can supply

    @abstractmethod
    def fetch(
        self,
        city: str,
        address: str,
        block:  Optional[str] = None,
        parcel: Optional[str] = None,
    ) -> CollectionResult:
        """Fetch data from this source and return a CollectionResult.

        Parameters
        ----------
        city:
            City name in Hebrew (e.g. "תל אביב").
        address:
            Street address in Hebrew (e.g. "רוטשילד 1").
        block:
            Optional gush number (may help narrow geocoding).
        parcel:
            Optional parcel number.

        Returns
        -------
        CollectionResult
        """
        ...

    def _unavailable(self, reason: str) -> CollectionResult:
        """Convenience factory: return an UNAVAILABLE result."""
        logger.warning("%s: unavailable — %s", self.name, reason)
        return CollectionResult(
            source_name = self.name,
            field_data  = {},
            status      = CollectionStatus.UNAVAILABLE,
            notes       = reason,
        )

    def _partial(self, field_data: Dict[str, Any], notes: str) -> CollectionResult:
        """Convenience factory: return a PARTIAL result."""
        return CollectionResult(
            source_name      = self.name,
            field_data       = field_data,
            status           = CollectionStatus.PARTIAL,
            notes            = notes,
            fields_collected = list(field_data.keys()),
        )

    def _ok(self, field_data: Dict[str, Any], notes: str = "") -> CollectionResult:
        """Convenience factory: return an OK result."""
        return CollectionResult(
            source_name      = self.name,
            field_data       = field_data,
            status           = CollectionStatus.OK,
            notes            = notes,
            fields_collected = list(field_data.keys()),
        )

    def _manual_required(self, notes: str) -> CollectionResult:
        """Convenience factory: return a MANUAL_REQUIRED result."""
        return CollectionResult(
            source_name = self.name,
            field_data  = {},
            status      = CollectionStatus.MANUAL_REQUIRED,
            notes       = notes,
        )
