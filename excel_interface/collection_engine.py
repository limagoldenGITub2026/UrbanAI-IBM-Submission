"""
collection_engine.py — CollectionEngine

Orchestrates all data connectors, writes results to ExcelDataInterface,
and returns a structured summary for the UI to display.

Usage
-----
    engine = CollectionEngine()
    summary = engine.run(city="תל אביב", address="רוטשילד 1")
    print(summary.output_path)
    for r in summary.results:
        print(r.source_name, r.status, r.notes)
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, List, Optional

from excel_interface.connectors.base import (
    BaseConnector,
    CollectionResult,
    CollectionStatus,
)
from excel_interface.connectors.cbs import CBSConnector
from excel_interface.connectors.data_gov import DataGovConnector
from excel_interface.connectors.govmap import GovmapConnector
from excel_interface.connectors.xplan import XplanConnector
from excel_interface.interface import ExcelDataInterface

logger = logging.getLogger(__name__)

# Template path — relative to this file's package parent
TEMPLATE_PATH = Path(__file__).parent.parent / "urban_renewal_template.xlsx"

# Sources that require authentication / payment — shown as manual in the UI
MANUAL_REQUIRED = [
    {
        "name":   "מבט — מנהל התכנון (בדיקה ידנית)",
        "fields": ["PLAN_001", "PLAN_002", "PLAN_003"],
        "reason": "XPlan אוסף אוטומטית — מבט לאימות ידני ופרטי תכנית מלאים",
        "url":    "https://mavat.iplan.gov.il/",
    },
    {
        "name":   "XPlan — קווים כחולים",
        "fields": ["PLAN_001", "PLAN_002", "PLAN_003"],
        "reason": "ממשק GIS לבדיקת שכבות תכנון ויעודי קרקע",
        "url":    "https://ags.iplan.gov.il/xplan/",
    },
    {
        "name":   "טאבו / לשכת רישום מקרקעין",
        "fields": ["PARCEL_001"],
        "reason": "שירות בתשלום — ספקים מורשים: TabuDirect, דאטה פקס",
        "url":    "https://www.gov.il/he/service/tabu_land_registration",
    },
    {
        "name":   "RishunNet / תיק בניין",
        "fields": ["BUILD_001", "BUILD_002", "BUILD_003"],
        "reason": "דורש כניסה לארכיון הנדסה עירוני",
        "url":    "https://www.iplan.gov.il/",
    },
]


@dataclass
class CollectionSummary:
    """Result of a full collection run.

    Attributes
    ----------
    results:
        One CollectionResult per connector (including failed ones).
    output_path:
        Path to the saved `.xlsx` workbook.
    manual_required:
        List of dicts describing sources that need manual entry.
    city:
        The city name that was queried.
    address:
        The address that was queried.
    neighborhood:
        Optional neighborhood name (used in cadastral mode).
    search_mode:
        One of ``"address"`` | ``"cadastral"`` | ``"streets"``.
    item_label:
        Short label for this item in multi-item results (e.g. "חלקה 5" or "דיזנגוף").
    """
    results:         List[CollectionResult]
    output_path:     Path
    manual_required: List[dict]
    city:            str
    address:         str
    neighborhood:    str = ""
    search_mode:     str = "address"   # "address" | "cadastral" | "streets"
    item_label:      str = ""          # short label for multi-item display

    @property
    def ok_count(self) -> int:
        return sum(1 for r in self.results if r.status == CollectionStatus.OK)

    @property
    def partial_count(self) -> int:
        return sum(1 for r in self.results if r.status == CollectionStatus.PARTIAL)

    @property
    def failed_count(self) -> int:
        return sum(1 for r in self.results
                   if r.status == CollectionStatus.UNAVAILABLE)

    @property
    def total_fields_collected(self) -> int:
        return sum(len(r.fields_collected) for r in self.results)


class CollectionEngine:
    """Orchestrate data collection from all free connectors.

    Parameters
    ----------
    output_dir:
        Directory where the output workbook is saved.
        Defaults to the same directory as the template.
    """

    def __init__(self, output_dir: Optional[Path] = None) -> None:
        self.output_dir = output_dir or TEMPLATE_PATH.parent
        self._connectors: List[BaseConnector] = [
            DataGovConnector(),
            GovmapConnector(),
            CBSConnector(),
            XplanConnector(),
        ]

    def run(
        self,
        city: str,
        address: str,
        block:  Optional[str] = None,
        parcel: Optional[str] = None,
        output_path: Optional[Path] = None,
        progress_callback=None,
    ) -> CollectionSummary:
        """Run all connectors and save results to a new Excel workbook.

        Parameters
        ----------
        city:
            City name in Hebrew.
        address:
            Street address in Hebrew.
        block:
            Optional gush number.
        parcel:
            Optional parcel number.
        output_path:
            Explicit output path.  Defaults to
            ``urban_renewal_<city>_<YYYYMMDD>.xlsx`` in output_dir.
        progress_callback:
            Optional callable(step: int, total: int, message: str) for
            streaming progress to the Streamlit UI.

        Returns
        -------
        CollectionSummary
        """
        if not TEMPLATE_PATH.exists():
            raise FileNotFoundError(
                f"Template not found: {TEMPLATE_PATH}. "
                "Run `python -m excel_interface.create_template` first."
            )

        # ── Determine output path ────────────────────────────────────
        if output_path is None:
            safe_city = city.replace(" ", "_").replace("/", "-")
            stamp = date.today().strftime("%Y%m%d")
            output_path = self.output_dir / f"urban_renewal_{safe_city}_{stamp}.xlsx"

        # ── Load a fresh workbook from the template ──────────────────
        iface = ExcelDataInterface()
        iface.load(TEMPLATE_PATH)

        # Write the city name immediately (it's a definite known value)
        iface.write({"SITE_001": city}, source_name="User Input")

        results: List[CollectionResult] = []
        total = len(self._connectors)

        # ── Run each connector ───────────────────────────────────────
        for i, connector in enumerate(self._connectors, start=1):
            if progress_callback:
                progress_callback(i, total, f"מריץ {connector.name}…")

            try:
                result = connector.fetch(
                    city=city,
                    address=address,
                    block=block,
                    parcel=parcel,
                )
            except Exception as exc:
                logger.exception("Connector %s raised unexpected exception", connector.name)
                from excel_interface.connectors.base import CollectionResult as CR
                result = CR(
                    source_name = connector.name,
                    field_data  = {},
                    status      = CollectionStatus.UNAVAILABLE,
                    notes       = f"שגיאה לא צפויה: {exc}",
                )

            results.append(result)

            # Write collected fields to workbook (skip empty results)
            if result.field_data:
                try:
                    iface.write(result.field_data, source_name=result.source_name)
                except Exception as exc:
                    logger.warning("Failed to write result from %s: %s", connector.name, exc)

        # ── Save workbook ─────────────────────────────────────────────
        if progress_callback:
            progress_callback(total, total, "שומר קובץ Excel…")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        iface.save(output_path)

        return CollectionSummary(
            results         = results,
            output_path     = output_path,
            manual_required = MANUAL_REQUIRED,
            city            = city,
            address         = address,
            search_mode     = "address",
        )

    # ------------------------------------------------------------------
    # Cadastral mode — one run per parcel
    # ------------------------------------------------------------------

    def run_cadastral(
        self,
        city: str,
        block: str,
        parcels: List[str],
        neighborhood: str = "",
        output_dir: Optional[Path] = None,
        progress_callback: Optional[Callable] = None,
    ) -> List[CollectionSummary]:
        """Run collection for each parcel in *parcels*.

        Parameters
        ----------
        city:           City name in Hebrew.
        block:          Gush number (גוש).
        parcels:        List of parcel numbers (חלקות).
        neighborhood:   Optional neighborhood name (hint for geocoding).
        output_dir:     Directory for output workbooks. Defaults to ``self.output_dir``.
        progress_callback: callable(step, total, message)

        Returns
        -------
        list[CollectionSummary] — one entry per parcel.
        """
        out_dir = output_dir or self.output_dir
        results = []
        total = len(parcels)

        for i, parcel in enumerate(parcels, start=1):
            label = f"חלקה {parcel} (גוש {block})"
            if progress_callback:
                progress_callback(i, total, f"אוסף נתונים עבור {label}…")

            # Build a descriptive address for geocoding
            address_hint = f"גוש {block} חלקה {parcel}"
            if neighborhood:
                address_hint = f"{neighborhood}, {address_hint}"

            safe = re.sub(r"[^\w]", "_", f"{city}_g{block}_p{parcel}")
            stamp = date.today().strftime("%Y%m%d")
            out_path = out_dir / f"urban_renewal_{safe}_{stamp}.xlsx"

            summary = self.run(
                city         = city,
                address      = address_hint,
                block        = block,
                parcel       = parcel,
                output_path  = out_path,
            )
            summary.neighborhood = neighborhood
            summary.search_mode  = "cadastral"
            summary.item_label   = label
            results.append(summary)

        return results

    # ------------------------------------------------------------------
    # Multi-street mode — one run per street
    # ------------------------------------------------------------------

    def run_streets(
        self,
        city: str,
        streets: List[dict],
        output_dir: Optional[Path] = None,
        progress_callback: Optional[Callable] = None,
    ) -> List[CollectionSummary]:
        """Run collection for each street in *streets*.

        Parameters
        ----------
        city:     City name in Hebrew.
        streets:  List of dicts: ``{"street": "אבן גבירול", "numbers": "1-10"}``
                  ``numbers`` is optional.
        output_dir:         Directory for output workbooks.
        progress_callback:  callable(step, total, message)

        Returns
        -------
        list[CollectionSummary] — one entry per street.
        """
        out_dir = output_dir or self.output_dir
        results = []
        total = len(streets)

        for i, item in enumerate(streets, start=1):
            street  = item.get("street", "").strip()
            numbers = item.get("numbers", "").strip()
            if not street:
                continue

            label   = f"{street} {numbers}".strip()
            address = label  # full address string passed to geocoder

            if progress_callback:
                progress_callback(i, total, f"אוסף נתונים עבור {label}…")

            safe  = re.sub(r"[^\w]", "_", f"{city}_{street}")[:50]
            stamp = date.today().strftime("%Y%m%d")
            out_path = out_dir / f"urban_renewal_{safe}_{stamp}.xlsx"

            summary = self.run(
                city        = city,
                address     = address,
                output_path = out_path,
            )
            summary.search_mode = "streets"
            summary.item_label  = label
            results.append(summary)

        return results
