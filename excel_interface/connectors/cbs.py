"""
cbs.py — CBSConnector

Fetches demographic and population statistics from the Israeli Central Bureau
of Statistics (הלשכה המרכזית לסטטיסטיקה / CBS) open data.

CBS provides open GeoJSON and CSV datasets for statistical areas.
We use data.gov.il as a proxy for CBS datasets (CBS publishes on that portal).

Dataset used
------------
  Population by city (from CBS via data.gov.il)
  Resource: אוכלוסייה לפי יישוב
  Endpoint: https://data.gov.il/api/3/action/datastore_search

Populates
---------
  SITE_005 — שטח מתחם כולל (estimated city area from CBS, in sqm)
             Note: this is the CITY area, not the specific site — flagged as "estimate"
"""

from __future__ import annotations

import logging
from typing import Optional
from urllib.parse import urlencode

import requests

from .base import BaseConnector, CollectionResult

logger = logging.getLogger(__name__)

DATA_GOV_BASE   = "https://data.gov.il/api/3/action/datastore_search"

# CBS population/area by settlement (resource published by CBS on data.gov.il)
# Source: https://data.gov.il/dataset/residents_in_israel_by_communities_and_age_groups
CBS_POPULATION_RESOURCE = "64edd0ee-3d5d-43ce-8562-c336c24dbc1f"

TIMEOUT = 10


class CBSConnector(BaseConnector):
    """Fetch city-level demographic statistics from CBS via data.gov.il."""

    name = "CBS / הלמ\"ס"
    required_fields = ["SITE_005"]

    def fetch(
        self,
        city: str,
        address: str,
        block: Optional[str] = None,
        parcel: Optional[str] = None,
    ) -> CollectionResult:
        try:
            record = self._search_city_stats(city)
        except requests.Timeout:
            return self._unavailable("CBS / data.gov.il לא הגיב בזמן (timeout 10s).")
        except requests.ConnectionError as exc:
            return self._unavailable(f"שגיאת חיבור ל-CBS: {exc}")
        except Exception as exc:
            return self._unavailable(f"שגיאה לא צפויה: {exc}")

        if record is None:
            return self._partial(
                {},
                f"העיר '{city}' לא נמצאה בנתוני CBS. "
                "שדה SITE_005 ידרוש מילוי ידני.",
            )

        field_data: dict = {}
        fields_found = []

        # שטח העיר — CBS field names vary; try common variants
        area_val = (
            record.get("שטח")
            or record.get("area")
            or record.get("AREA")
            or record.get("area_km2")
            or record.get("שטח_קמ2")
        )
        if area_val:
            try:
                area_sqm = float(str(area_val).replace(",", "")) * 1_000_000
                field_data["SITE_005"] = round(area_sqm, 0)
                fields_found.append("SITE_005")
            except (TypeError, ValueError):
                pass

        if not field_data:
            return self._partial(
                {},
                f"נמצא רשומת CBS עבור '{city}' אך שדה שטח לא זמין.",
            )

        from .base import CollectionStatus
        return CollectionResult(
            source_name      = self.name,
            field_data       = field_data,
            status           = CollectionStatus.PARTIAL,
            notes            = (
                f"שטח עיר מ-CBS: {field_data.get('SITE_005', '')} מ\"ר "
                "(הערכה ברמת עיר — יש לעדכן לשטח המתחם הספציפי)."
            ),
            fields_collected = fields_found,
        )

    def _search_city_stats(self, city: str) -> dict | None:
        """Search CBS dataset for *city*. Return first matching record or None."""
        params = {
            "resource_id": CBS_POPULATION_RESOURCE,
            "q":           city,
            "limit":       5,
        }
        url = f"{DATA_GOV_BASE}?{urlencode(params)}"
        resp = requests.get(url, timeout=TIMEOUT)
        resp.raise_for_status()
        data = resp.json()

        if not data.get("success"):
            raise ValueError(f"data.gov.il/CBS returned success=false: {data.get('error')}")

        records = data.get("result", {}).get("records", [])
        if not records:
            return None

        city_lower = city.strip().lower()
        for rec in records:
            for v in rec.values():
                if isinstance(v, str) and v.strip().lower() == city_lower:
                    return rec
        return records[0]
