"""
data_gov.py — DataGovConnector

Fetches city-level administrative metadata from the Israeli open data portal
(data.gov.il) using the CKAN datastore API.

Dataset used
------------
  Resource: מחוזות ונפות — city/district code lookup
  Endpoint: https://data.gov.il/api/3/action/datastore_search

Populates
---------
  SITE_001 — שם עיר (confirmed city name from the dataset)
  SITE_002 — שכונה / מחוז (district name)
"""

from __future__ import annotations

import logging
from typing import Optional
from urllib.parse import urlencode

import requests

from .base import BaseConnector, CollectionResult

logger = logging.getLogger(__name__)

# data.gov.il CKAN API
DATA_GOV_BASE = "https://data.gov.il/api/3/action/datastore_search"

# Resource ID for Israeli city list (ישובים ברשות המקומית)
# Source: https://data.gov.il/dataset/citiesandsettelments
CITIES_RESOURCE_ID = "5c78e9fa-c2e2-4771-93ff-7f400a12f7ba"

TIMEOUT = 10


class DataGovConnector(BaseConnector):
    """Fetch city administrative data from data.gov.il."""

    name = "data.gov.il"
    required_fields = ["SITE_001", "SITE_002"]

    def fetch(
        self,
        city: str,
        address: str,
        block: Optional[str] = None,
        parcel: Optional[str] = None,
    ) -> CollectionResult:
        try:
            record = self._search_city(city)
        except requests.Timeout:
            return self._unavailable("data.gov.il לא הגיב בזמן (timeout 10s).")
        except requests.ConnectionError as exc:
            return self._unavailable(f"שגיאת חיבור ל-data.gov.il: {exc}")
        except Exception as exc:
            return self._unavailable(f"שגיאה לא צפויה: {exc}")

        if record is None:
            return self._partial(
                {},
                f"העיר '{city}' לא נמצאה ב-data.gov.il. "
                "ייתכן שהשם שגוי או שלא קיים במאגר.",
            )

        field_data: dict = {}
        fields_found = []

        # שם עיר רשמי
        city_name = (
            record.get("שם_ישוב")
            or record.get("name_he")
            or record.get("city_name_he")
            or city
        )
        field_data["SITE_001"] = city_name.strip()
        fields_found.append("SITE_001")

        # מחוז / נפה
        district = (
            record.get("מחוז")
            or record.get("district")
            or record.get("nafa")
            or ""
        )
        if district:
            field_data["SITE_002"] = str(district).strip()
            fields_found.append("SITE_002")

        missing = [f for f in self.required_fields if f not in field_data]
        from .base import CollectionStatus
        status = CollectionStatus.OK if not missing else CollectionStatus.PARTIAL
        notes  = (
            f"נמצאו נתוני עיר עבור '{city_name}'."
            if not missing
            else f"נמצא חלקי: חסרים {', '.join(missing)}."
        )
        return CollectionResult(
            source_name      = self.name,
            field_data       = field_data,
            status           = status,
            notes            = notes,
            fields_collected = fields_found,
        )

    def _search_city(self, city: str) -> dict | None:
        """Search the cities dataset for *city*. Return first matching record or None."""
        params = {
            "resource_id": CITIES_RESOURCE_ID,
            "q":           city,
            "limit":       5,
        }
        url = f"{DATA_GOV_BASE}?{urlencode(params)}"
        resp = requests.get(url, timeout=TIMEOUT)
        resp.raise_for_status()
        data = resp.json()

        if not data.get("success"):
            raise ValueError(f"data.gov.il returned success=false: {data.get('error')}")

        records = data.get("result", {}).get("records", [])
        if not records:
            return None

        # Try exact match first, then partial
        city_lower = city.strip().lower()
        for rec in records:
            for v in rec.values():
                if isinstance(v, str) and v.strip().lower() == city_lower:
                    return rec
        # Fallback: return first result
        return records[0]
