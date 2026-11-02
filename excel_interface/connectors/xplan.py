"""
xplan.py — XplanConnector

Fetches planning data from the Israeli planning authority XPlan viewer
(ags.iplan.gov.il/xplan) using a headless Chromium browser session.

The public ArcGIS REST API requires a valid browser session (cookies) from
the XPlan viewer page — direct HTTP requests without a session are rejected.
This connector uses Playwright to establish the session and then calls the
ArcGIS REST API from within the browser context.

Two-step process
----------------
1. Geocode the address via Nominatim → lat, lon
2. Open XPlan in headless Chromium to establish a valid session
3. Query the ArcGIS REST layers via page.evaluate (fetch with credentials)
   - Layer 1: תכניות מקוונות (plans) → PLAN_001, PLAN_002, PLAN_003
   - Layer 4: יעודי קרקע (land use) → PLAN_003 enrichment

Populates
---------
  PLAN_001 — Plan_Number (מספר תכנית — most recent active plan)
  PLAN_002 — Plan_Status (סטטוס תכנית)
  PLAN_003 — Land_Use (ייעוד קרקע)
"""

from __future__ import annotations

import logging
from typing import Optional

from .base import BaseConnector, CollectionResult
from .playwright_utils import INSTALL_NOTE, get_page, is_playwright_available

logger = logging.getLogger(__name__)

# Base URL for the XPlan ArcGIS REST (internal path discovered from network intercept)
XPLAN_BASE = (
    "https://ags.iplan.gov.il/arcgisiplan/rest/services/PlanningPublic/Xplan/MapServer"
)

# Status values considered "active" — preferred for PLAN_001/002
_ACTIVE_STATUSES = {"התכנית אושרה", "פרסום אישור", "בתוקף"}


class XplanConnector(BaseConnector):
    """Fetch planning data from XPlan via headless browser session."""

    name = "XPlan / מנהל התכנון"
    required_fields = ["PLAN_001", "PLAN_002", "PLAN_003", "SITE_003", "PARCEL_002", "PARCEL_003", "PARCEL_004"]

    def fetch(
        self,
        city: str,
        address: str,
        block: Optional[str] = None,
        parcel: Optional[str] = None,
    ) -> CollectionResult:

        if not is_playwright_available():
            return self._unavailable(INSTALL_NOTE)

        # Geocode + query — all inside headless browser to bypass IP rate-limits
        try:
            result = self._geocode_and_query(city, address)
        except Exception as exc:
            logger.exception("XplanConnector failed")
            return self._unavailable(f"שגיאה ב-XPlan: {exc}")

        if result.get("error") == "geocode_not_found":
            return self._unavailable(
                f"לא נמצאה כתובת: '{address}, {city}'. נסה לדייק."
            )

        plans     = result.get("plans", [])
        land_uses = result.get("land_uses", [])
        lat       = result.get("lat", 0)
        lon       = result.get("lon", 0)

        if not plans and not land_uses:
            return self._partial(
                {},
                f"לא נמצאו תכניות בנקודה זו ב-XPlan ({lat:.4f}, {lon:.4f}).",
            )

        # Step 3: Build field_data
        field_data: dict = {}
        fields_found = []

        # Address / location from geocoding
        display_name = result.get("display_name", "")
        if display_name:
            field_data["SITE_003"] = display_name
            fields_found.append("SITE_003")

        # Parcel data from layer 4 land-use records (gush/helka/area)
        if land_uses:
            lu = land_uses[0]
            # layer 4 has pl_number as gush in some cases — use plan-based gush from result
            gush  = result.get("gush")
            helka = result.get("helka")
            area  = result.get("area")
            if gush:
                field_data["PARCEL_002"] = str(gush)
                fields_found.append("PARCEL_002")
            if helka:
                field_data["PARCEL_003"] = str(helka)
                fields_found.append("PARCEL_003")
            if area:
                try:
                    field_data["PARCEL_004"] = float(area)
                    fields_found.append("PARCEL_004")
                except (TypeError, ValueError):
                    pass

        # Pick the most relevant plan (prefer approved/active)
        chosen_plan = None
        for p in plans:
            status = p.get("internet_short_status", "") or ""
            if any(s in status for s in _ACTIVE_STATUSES):
                chosen_plan = p
                break
        if chosen_plan is None and plans:
            chosen_plan = plans[0]

        if chosen_plan:
            pl_num = chosen_plan.get("pl_number")
            if pl_num:
                field_data["PLAN_001"] = str(pl_num)
                fields_found.append("PLAN_001")

            pl_status = chosen_plan.get("internet_short_status")
            if pl_status:
                field_data["PLAN_002"] = str(pl_status)
                fields_found.append("PLAN_002")

            # Land use: prefer from layer 4 (more specific), fallback layer 1
            land_use = (
                land_uses[0].get("mavat_name") if land_uses
                else chosen_plan.get("pl_landuse_string")
            )
            if land_use:
                field_data["PLAN_003"] = str(land_use)
                fields_found.append("PLAN_003")

        plan_fields = ["PLAN_001", "PLAN_002", "PLAN_003"]
        missing = [f for f in plan_fields if f not in field_data]
        from .base import CollectionStatus
        status_enum = CollectionStatus.OK if not missing else CollectionStatus.PARTIAL
        notes = (
            f"נמצאו {len(plans)} תכניות ב-XPlan. "
            f"תכנית ראשית: {field_data.get('PLAN_001', '—')}."
            if not missing
            else f"נמצאו {len(plans)} תכניות; חסרים: {', '.join(missing)}."
        )
        return CollectionResult(
            source_name=self.name,
            field_data=field_data,
            status=status_enum,
            notes=notes,
            fields_collected=fields_found,
        )

    # ------------------------------------------------------------------

    def _geocode_and_query(self, city: str, address: str):
        """Geocode address via Nominatim and query XPlan directly via HTTP.
        No browser needed — the ArcGIS REST API is publicly accessible."""
        import ssl
        import certifi
        import requests as _req
        from requests.adapters import HTTPAdapter

        # TLS adapter for ags.iplan.gov.il (TLS 1.2 + AES128-SHA)
        class _TLSAdapter(HTTPAdapter):
            def init_poolmanager(self, *args, **kwargs):
                ctx = ssl.create_default_context(cafile=certifi.where())
                ctx.set_ciphers("DEFAULT")
                kwargs["ssl_context"] = ctx
                super().init_poolmanager(*args, **kwargs)

        session = _req.Session()
        session.mount("https://ags.iplan.gov.il/", _TLSAdapter())
        session.headers["User-Agent"] = "UrbanRenewalAI/1.0"

        nom_headers = {
            "User-Agent": "Mozilla/5.0 (compatible; UrbanRenewalAI/1.0)",
            "Accept": "application/json",
        }

        # Step 1: Geocode via Nominatim
        clean_address = address
        for prefix in ("רחוב ", "רח' ", "רח ", "שדרות ", "שד' ", "שד "):
            if clean_address.startswith(prefix):
                clean_address = clean_address[len(prefix):]
                break

        lat = lon = display_name = None
        for query in list(dict.fromkeys([
            f"{address}, {city}, ישראל",
            f"{clean_address}, {city}, ישראל",
            f"{city}, {clean_address}, ישראל",
        ])):
            try:
                r = _req.get(
                    "https://nominatim.openstreetmap.org/search",
                    params={"q": query, "format": "json", "limit": 1, "accept-language": "he"},
                    headers=nom_headers, timeout=10,
                )
                data = r.json()
                if data:
                    lat = float(data[0]["lat"])
                    lon = float(data[0]["lon"])
                    display_name = data[0].get("display_name", "")
                    break
            except Exception:
                continue

        if not lat:
            return {"error": "geocode_not_found"}

        geo = f"{lon},{lat}"

        # Step 2: Query XPlan layer 1 (plan boundaries) via HTTP
        plans = []
        land_uses = []
        try:
            r1 = session.get(
                f"{XPLAN_BASE}/1/query",
                params={
                    "geometry": geo, "geometryType": "esriGeometryPoint",
                    "inSR": "4326", "spatialRel": "esriSpatialRelIntersects",
                    "outFields": "pl_number,pl_name,station_desc,internet_short_status",
                    "returnGeometry": "false", "f": "json", "resultRecordCount": "10",
                },
                timeout=20,
            )
            body1 = r1.json()
            if "error" not in body1:
                plans = [f["attributes"] for f in body1.get("features", [])]
        except Exception:
            pass

        # Step 3: Query XPlan layer 4 (land use) via HTTP
        try:
            r4 = session.get(
                f"{XPLAN_BASE}/4/query",
                params={
                    "geometry": geo, "geometryType": "esriGeometryPoint",
                    "inSR": "4326", "spatialRel": "esriSpatialRelIntersects",
                    "outFields": "mavat_name,pl_number,legal_area,station_desc,shape_area",
                    "returnGeometry": "false", "f": "json", "resultRecordCount": "3",
                },
                timeout=20,
            )
            body4 = r4.json()
            if "error" not in body4:
                land_uses = [f["attributes"] for f in body4.get("features", [])]
        except Exception:
            pass

        return {
            "lat":          lat,
            "lon":          lon,
            "display_name": display_name or f"{address}, {city}",
            "plans":        plans,
            "land_uses":    land_uses,
            "gush":         None,
            "helka":        None,
            "area":         land_uses[0].get("shape_area") if land_uses else None,
        }
