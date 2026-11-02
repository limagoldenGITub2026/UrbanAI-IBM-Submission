"""
govmap.py — GovmapConnector

Collects parcel data (gush, parcel number, area) for a given Hebrew address.

Two-step process (all inside headless Chromium to bypass IP rate-limits)
------------------------------------------------------------------------
1. Open govmap.gov.il to establish a valid browser session
2. Geocode via Nominatim from within the browser (browser IP, no rate-limit)
3. Query Govmap ArcGIS parcel layer with browser session cookies

Populates
---------
  SITE_003   — address as confirmed by geocoder
  PARCEL_002 — Block_Number (גוש)
  PARCEL_003 — Parcel_Number (חלקה)
  PARCEL_004 — Area_sqm (שטח)
"""

from __future__ import annotations

import logging
from typing import Optional

import requests

from .base import BaseConnector, CollectionResult
from .playwright_utils import INSTALL_NOTE, get_page, is_playwright_available

logger = logging.getLogger(__name__)

NOMINATIM_URL    = "https://nominatim.openstreetmap.org/search"
NOMINATIM_HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; UrbanRenewalAI/1.0; +https://github.com/urban-renewal)",
    "Accept-Language": "he,en;q=0.9",
    "Accept": "application/json",
}
TIMEOUT = 15

# Govmap ArcGIS REST — public parcel layer (CADASTRE)
GOVMAP_PARCEL_URL = (
    "https://ags.govmap.gov.il/arcgis/rest/services/CADASTRE/Parcels/MapServer/0/query"
)


class GovmapConnector(BaseConnector):
    """Fetch parcel data from Govmap using Hebrew address geocoding."""

    name = "Govmap / OpenStreetMap"
    required_fields = ["SITE_003", "PARCEL_002", "PARCEL_003", "PARCEL_004"]

    def fetch(
        self,
        city: str,
        address: str,
        block: Optional[str] = None,
        parcel: Optional[str] = None,
    ) -> CollectionResult:

        # Try direct HTTP first (fast) — Playwright only as last resort
        try:
            result = self._fetch_direct(city, address)
            if result.status.value != "unavailable":
                return result
        except Exception:
            pass

        # Fallback: browser-based geocoding (slower, bypasses IP rate-limit)
        if is_playwright_available():
            try:
                return self._fetch_via_browser(city, address)
            except Exception as exc:
                return self._unavailable(f"Govmap לא זמין: {exc}")

        return self._unavailable("Govmap לא זמין")

    # ------------------------------------------------------------------
    # Browser path (Playwright)
    # ------------------------------------------------------------------

    def _fetch_via_browser(self, city: str, address: str) -> CollectionResult:
        """Geocode via Nominatim inside headless browser (avoids IP rate-limit),
        then query Govmap ArcGIS directly from Python (no CORS issue)."""
        # Strip common Hebrew prefixes that confuse geocoders
        clean_address = address
        for prefix in ("רחוב ", "רח' ", "רח ", "שדרות ", "שד' ", "שד "):
            if clean_address.startswith(prefix):
                clean_address = clean_address[len(prefix):]
                break

        # Try both original and cleaned address
        queries = list(dict.fromkeys([
            f"{address}, {city}, ישראל",
            f"{clean_address}, {city}, ישראל",
            f"{city}, {clean_address}, ישראל",
        ]))

        # Build JS that tries each query in order until one succeeds
        queries_js = str(queries).replace("'", '"')
        nom_url = (
            f"https://nominatim.openstreetmap.org/search"
            f"?q=QUERY_PLACEHOLDER"
            f"&format=json&limit=1&addressdetails=1&accept-language=he"
        )

        # Step 1: geocode from within the browser — try multiple query variants
        base_url = "https://nominatim.openstreetmap.org/search?format=json&limit=1&addressdetails=1&accept-language=he&q="
        js = f"""async () => {{
            const queries = {queries_js};
            for (const q of queries) {{
                const url = "{base_url}" + encodeURIComponent(q);
                const resp = await fetch(url, {{headers: {{"User-Agent": "UrbanRenewalAI/1.0"}}}});
                const data = await resp.json();
                if (data && data.length) {{
                    return {{
                        lat:          parseFloat(data[0].lat),
                        lon:          parseFloat(data[0].lon),
                        display_name: data[0].display_name || ""
                    }};
                }}
            }}
            return {{error: "not_found"}};
        }}"""

        try:
            with get_page(timeout_ms=30_000) as page:
                page.goto("https://www.govmap.gov.il/", wait_until="domcontentloaded", timeout=25_000)
                page.wait_for_timeout(500)
                geo = page.evaluate(js)
        except Exception as exc:
            return self._unavailable(f"שגיאה בגיאוקודינג: {exc}")

        if geo.get("error") == "not_found":
            return self._unavailable(f"לא נמצאה כתובת: '{address}, {city}'.")

        lat          = geo["lat"]
        lon          = geo["lon"]
        display_name = geo.get("display_name", f"{address}, {city}")

        # Step 2: query Govmap ArcGIS parcel layer
        # Try Python direct first; if blocked, fall back to browser fetch
        parcel_attrs = self._query_parcel_python(lon, lat)
        if parcel_attrs is None:
            parcel_attrs = self._query_parcel_browser(lon, lat)

        return self._build_result(display_name, parcel_attrs, lat, lon)

    def _query_parcel_python(self, lon: float, lat: float):
        """Query Govmap ArcGIS directly via Python requests. Returns attrs dict or None."""
        try:
            from urllib.parse import urlencode
            params = {
                "geometry":       f"{lon},{lat}",
                "geometryType":   "esriGeometryPoint",
                "inSR":           "4326",
                "spatialRel":     "esriSpatialRelIntersects",
                "outFields":      "GUSH,HELKA,AREA",
                "returnGeometry": "false",
                "f":              "json",
            }
            resp = requests.get(f"{GOVMAP_PARCEL_URL}?{urlencode(params)}", timeout=TIMEOUT)
            resp.raise_for_status()
            features = resp.json().get("features", [])
            return features[0].get("attributes", {}) if features else None
        except Exception:
            return None

    def _query_parcel_browser(self, lon: float, lat: float):
        """Query Govmap ArcGIS from within browser context (bypasses network blocks)."""
        try:
            from urllib.parse import urlencode
            params = {
                "geometry": f"{lon},{lat}", "geometryType": "esriGeometryPoint",
                "inSR": "4326", "spatialRel": "esriSpatialRelIntersects",
                "outFields": "GUSH,HELKA,AREA", "returnGeometry": "false", "f": "json",
            }
            purl = f"{GOVMAP_PARCEL_URL}?{urlencode(params)}"
            js = f"""async () => {{
                const r = await fetch("{purl}");
                const d = await r.json();
                return (d.features && d.features.length) ? d.features[0].attributes : null;
            }}"""
            with get_page(timeout_ms=25_000) as page:
                page.goto("https://www.govmap.gov.il/", wait_until="domcontentloaded", timeout=20_000)
                return page.evaluate(js)
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Direct HTTP fallback
    # ------------------------------------------------------------------

    def _fetch_direct(self, city: str, address: str) -> CollectionResult:
        """Direct HTTP — may fail with 403 on some IPs."""
        from urllib.parse import urlencode
        query = f"{address}, {city}, ישראל"
        params = {"q": query, "format": "json", "limit": 1,
                  "addressdetails": 1, "accept-language": "he"}
        url = f"{NOMINATIM_URL}?{urlencode(params)}"
        try:
            resp = requests.get(url, headers=NOMINATIM_HEADERS, timeout=TIMEOUT)
            resp.raise_for_status()
            results = resp.json()
        except Exception as exc:
            return self._unavailable(f"Geocoding failed: {exc}")

        if not results:
            return self._unavailable(f"לא נמצאה כתובת: '{address}, {city}'.")

        r = results[0]
        lat, lon = float(r["lat"]), float(r["lon"])
        display_name = r.get("display_name", f"{address}, {city}")

        try:
            from urllib.parse import urlencode as ue
            params2 = {
                "geometry": f"{lon},{lat}", "geometryType": "esriGeometryPoint",
                "inSR": "4326", "spatialRel": "esriSpatialRelIntersects",
                "outFields": "GUSH,HELKA,AREA", "returnGeometry": "false", "f": "json",
            }
            resp2 = requests.get(f"{GOVMAP_PARCEL_URL}?{ue(params2)}", timeout=TIMEOUT)
            resp2.raise_for_status()
            features = resp2.json().get("features", [])
            parcel_attrs = features[0].get("attributes", {}) if features else None
        except Exception as exc:
            return self._partial(
                {"SITE_003": display_name},
                f"כתובת אותרה אך שליפת חלקה נכשלה: {exc}",
            )

        return self._build_result(display_name, parcel_attrs, lat, lon)

    # ------------------------------------------------------------------
    # Shared result builder
    # ------------------------------------------------------------------

    def _build_result(self, display_name: str, parcel_attrs, lat: float, lon: float) -> CollectionResult:
        if not parcel_attrs:
            return self._partial(
                {"SITE_003": display_name},
                f"כתובת אותרה ({display_name}) אך לא נמצאה חלקה קדסטרלית.",
            )

        field_data: dict = {"SITE_003": display_name}
        fields_found = ["SITE_003"]

        if parcel_attrs.get("GUSH"):
            field_data["PARCEL_002"] = str(parcel_attrs["GUSH"])
            fields_found.append("PARCEL_002")
        if parcel_attrs.get("HELKA"):
            field_data["PARCEL_003"] = str(parcel_attrs["HELKA"])
            fields_found.append("PARCEL_003")
        if parcel_attrs.get("AREA"):
            try:
                field_data["PARCEL_004"] = float(parcel_attrs["AREA"])
                fields_found.append("PARCEL_004")
            except (TypeError, ValueError):
                pass

        from .base import CollectionStatus
        missing = [f for f in self.required_fields if f not in field_data]
        status = CollectionStatus.OK if not missing else CollectionStatus.PARTIAL
        notes = (
            f"אותרו {len(fields_found)} שדות מ-Govmap."
            if not missing
            else f"אותרו {len(fields_found)} שדות; חסרים: {', '.join(missing)}."
        )
        return CollectionResult(
            source_name=self.name,
            field_data=field_data,
            status=status,
            notes=notes,
            fields_collected=fields_found,
        )
