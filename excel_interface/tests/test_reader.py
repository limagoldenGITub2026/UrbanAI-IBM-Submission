"""
test_reader.py — Unit tests for ExcelReader

Run from the 2026.11.02/ directory:
    python -m pytest excel_interface/tests/test_reader.py -v
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from excel_interface.reader import ExcelReader

# Path to the canonical template — tests work on a temporary copy
TEMPLATE = Path(__file__).parent.parent.parent / "urban_renewal_template.xlsx"


@pytest.fixture
def tmp_workbook(tmp_path: Path) -> Path:
    """Copy the template to a temp directory and return the copy path."""
    dest = tmp_path / "urban_renewal_template.xlsx"
    shutil.copy(TEMPLATE, dest)
    return dest


# ── load_data_mapping ──────────────────────────────────────────────────────────

class TestLoadDataMapping:
    def test_returns_17_entries(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        mapping = reader.load_data_mapping()
        assert len(mapping) == 17, f"Expected 17 Field_IDs, got {len(mapping)}"

    def test_stores_on_instance(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        reader.load_data_mapping()
        assert len(reader.field_locations) == 17

    def test_field_location_attrs(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        mapping = reader.load_data_mapping()
        site001 = mapping["SITE_001"]
        assert site001.sheet_name == "SITE_DATA"
        assert site001.table_name == "tbl_site_data"
        assert site001.column_name == "Value"
        assert site001.data_type == "text"
        assert site001.required is True
        assert site001.write_permission is True

    def test_parcel_001_write_permission_false(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        mapping = reader.load_data_mapping()
        assert mapping["PARCEL_001"].write_permission is False

    def test_raises_on_missing_data_mapping_sheet(self, tmp_workbook):
        import openpyxl
        wb = openpyxl.load_workbook(str(tmp_workbook))
        del wb["DATA_MAPPING"]
        wb.save(str(tmp_workbook))
        reader = ExcelReader(tmp_workbook)
        with pytest.raises(KeyError, match="DATA_MAPPING"):
            reader.load_data_mapping()

    def test_populates_sheet_names_correctly(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        mapping = reader.load_data_mapping()
        sheets = {loc.sheet_name for loc in mapping.values()}
        assert sheets == {"SITE_DATA", "PARCELS", "PLANS", "BUILDINGS", "CONSTRAINTS", "SOURCES_LOG"}


# ── sheet_names ────────────────────────────────────────────────────────────────

class TestSheetNames:
    def test_all_8_sheets_present(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        names = reader.sheet_names()
        expected = {"README", "SITE_DATA", "PARCELS", "PLANS", "BUILDINGS",
                    "CONSTRAINTS", "SOURCES_LOG", "DATA_MAPPING"}
        assert set(names) == expected


# ── read_sheet ─────────────────────────────────────────────────────────────────

class TestReadSheet:
    def test_site_data_returns_5_rows(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        rows = reader.read_sheet("SITE_DATA")
        assert len(rows) == 5

    def test_site_data_row_has_field_id_and_value(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        rows = reader.read_sheet("SITE_DATA")
        first = rows[0]
        assert "Field_ID" in first
        assert "Value" in first
        assert "Status" in first

    def test_parcels_returns_0_data_rows(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        rows = reader.read_sheet("PARCELS")
        assert rows == [], f"Expected 0 rows, got {len(rows)}"

    def test_plans_returns_0_data_rows(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        rows = reader.read_sheet("PLANS")
        assert rows == []

    def test_data_mapping_returns_17_rows(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        rows = reader.read_sheet("DATA_MAPPING")
        assert len(rows) == 17

    def test_raises_on_unknown_sheet(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        with pytest.raises(KeyError, match="NONEXISTENT"):
            reader.read_sheet("NONEXISTENT")

    def test_parcels_column_headers(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        headers = reader.column_headers("PARCELS")
        assert headers == ["Parcel_ID", "Block_Number", "Parcel_Number", "Area_sqm", "Source_ID"]


# ── get_record_by_id ───────────────────────────────────────────────────────────

class TestGetRecordById:
    def test_returns_none_on_empty_sheet(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        result = reader.get_record_by_id("PARCELS", "Parcel_ID", "P_001")
        assert result is None

    def test_finds_existing_record(self, tmp_workbook):
        """Add a row manually and confirm get_record_by_id finds it."""
        import openpyxl
        wb = openpyxl.load_workbook(str(tmp_workbook))
        ws = wb["PARCELS"]
        ws.append(["P_6660_001", "6660", "1", 450.0, "SRC_001"])
        wb.save(str(tmp_workbook))

        reader = ExcelReader(tmp_workbook)
        row = reader.get_record_by_id("PARCELS", "Parcel_ID", "P_6660_001")
        assert row is not None
        assert row["Block_Number"] == "6660"
        assert row["Area_sqm"] == 450.0

    def test_returns_none_for_wrong_id(self, tmp_workbook):
        import openpyxl
        wb = openpyxl.load_workbook(str(tmp_workbook))
        ws = wb["PARCELS"]
        ws.append(["P_6660_001", "6660", "1", 450.0, "SRC_001"])
        wb.save(str(tmp_workbook))

        reader = ExcelReader(tmp_workbook)
        result = reader.get_record_by_id("PARCELS", "Parcel_ID", "P_NOTEXIST")
        assert result is None


# ── identify_missing_fields ────────────────────────────────────────────────────

class TestIdentifyMissingFields:
    def test_fresh_template_missing_site_001(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        reader.load_data_mapping()
        missing = reader.identify_missing_fields()
        # SITE_001 is required and writable; SITE_DATA Value is empty in the template
        assert "SITE_001" in missing

    def test_raises_without_load_data_mapping(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        with pytest.raises(RuntimeError, match="load_data_mapping"):
            reader.identify_missing_fields()

    def test_after_filling_site_001_not_missing(self, tmp_workbook):
        import openpyxl
        wb = openpyxl.load_workbook(str(tmp_workbook))
        ws = wb["SITE_DATA"]
        # Row 2 is SITE_001; column C is Value (index 3)
        for row in ws.iter_rows(min_row=2):
            if row[0].value == "SITE_001":
                row[2].value = "Tel Aviv"
                break
        wb.save(str(tmp_workbook))

        reader = ExcelReader(tmp_workbook)
        reader.load_data_mapping()
        missing = reader.identify_missing_fields()
        assert "SITE_001" not in missing

    def test_empty_parcels_sheet_flags_parcel_required_fields(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        reader.load_data_mapping()
        missing = reader.identify_missing_fields()
        # PARCEL_002 and PARCEL_003 are required; PARCELS sheet is empty
        assert "PARCEL_002" in missing
        assert "PARCEL_003" in missing

    def test_parcel_001_not_in_missing_write_permission_false(self, tmp_workbook):
        reader = ExcelReader(tmp_workbook)
        reader.load_data_mapping()
        missing = reader.identify_missing_fields()
        # PARCEL_001 has write_permission=False → must NOT appear in missing list
        assert "PARCEL_001" not in missing


# ── FileNotFoundError ──────────────────────────────────────────────────────────

class TestFileNotFound:
    def test_raises_on_missing_file(self):
        with pytest.raises(FileNotFoundError):
            ExcelReader("/nonexistent/path/to/file.xlsx")
