"""
test_mapper.py — Unit tests for DataMapper

Run from the 2026.11.02/ directory:
    python -m pytest excel_interface/tests/test_mapper.py -v
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from excel_interface.mapper import DataMapper, WritePermissionError
from excel_interface.reader import ExcelReader

TEMPLATE = Path(__file__).parent.parent.parent / "urban_renewal_template.xlsx"


@pytest.fixture
def field_locations(tmp_path):
    """Load a fresh field_locations dict from the template."""
    dest = tmp_path / "template.xlsx"
    shutil.copy(TEMPLATE, dest)
    reader = ExcelReader(dest)
    return reader.load_data_mapping()


@pytest.fixture
def mapper(field_locations):
    return DataMapper(field_locations)


# ── map_input ──────────────────────────────────────────────────────────────────

class TestMapInput:
    def test_returns_write_operation_for_valid_field(self, mapper):
        ops = mapper.map_input({"SITE_001": "Haifa"})
        assert len(ops) == 1
        op = ops[0]
        assert op.field_id == "SITE_001"
        assert op.sheet_name == "SITE_DATA"
        assert op.table_name == "tbl_site_data"
        assert op.column_name == "Value"
        assert op.value == "Haifa"

    def test_raises_write_permission_error_for_parcel_001(self, mapper):
        with pytest.raises(WritePermissionError) as exc_info:
            mapper.map_input({"PARCEL_001": "P_001"})
        assert exc_info.value.field_id == "PARCEL_001"

    def test_raises_on_first_protected_field_in_mixed_input(self, mapper):
        """WritePermissionError raised even if other fields are valid."""
        with pytest.raises(WritePermissionError):
            mapper.map_input({"SITE_001": "Tel Aviv", "PARCEL_001": "X"})

    def test_skips_unknown_field_id_with_warning(self, mapper, caplog):
        import logging
        with caplog.at_level(logging.WARNING):
            ops = mapper.map_input({"UNKNOWN_999": "value"})
        assert ops == []
        assert "UNKNOWN_999" in caplog.text

    def test_maps_multiple_fields(self, mapper):
        ops = mapper.map_input({
            "SITE_001": "Tel Aviv",
            "SITE_002": "Florentin",
            "PARCEL_002": "6660",
        })
        assert len(ops) == 3
        field_ids = {op.field_id for op in ops}
        assert field_ids == {"SITE_001", "SITE_002", "PARCEL_002"}

    def test_attaches_source_id(self, mapper):
        ops = mapper.map_input({"SITE_001": "Haifa"}, source_id="SRC_001")
        assert ops[0].source_id == "SRC_001"

    def test_empty_input_returns_empty_list(self, mapper):
        assert mapper.map_input({}) == []

    def test_parcel_002_maps_to_correct_column(self, mapper):
        ops = mapper.map_input({"PARCEL_002": "6660"})
        op = ops[0]
        assert op.sheet_name == "PARCELS"
        assert op.column_name == "Block_Number"

    def test_build_001_maps_to_buildings_sheet(self, mapper):
        ops = mapper.map_input({"BUILD_001": 5})
        op = ops[0]
        assert op.sheet_name == "BUILDINGS"
        assert op.column_name == "Floors_Above_Ground"


# ── group_by_sheet ─────────────────────────────────────────────────────────────

class TestGroupBySheet:
    def test_groups_correctly(self, mapper):
        ops = mapper.map_input({
            "SITE_001": "Tel Aviv",
            "PARCEL_002": "6660",
            "PARCEL_003": "1",
            "PLAN_001": "TMA2050",
        })
        grouped = mapper.group_by_sheet(ops)
        assert set(grouped.keys()) == {"SITE_DATA", "PARCELS", "PLANS"}
        assert len(grouped["PARCELS"]) == 2
        assert len(grouped["SITE_DATA"]) == 1
        assert len(grouped["PLANS"]) == 1

    def test_empty_ops_returns_empty_dict(self, mapper):
        assert mapper.group_by_sheet([]) == {}

    def test_single_sheet_all_in_one_group(self, mapper):
        ops = mapper.map_input({"SITE_001": "A", "SITE_002": "B", "SITE_003": "C"})
        grouped = mapper.group_by_sheet(ops)
        assert list(grouped.keys()) == ["SITE_DATA"]
        assert len(grouped["SITE_DATA"]) == 3
