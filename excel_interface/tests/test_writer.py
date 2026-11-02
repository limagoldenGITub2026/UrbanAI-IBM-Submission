"""
test_writer.py — Unit tests for ExcelWriter

Run from the 2026.11.02/ directory:
    python -m pytest excel_interface/tests/test_writer.py -v
"""

from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
import pytest

from excel_interface.reader import ExcelReader
from excel_interface.writer import ExcelWriter

TEMPLATE = Path(__file__).parent.parent.parent / "urban_renewal_template.xlsx"


@pytest.fixture
def setup(tmp_path):
    """Return (reader, writer, workbook, tmp_path) loaded from a fresh template copy."""
    dest = tmp_path / "test_workbook.xlsx"
    shutil.copy(TEMPLATE, dest)
    wb = openpyxl.load_workbook(str(dest))
    reader = ExcelReader(dest)
    reader.load_data_mapping()
    writer = ExcelWriter(wb, reader, dest)
    return reader, writer, wb, dest


# ── write_field (SITE_DATA vertical) ──────────────────────────────────────────

class TestWriteField:
    def test_writes_value_to_site_001(self, setup):
        reader, writer, wb, dest = setup
        writer.write_field("SITE_001", "Jerusalem")
        ws = wb["SITE_DATA"]
        # Find row for SITE_001 and read Value column (C = index 3)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == "SITE_001":
                assert row[2] == "Jerusalem"
                return
        pytest.fail("SITE_001 row not found")

    def test_updates_status_to_collected(self, setup):
        reader, writer, wb, dest = setup
        writer.write_field("SITE_002", "Florentin")
        ws = wb["SITE_DATA"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == "SITE_002":
                assert row[3] == "Collected"
                return
        pytest.fail("SITE_002 row not found")

    def test_raises_for_unknown_field_id(self, setup):
        _, writer, _, _ = setup
        with pytest.raises(ValueError, match="SITE_999"):
            writer.write_field("SITE_999", "X")

    def test_write_persists_after_save(self, setup, tmp_path):
        reader, writer, wb, dest = setup
        writer.write_field("SITE_001", "Haifa")
        out_path = tmp_path / "out.xlsx"
        writer.save(out_path)

        # Reload and verify
        wb2 = openpyxl.load_workbook(str(out_path), data_only=True)
        ws2 = wb2["SITE_DATA"]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if row[0] == "SITE_001":
                assert row[2] == "Haifa"
                return
        pytest.fail("SITE_001 not found after save/reload")


# ── append_row ─────────────────────────────────────────────────────────────────

class TestAppendRow:
    def test_appends_parcel_row(self, setup):
        reader, writer, wb, dest = setup
        writer.append_row("PARCELS", {
            "Parcel_ID": "P_6660_001",
            "Block_Number": "6660",
            "Parcel_Number": "1",
            "Area_sqm": 450.0,
            "Source_ID": "SRC_001",
        })
        ws = wb["PARCELS"]
        # Header is row 1; data should now be in row 2
        assert ws.max_row == 2
        assert ws["A2"].value == "P_6660_001"
        assert ws["B2"].value == "6660"
        assert ws["D2"].value == 450.0

    def test_append_two_rows_increases_count(self, setup):
        reader, writer, wb, dest = setup
        for i in range(2):
            writer.append_row("PARCELS", {"Parcel_ID": f"P_00{i}", "Block_Number": str(6660 + i)})
        assert wb["PARCELS"].max_row == 3  # header + 2 data rows

    def test_missing_columns_filled_with_none(self, setup):
        _, writer, wb, _ = setup
        writer.append_row("PARCELS", {"Block_Number": "1111"})
        ws = wb["PARCELS"]
        assert ws["A2"].value is None   # Parcel_ID not provided
        assert ws["B2"].value == "1111"


# ── update_row ─────────────────────────────────────────────────────────────────

class TestUpdateRow:
    def _add_parcel(self, writer):
        writer.append_row("PARCELS", {
            "Parcel_ID": "P_001",
            "Block_Number": "6660",
            "Parcel_Number": "1",
            "Area_sqm": 300.0,
            "Source_ID": "SRC_001",
        })

    def test_updates_existing_row(self, setup):
        reader, writer, wb, _ = setup
        self._add_parcel(writer)
        result = writer.update_row("PARCELS", "Parcel_ID", "P_001", {"Area_sqm": 450.0})
        assert result is True
        ws = wb["PARCELS"]
        assert ws["D2"].value == 450.0   # Area_sqm column

    def test_does_not_change_id_column(self, setup):
        reader, writer, wb, _ = setup
        self._add_parcel(writer)
        writer.update_row("PARCELS", "Parcel_ID", "P_001", {"Parcel_ID": "HACKED", "Area_sqm": 500})
        ws = wb["PARCELS"]
        assert ws["A2"].value == "P_001"  # ID must not be changed

    def test_returns_false_when_id_not_found(self, setup):
        reader, writer, wb, _ = setup
        self._add_parcel(writer)
        result = writer.update_row("PARCELS", "Parcel_ID", "P_NOTEXIST", {"Area_sqm": 999})
        assert result is False

    def test_row_count_unchanged_after_update(self, setup):
        reader, writer, wb, _ = setup
        self._add_parcel(writer)
        writer.update_row("PARCELS", "Parcel_ID", "P_001", {"Area_sqm": 999})
        assert wb["PARCELS"].max_row == 2  # header + 1 data row


# ── save ───────────────────────────────────────────────────────────────────────

class TestSave:
    def test_creates_updated_file(self, setup, tmp_path):
        reader, writer, wb, dest = setup
        out = writer.save(tmp_path / "output.xlsx")
        assert out.exists()

    def test_default_output_path_has_updated_suffix(self, setup):
        reader, writer, wb, dest = setup
        out = writer.save()
        assert "_updated" in out.name
        assert out.exists()
        out.unlink(missing_ok=True)

    def test_creates_backup_of_original(self, setup):
        reader, writer, wb, dest = setup
        writer.save(dest.parent / "output.xlsx")
        from datetime import date
        stamp = date.today().strftime("%Y%m%d")
        backup = dest.with_name(f"{dest.stem}_backup_{stamp}{dest.suffix}")
        assert backup.exists()

    def test_backup_created_only_once(self, setup, tmp_path):
        reader, writer, wb, dest = setup
        writer.save(tmp_path / "out1.xlsx")
        writer.save(tmp_path / "out2.xlsx")
        from datetime import date
        stamp = date.today().strftime("%Y%m%d")
        backup = dest.with_name(f"{dest.stem}_backup_{stamp}{dest.suffix}")
        # Backup should exist but there should be only one backup file
        assert backup.exists()
