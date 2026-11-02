# Plan: Excel Data Interface – Urban Renewal AI (Python / openpyxl)

## Overview

Build a Python module (`excel_interface`) that reads and writes a **single Excel workbook**
with 8 sheets, serving as the structured data layer for an urban renewal AI agent.

**Input**: Python `dict` keyed by `Field_ID` (e.g. `{"SITE_001": "Tel Aviv", "SRC_001": "XPlan"}`)
**Output**: Updated `.xlsx` workbook + error/missing-fields reports
**Library**: `openpyxl` only — no external recalculation tools
**Rule**: All field access goes through `Field_ID` → `DATA_MAPPING`, never by cell position or color

The 7 separate template `.xlsx` files in `Excell/` are reference material only.
The Python module always works with a single unified workbook.

---

## Sub-Task 1 — Define Architecture & Interfaces

**Status**: [x] done

**Intent**
Establish the module structure, class interfaces, and data contracts before writing any logic.
This prevents rework across all later sub-tasks.

**Expected Outcomes**
- `excel_interface/` package folder with `__init__.py` and one file per component
- Each component has a documented class signature (no logic yet — stubs only)
- `DataStatus` and `FieldLocation` dataclasses defined and shared

**Todo List**
1. Create `2026.11.02/excel_interface/` package
2. Define `FieldLocation` dataclass: `field_id, sheet_name, table_name, column_name, data_type, required, write_permission`
3. Define `DataStatus` enum: `MISSING | COLLECTED | REQUIRES_VERIFICATION | VERIFIED | CONFLICTING | NOT_APPLICABLE`
4. Create stub classes in separate files: `reader.py`, `mapper.py`, `writer.py`, `validator.py`, `sources_logger.py`, `interface.py`
5. `interface.py` is the single public entry point — exposes `load()`, `write()`, `save()`, `report()`

**Relevant Context**
- `2026.11.02/Excell/INSTRUCTIONS.md.docx` — field identification hierarchy, dev stages
- `2026.11.02/Excell/Readme.xlsx` — all mandatory working rules
- `2026.11.02/analysis.md` — full field dictionary and class architecture

---

## Sub-Task 2 — Create Unified Excel Template

**Status**: [x] done

**Intent**
Generate a single `.xlsx` workbook with all 8 sheets, correct table names, and column headers
matching DATA_MAPPING. This becomes the canonical workbook the system reads and writes.

**Expected Outcomes**
- `urban_renewal_template.xlsx` created in `2026.11.02/`
- 8 sheets present: `README`, `SITE_DATA`, `PARCELS`, `PLANNING_DATA`, `EXISTING_BUILDINGS`, `CONSTRAINTS`, `SOURCES_LOG`, `DATA_MAPPING`
- Each sheet has its named table (`tbl_site_data`, `tbl_parcels`, `tbl_plans`, `tbl_buildings`, `tbl_constraints`, `tbl_sources`) with correct column headers
- `DATA_MAPPING` sheet populated with all 17 Field_IDs from the reference files
- `README` sheet contains operating rules summary

**Todo List**
1. Read all 7 template `.xlsx` files from `Excell/` to extract exact column headers and table names
2. Create `create_template.py` script using openpyxl
3. Add each sheet with its named Excel table (using `openpyxl.worksheet.table.Table`)
4. Populate `DATA_MAPPING` sheet with all Field_IDs
5. Run script → produce `urban_renewal_template.xlsx`
6. Verify with `office_read mode:outline` that all 8 sheets and tables are present

**Relevant Context**
- Source column headers: `2026.11.02/Excell/גיליון PARCELS.xlsx` etc.
- Full Field_ID list: `2026.11.02/Excell/DATA_MAPPING.xlsx`
- Note: CONSTRAINTS sheet currently uses `tbl_constraints` as table name in SOURCES_LOG too — use `tbl_constraints` for CONSTRAINTS, `tbl_sources` for SOURCES_LOG

---

## Sub-Task 3 — Develop ExcelReader

**Status**: [x] done

**Intent**
Read an existing workbook, load DATA_MAPPING into memory, and expose methods to retrieve
existing records and identify missing required fields.

**Expected Outcomes**
- `reader.py`: `ExcelReader` class functional
- `load_data_mapping()` returns `dict[field_id → FieldLocation]`
- `read_sheet(sheet_name)` returns list of row dicts keyed by column name
- `get_record_by_id(sheet, id_column, id_value)` returns single row dict or None
- `identify_missing_fields()` returns list of Field_IDs where required=yes and value is empty

**Todo List**
1. Implement `ExcelReader.__init__(path)` — opens workbook with `openpyxl.load_workbook()`
2. Implement `load_data_mapping()` — reads `DATA_MAPPING` sheet, builds `dict[str, FieldLocation]`
3. Implement `read_sheet(sheet_name)` — finds sheet, detects table headers, returns list of row dicts
4. Implement `get_record_by_id(sheet_name, id_column, id_value)` — linear search on the table rows
5. Implement `identify_missing_fields()` — cross-reference DATA_MAPPING required fields vs actual values
6. Unit test: load `urban_renewal_template.xlsx`, assert DATA_MAPPING has 17 entries, PARCELS has 0 data rows

**Relevant Context**
- Field_ID is the lookup key — never use cell coordinates
- `SITE_DATA` uses a vertical layout (Field_ID | Value rows) — different from other sheets which are horizontal tables
- Read DATA_MAPPING first, always — other reads depend on it

---

## Sub-Task 4 — Develop DataMapper

**Status**: [x] done

**Intent**
Translate an incoming `dict[Field_ID → value]` into a set of targeted write operations,
grouped by sheet and table, with write permission checks applied.

**Expected Outcomes**
- `mapper.py`: `DataMapper` class functional
- `map_input(data_dict)` returns `list[WriteOperation]` (each op has: field_id, sheet, table, column, value)
- Raises `PermissionError` for any Field_ID where `write_permission = no`
- Groups operations by sheet for efficient batched writes

**Todo List**
1. Define `WriteOperation` dataclass: `field_id, sheet_name, table_name, column_name, value, source_id`
2. Implement `DataMapper.__init__(field_locations: dict[str, FieldLocation])`
3. Implement `map_input(data_dict: dict) → list[WriteOperation]`
   - Skip unknown Field_IDs with a warning (do not raise)
   - Raise `PermissionError` if `write_permission = no`
4. Implement `group_by_sheet(ops) → dict[sheet_name → list[WriteOperation]]`
5. Unit test: input `{"PARCEL_001": "X"}` raises PermissionError; input `{"SITE_001": "Haifa"}` maps to correct sheet/column

**Relevant Context**
- `PARCEL_001` has `write_permission = no` (system-generated ID) — must be blocked
- All other IDs in DATA_MAPPING have `write_permission = yes`

---

## Sub-Task 5 — Develop ExcelWriter

**Status**: [x] done

**Intent**
Apply write operations to the workbook: update existing records by ID, append new rows,
and never overwrite data without an explicit ID-based lookup.

**Expected Outcomes**
- `writer.py`: `ExcelWriter` class functional
- `write_field(field_id, value)` updates a single SITE_DATA field (vertical layout)
- `append_row(sheet, table, row_data)` adds a new row — only after duplicate check
- `update_row(sheet, id_column, id_value, updates)` updates an existing row in-place
- `save(output_path)` writes `_updated.xlsx` and keeps `_backup_YYYYMMDD.xlsx` untouched

**Todo List**
1. Implement `ExcelWriter.__init__(workbook, reader: ExcelReader)`
2. Implement `write_field(field_id, value)` — for SITE_DATA vertical layout only
3. Implement `append_row(sheet_name, row_data: dict)` — adds row at end of table
4. Implement `update_row(sheet_name, id_column, id_value, updates: dict)` — finds row, updates only specified columns
5. Implement `save(output_path)` — saves to `_updated.xlsx`, backs up original to `_backup_YYYYMMDD.xlsx`
6. Unit test: write `{"SITE_001": "Jerusalem"}` → read back and assert value. Append parcel → assert row count +1.
   Append same parcel ID again → assert row count unchanged, record updated instead.

**Relevant Context**
- Duplicate check: before `append_row` on PARCELS/PLANNING_DATA/BUILDINGS/CONSTRAINTS,
  call `reader.get_record_by_id()` — if found, call `update_row()` instead
- Do NOT change existing IDs (Parcel_ID, Plan_ID etc.) even during update

---

## Sub-Task 6 — Develop DataValidator

**Status**: [x] done

**Intent**
Validate incoming values against their declared types, detect conflicts with existing data,
and produce structured error and missing-field reports.

**Expected Outcomes**
- `validator.py`: `DataValidator` class functional
- `validate_type(field_id, value)` returns `(bool, error_message)`
- `check_duplicate(sheet, id_column, id_value)` returns existing row or None
- `detect_conflict(field_id, new_value, existing_value)` returns bool
- `generate_error_report()` returns structured dict
- `generate_missing_fields_report()` returns list of required Field_IDs with empty values

**Todo List**
1. Implement type validators per DATA_MAPPING types: `text`, `integer`, `number`, `date`, `boolean`
2. Implement `validate_type(field_id, value)` — uses DATA_MAPPING type for field_id
3. Implement `check_duplicate(sheet_name, id_column, id_value)`
4. Implement `detect_conflict(field_id, new_value, existing_value)` — flags when values differ and both are non-empty
5. Implement `generate_error_report()` — returns `{errors: [...], conflicts: [...], unauthorized_writes: [...]}`
6. Implement `generate_missing_fields_report()` — returns `{required_missing: [...], optional_missing: [...]}`
7. Validation rules from INSTRUCTIONS: area > 0, floors >= 0, units >= 0, year 1800–2100

**Relevant Context**
- Conflicting data → set `DataStatus.CONFLICTING` and write `Conflict_Notes` in SOURCES_LOG
- Do not raise on conflict — log and continue

---

## Sub-Task 7 — Develop SourcesLogger

**Status**: [x] done

**Intent**
Maintain a full audit trail: every value written to the workbook is linked to a Source_ID in SOURCES_LOG.

**Expected Outcomes**
- `sources_logger.py`: `SourcesLogger` class functional
- `log_source(source_name, source_type, document_date, reliability)` → returns new `Source_ID`
- `get_or_create_source(source_name)` → returns existing Source_ID if already logged, else creates new
- `link_source_to_field(source_id, field_id, record_id)` — records which source provided which field value
- Source_IDs auto-generated as `SRC_NNN` (SRC_001, SRC_002...)

**Todo List**
1. Implement `SourcesLogger.__init__(workbook, reader: ExcelReader)` — loads existing SOURCES_LOG rows
2. Implement `get_next_source_id()` — reads max existing SRC_NNN, increments
3. Implement `log_source(...)` — appends to `tbl_sources`, returns new Source_ID
4. Implement `get_or_create_source(source_name)` — deduplicates by source name
5. Integrate with `ExcelWriter`: every `write_field()` and `append_row()` call passes a `source_id`
6. Unit test: log same source twice → assert only one row, same Source_ID returned both times

**Relevant Context**
- SOURCES_LOG columns (from template): Source_ID, Source_Name, Source_Type, Document_Date, Reliability, Conflict_Notes
- `SRC_001` appears as default placeholder in PARCELS template rows — handle gracefully

---

## Sub-Task 8 — Integration, Public Interface & Tests

**Status**: [x] done

**Intent**
Wire all components into a single public `ExcelDataInterface` class and run end-to-end tests
covering all 7 functional requirements.

**Expected Outcomes**
- `interface.py` exposes: `load(path)`, `write(data_dict, source_name)`, `save(output_path)`, `report()`
- All 7 test scenarios pass (see Todo List)
- `excel_interface/` importable as a Python package

**Todo List**
1. Implement `ExcelDataInterface.load(path)` — creates Reader, loads DATA_MAPPING, creates Mapper/Writer/Validator/Logger
2. Implement `ExcelDataInterface.write(data_dict, source_name)`:
   a. `SourcesLogger.get_or_create_source(source_name)` → source_id
   b. `DataMapper.map_input(data_dict)` → write_ops
   c. For each op: `DataValidator.validate_type()`, then `DataValidator.check_duplicate()`
   d. Write via `ExcelWriter` (append or update)
   e. Log source link
3. Implement `ExcelDataInterface.save(output_path)` → delegates to `ExcelWriter.save()`
4. Implement `ExcelDataInterface.report()` → merges error + missing-fields reports
5. Test 1 — Read: load template, assert 8 sheets, DATA_MAPPING has 17 entries
6. Test 2 — Write SITE_DATA: write `{"SITE_001": "Haifa"}`, save, reload, assert value
7. Test 3 — Append parcel: write parcel fields, assert new row in PARCELS
8. Test 4 — Update existing: write same parcel again with new area, assert row count unchanged, value updated
9. Test 5 — Duplicate prevention: write parcel with existing Parcel_ID → assert update, not append
10. Test 6 — Permission check: write `{"PARCEL_001": "X"}` → assert PermissionError logged, no write
11. Test 7 — Missing fields report: on empty template, `report()` lists all required Field_IDs as missing

**Relevant Context**
- Entry point pattern: `interface = ExcelDataInterface(); interface.load("urban_renewal_template.xlsx"); interface.write({...}, "XPlan"); interface.save("output.xlsx")`
- All tests operate on a copy of `urban_renewal_template.xlsx` — never on the original

---

## File Layout (after all sub-tasks complete)

```
2026.11.02/
├── analysis.md                        ← this analysis
├── excel-interface-plan.md            ← this plan
├── urban_renewal_template.xlsx        ← unified 8-sheet template (Sub-Task 2)
└── excel_interface/
    ├── __init__.py
    ├── interface.py                   ← public entry point
    ├── reader.py
    ├── mapper.py
    ├── writer.py
    ├── validator.py
    ├── sources_logger.py
    ├── create_template.py             ← one-time template generator
    └── tests/
        ├── test_reader.py
        ├── test_writer.py
        ├── test_validator.py
        └── test_integration.py
```
