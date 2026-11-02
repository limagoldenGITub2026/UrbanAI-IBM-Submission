"""
app.py — Streamlit UI for the Urban Renewal Excel Data Interface

Run from the 2026.11.02/ directory:
    streamlit run excel_interface/app.py

Features
--------
- Load any .xlsx workbook (or use the bundled template)
- Browse existing data in each sheet
- Write new field values (Field_ID → value) with a source name
- View validation errors, conflicts, and missing-fields reports
- Download the updated workbook
"""

from __future__ import annotations

import io
import shutil
import tempfile
from pathlib import Path

import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Urban Renewal — Data Interface",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH = Path(__file__).parent.parent / "urban_renewal_template.xlsx"

# ── Helpers ────────────────────────────────────────────────────────────────────

def _get_iface(path: Path):
    """Load an ExcelDataInterface from *path*, cached by path + mtime."""
    from excel_interface import ExcelDataInterface
    iface = ExcelDataInterface()
    iface.load(path)
    return iface


def _bytes_to_tmp(uploaded) -> Path:
    """Write an UploadedFile to a temp file and return its path."""
    suffix = Path(uploaded.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.read())
    tmp.flush()
    return Path(tmp.name)


def _workbook_to_bytes(iface) -> bytes:
    """Save the current workbook to an in-memory buffer and return bytes."""
    buf = io.BytesIO()
    iface._writer.workbook.save(buf)
    return buf.getvalue()


# ── Session state initialisation ──────────────────────────────────────────────

if "iface" not in st.session_state:
    st.session_state.iface = None
if "wb_path" not in st.session_state:
    st.session_state.wb_path = None
if "write_log" not in st.session_state:
    st.session_state.write_log = []  # list of (field_id, value, source, status)


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("🏗️ Urban Renewal\nData Interface")
    st.divider()

    st.subheader("📂 Load Workbook")
    use_template = st.checkbox("Use bundled template", value=True)

    if use_template:
        if TEMPLATE_PATH.exists():
            if st.button("Load Template", type="primary"):
                # Copy to a temp file so we can write to it
                tmp = Path(tempfile.mktemp(suffix=".xlsx"))
                shutil.copy(TEMPLATE_PATH, tmp)
                st.session_state.wb_path = tmp
                st.session_state.iface = _get_iface(tmp)
                st.session_state.write_log = []
                st.success("Template loaded.")
        else:
            st.error(f"Template not found:\n{TEMPLATE_PATH}")
    else:
        uploaded = st.file_uploader("Upload .xlsx", type=["xlsx"])
        if uploaded and st.button("Load Uploaded File", type="primary"):
            tmp_path = _bytes_to_tmp(uploaded)
            st.session_state.wb_path = tmp_path
            st.session_state.iface = _get_iface(tmp_path)
            st.session_state.write_log = []
            st.success(f"Loaded: {uploaded.name}")

    if st.session_state.iface:
        st.divider()
        st.caption("✅ Workbook loaded")
        sheets = st.session_state.iface._reader.sheet_names()
        st.caption(f"{len(sheets)} sheets · "
                   f"{len(st.session_state.iface._reader.field_locations)} fields")


# ── Main area ──────────────────────────────────────────────────────────────────

if not st.session_state.iface:
    st.info("👈 Load a workbook from the sidebar to get started.")
    st.stop()

iface = st.session_state.iface

tab_browse, tab_write, tab_report, tab_download = st.tabs([
    "📋 Browse Data",
    "✏️ Write Values",
    "📊 Reports",
    "⬇️ Download",
])


# ── Tab 1: Browse ──────────────────────────────────────────────────────────────

with tab_browse:
    st.header("Browse Workbook Sheets")

    reader = iface._reader
    sheet_names = [s for s in reader.sheet_names() if s != "README"]
    selected_sheet = st.selectbox("Select sheet", sheet_names)

    if selected_sheet:
        try:
            rows = reader.read_sheet(selected_sheet)
        except Exception as e:
            st.error(f"Error reading sheet: {e}")
            rows = []

        if rows:
            st.dataframe(rows, use_container_width=True)
            st.caption(f"{len(rows)} row(s)")
        else:
            st.info(f"Sheet **{selected_sheet}** has no data rows yet.")

    st.divider()
    with st.expander("📖 DATA_MAPPING — Field Registry"):
        try:
            dm_rows = reader.read_sheet("DATA_MAPPING")
            st.dataframe(dm_rows, use_container_width=True)
        except Exception:
            st.warning("DATA_MAPPING sheet not found.")


# ── Tab 2: Write Values ────────────────────────────────────────────────────────

with tab_write:
    st.header("Write Field Values")

    field_locations = iface._reader.field_locations
    # Build options: only writable fields
    writable = {
        fid: loc for fid, loc in field_locations.items()
        if loc.write_permission
    }
    if not writable:
        st.warning("No writable fields found in DATA_MAPPING.")
    else:
        col1, col2 = st.columns([1, 2])

        with col1:
            st.subheader("Single Field")
            field_options = [
                f"{fid} — {loc.sheet_name} / {loc.column_name}"
                for fid, loc in writable.items()
            ]
            selected_field_str = st.selectbox("Field_ID", field_options)
            selected_fid = selected_field_str.split(" — ")[0] if selected_field_str else None
            loc = writable.get(selected_fid) if selected_fid else None

            if loc:
                st.caption(
                    f"Sheet: **{loc.sheet_name}** · Table: `{loc.table_name}` · "
                    f"Column: `{loc.column_name}` · Type: `{loc.data_type}`"
                    + (" · **Required**" if loc.required else "")
                )

            value_input = st.text_input(
                "Value",
                placeholder="Enter value...",
                key="single_value",
            )
            source_input = st.text_input(
                "Source name",
                value="XPlan",
                key="single_source",
            )

            if st.button("Write Field", type="primary", disabled=not selected_fid):
                if not value_input.strip():
                    st.warning("Value cannot be empty.")
                else:
                    # Type-coerce for non-text fields
                    val = value_input
                    if loc and loc.data_type in ("integer",):
                        try:
                            val = int(value_input)
                        except ValueError:
                            pass
                    elif loc and loc.data_type == "number":
                        try:
                            val = float(value_input)
                        except ValueError:
                            pass

                    try:
                        iface.write({selected_fid: val}, source_name=source_input or "Manual")
                        st.success(f"✅ Written: `{selected_fid}` = `{val}`")
                        st.session_state.write_log.append(
                            (selected_fid, val, source_input, "✅ OK")
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")
                        st.session_state.write_log.append(
                            (selected_fid, val, source_input, f"❌ {e}")
                        )

        with col2:
            st.subheader("Bulk Write (JSON)")
            st.caption(
                'Enter a JSON dict keyed by Field_ID, e.g.  '
                '`{"SITE_001": "Tel Aviv", "PARCEL_002": "6660"}`'
            )
            bulk_json = st.text_area(
                "Data dict (JSON)",
                height=140,
                placeholder='{"SITE_001": "Tel Aviv", "PARCEL_002": "6660"}',
                key="bulk_json",
            )
            bulk_source = st.text_input("Source name", value="XPlan", key="bulk_source")

            if st.button("Write Bulk", type="primary"):
                if not bulk_json.strip():
                    st.warning("JSON input is empty.")
                else:
                    import json
                    try:
                        data_dict = json.loads(bulk_json)
                        if not isinstance(data_dict, dict):
                            st.error("Input must be a JSON object (dict).")
                        else:
                            iface.write(data_dict, source_name=bulk_source or "Manual")
                            st.success(f"✅ Bulk write: {len(data_dict)} field(s) processed.")
                            for fid, val in data_dict.items():
                                st.session_state.write_log.append(
                                    (fid, val, bulk_source, "✅ OK")
                                )
                    except json.JSONDecodeError as e:
                        st.error(f"Invalid JSON: {e}")

        st.divider()
        if st.session_state.write_log:
            st.subheader("Write Log (this session)")
            import pandas as pd
            log_df = pd.DataFrame(
                st.session_state.write_log,
                columns=["Field_ID", "Value", "Source", "Status"],
            )
            st.dataframe(log_df, use_container_width=True)
            if st.button("Clear Log"):
                st.session_state.write_log = []
                st.rerun()


# ── Tab 3: Reports ─────────────────────────────────────────────────────────────

with tab_report:
    st.header("Validation & Missing-Fields Reports")

    if st.button("🔄 Refresh Report"):
        st.rerun()

    try:
        rep = iface.report()
    except Exception as e:
        st.error(f"Could not generate report: {e}")
        st.stop()

    col_r, col_m = st.columns(2)

    with col_r:
        st.subheader("🔴 Errors")
        if rep["errors"]:
            import pandas as pd
            st.dataframe(pd.DataFrame(rep["errors"]), use_container_width=True)
        else:
            st.success("No validation errors.")

        st.subheader("⚠️ Conflicts")
        if rep["conflicts"]:
            import pandas as pd
            st.dataframe(pd.DataFrame(rep["conflicts"]), use_container_width=True)
        else:
            st.success("No conflicts detected.")

        st.subheader("🚫 Unauthorized Writes")
        if rep["unauthorized_writes"]:
            for fid in rep["unauthorized_writes"]:
                st.warning(f"`{fid}` — write_permission = no")
        else:
            st.success("No unauthorized write attempts.")

    with col_m:
        st.subheader("📌 Required Fields Missing")
        if rep["required_missing"]:
            for fid in rep["required_missing"]:
                loc = iface._reader.field_locations.get(fid)
                label = loc.sheet_name + " / " + loc.column_name if loc else ""
                st.error(f"`{fid}` — {label}")
        else:
            st.success("All required fields are filled.")

        st.subheader("💡 Optional Fields Missing")
        if rep["optional_missing"]:
            for fid in rep["optional_missing"]:
                loc = iface._reader.field_locations.get(fid)
                label = loc.sheet_name + " / " + loc.column_name if loc else ""
                st.info(f"`{fid}` — {label}")
        else:
            st.success("All optional fields are filled.")


# ── Tab 4: Download ────────────────────────────────────────────────────────────

with tab_download:
    st.header("Save & Download Updated Workbook")

    st.info(
        "Clicking **Save & Download** saves a backup of the original file next to it "
        "and prepares the updated workbook for download."
    )

    col_d1, col_d2 = st.columns(2)

    with col_d1:
        output_name = st.text_input(
            "Output filename",
            value="urban_renewal_updated.xlsx",
            key="output_name",
        )

    with col_d2:
        st.write("")  # vertical alignment spacer
        st.write("")

    if st.button("💾 Save & Download", type="primary"):
        try:
            wb_bytes = _workbook_to_bytes(iface)
            st.download_button(
                label="⬇️ Download updated workbook",
                data=wb_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("Ready to download.")
        except Exception as e:
            st.error(f"Error saving workbook: {e}")

    st.divider()
    st.subheader("📋 SOURCES_LOG Preview")
    try:
        src_rows = iface._reader.read_sheet("SOURCES_LOG")
        if src_rows:
            import pandas as pd
            st.dataframe(pd.DataFrame(src_rows), use_container_width=True)
        else:
            st.info("No sources logged yet.")
    except Exception:
        st.info("SOURCES_LOG not available.")
