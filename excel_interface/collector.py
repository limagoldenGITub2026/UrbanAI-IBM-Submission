"""
collector.py — Automatic Data Collector UI (Streamlit)

Run from the 2026.11.02/ directory:
    streamlit run excel_interface/collector.py

Three search modes
------------------
Tab A — כתובת : city + street address  (+ optional block/parcel hints)
Tab B — גוש/חלקות : city + block + list of parcels (one run per parcel)
Tab C — רחובות : city + multi-street list (one run per street)

Sidebar shows a history of every run in this session with quick-jump links.
"""

from __future__ import annotations

import io
import sys
from pathlib import Path
from typing import List

import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="איסוף נתונים אוטומטי — התחדשות עירונית",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── RTL + shared styling ───────────────────────────────────────────────────────
st.markdown("""
<style>
    body, .stApp { direction: rtl; }
    .stTextInput input, .stTextArea textarea { text-align: right; direction: rtl; }
    h1, h2, h3, h4, p, li { direction: rtl; text-align: right; }
    .status-ok       { color: #16a34a; font-weight: bold; }
    .status-partial  { color: #d97706; font-weight: bold; }
    .status-error    { color: #dc2626; font-weight: bold; }
    .status-manual   { color: #7c3aed; font-weight: bold; }
    .field-tag {
        display: inline-block;
        background: #e0f2fe;
        color: #0369a1;
        border-radius: 4px;
        padding: 1px 6px;
        margin: 2px;
        font-size: 0.82em;
        font-family: monospace;
    }
    .history-item {
        padding: 4px 0;
        border-bottom: 1px solid #e5e7eb;
        font-size: 0.88em;
    }
</style>
""", unsafe_allow_html=True)

# ── Session state ──────────────────────────────────────────────────────────────
if "history" not in st.session_state:
    st.session_state.history = []          # list of CollectionSummary (or list thereof)
if "active_result" not in st.session_state:
    st.session_state.active_result = None  # currently displayed summary/list


# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════

STATUS_LABELS = {
    "ok":              ("✅", "הצליח"),
    "partial":         ("⚠️", "חלקי"),
    "unavailable":     ("❌", "לא זמין"),
    "manual_required": ("🔒", "ידני"),
}


def _render_single_summary(summary) -> None:
    """Render KPIs, connector table, flat fields, manual list and download for one summary."""
    st.header(f"📊 תוצאות — {summary.city}  |  {summary.item_label or summary.address}")

    # KPI row
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("✅ מקורות הצליחו", summary.ok_count)
    kpi2.metric("⚠️ חלקי", summary.partial_count)
    kpi3.metric("❌ לא זמין", summary.failed_count)
    kpi4.metric("📋 שדות שנאספו", summary.total_fields_collected)

    st.divider()

    # Per-connector status
    st.subheader("📡 סטטוס לפי מקור")
    for result in summary.results:
        emoji, label = STATUS_LABELS.get(result.status.value, ("❓", result.status.value))
        with st.expander(
            f"{emoji} **{result.source_name}** — {label}",
            expanded=(result.status.value in ("ok", "partial")),
        ):
            st.write(result.notes)
            if result.fields_collected:
                tags = " ".join(
                    f'<span class="field-tag">{f}</span>'
                    for f in result.fields_collected
                )
                st.markdown(f"**שדות שנאספו:** {tags}", unsafe_allow_html=True)
            if result.field_data:
                st.json(result.field_data, expanded=False)

    st.divider()

    # Flat field table
    all_field_data: dict = {}
    for r in summary.results:
        all_field_data.update(r.field_data)

    if all_field_data:
        st.subheader("📋 כל השדות שנאספו")
        import pandas as pd
        rows = [
            {
                "Field_ID": k,
                "ערך": v,
                "מקור": next(
                    (r.source_name for r in summary.results if k in r.field_data), "—"
                ),
            }
            for k, v in all_field_data.items()
        ]
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
        st.divider()

    # Manual required
    st.subheader("🔒 מקורות הדורשים הגדרה ידנית")
    st.caption("המקורות הבאים דורשים אימות או תשלום — יש להשלים ידנית בקובץ Excel.")
    for item in summary.manual_required:
        fields_str = " ".join(
            f'<span class="field-tag">{f}</span>' for f in item["fields"]
        )
        st.markdown(
            f"**{item['name']}** — {item['reason']}<br>"
            f"שדות: {fields_str}<br>"
            f"<a href='{item['url']}' target='_blank'>🔗 פתח מקור</a>",
            unsafe_allow_html=True,
        )
        st.write("")

    st.divider()

    # Download
    st.subheader("⬇️ הורדת הקובץ")
    try:
        buf = io.BytesIO()
        from openpyxl import load_workbook as _lw
        wb = _lw(str(summary.output_path))
        wb.save(buf)
        buf.seek(0)
        st.download_button(
            label=f"⬇️ הורד: {summary.output_path.name}",
            data=buf.getvalue(),
            file_name=summary.output_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption(f"הקובץ נשמר גם ב: `{summary.output_path}`")
    except Exception as exc:
        st.error(f"לא ניתן לטעון את הקובץ להורדה: {exc}")


def _render_multi_summary(summaries: list, label: str) -> None:
    """Render a tabbed view for a list of summaries (cadastral / streets mode)."""
    st.header(f"📊 תוצאות — {label}")

    if not summaries:
        st.warning("לא הוחזרו תוצאות.")
        return

    # Aggregate KPIs across all items
    total_ok      = sum(s.ok_count for s in summaries)
    total_partial = sum(s.partial_count for s in summaries)
    total_failed  = sum(s.failed_count for s in summaries)
    total_fields  = sum(s.total_fields_collected for s in summaries)

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("✅ מקורות הצליחו (סה״כ)", total_ok)
    kpi2.metric("⚠️ חלקי (סה״כ)", total_partial)
    kpi3.metric("❌ לא זמין (סה״כ)", total_failed)
    kpi4.metric("📋 שדות שנאספו (סה״כ)", total_fields)

    st.divider()

    # One sub-tab per item
    tab_labels = [s.item_label or s.address for s in summaries]
    tabs = st.tabs(tab_labels)
    for tab, summary in zip(tabs, summaries):
        with tab:
            _render_single_summary(summary)


def _make_progress_callback(bar, container, msgs):
    def callback(step: int, total: int, message: str) -> None:
        pct = int(step / max(total, 1) * 100)
        bar.progress(pct, text=message)
        msgs.append(f"✔ {message}")
        container.markdown("\n".join(msgs))
    return callback


# ══════════════════════════════════════════════════════════════════════════════
# Sidebar — search history
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.header("🕑 היסטוריית חיפושים")
    if not st.session_state.history:
        st.caption("אין חיפושים בסשן הנוכחי.")
    else:
        for idx, entry in enumerate(reversed(st.session_state.history)):
            # entry is either a single summary or a list of summaries
            if isinstance(entry, list):
                first = entry[0]
                mode_icon = "🗺️" if first.search_mode == "cadastral" else "🛣️"
                lbl = f"{mode_icon} {first.city} ({len(entry)} פריטים)"
            else:
                lbl = f"📍 {entry.city} — {entry.address}"

            if st.button(lbl, key=f"hist_{idx}", use_container_width=True):
                st.session_state.active_result = entry

    if st.session_state.history:
        st.divider()
        if st.button("🗑️ נקה היסטוריה", use_container_width=True):
            st.session_state.history = []
            st.session_state.active_result = None
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# Main — header
# ══════════════════════════════════════════════════════════════════════════════

st.title("🏗️ איסוף נתונים אוטומטי — מתחם התחדשות עירונית")
st.caption(
    "בחר מצב חיפוש: כתובת, גוש/חלקות, או רשימת רחובות. "
    "הסוכן יאסוף נתוני חלקה, תכנון ודמוגרפיה ממקורות ממשלתיים פתוחים."
)
st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# Three search tabs
# ══════════════════════════════════════════════════════════════════════════════

tab_a, tab_b, tab_c = st.tabs(["📍 כתובת", "🗺️ גוש / חלקות", "🛣️ רשימת רחובות"])


# ─────────────────────────────────────────────────────────────────────────────
# TAB A — Address (original mode, unchanged logic)
# ─────────────────────────────────────────────────────────────────────────────

with tab_a:
    with st.form("form_address", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            city_a = st.text_input(
                "🏙️ שם עיר *",
                placeholder="לדוגמה: תל אביב",
                help="שם העיר בעברית",
                key="city_a",
            )
            address_a = st.text_input(
                "📍 כתובת (רחוב ומספר) *",
                placeholder="לדוגמה: רוטשילד 1",
                help="שם הרחוב ומספר הבית בעברית",
                key="address_a",
            )

        with col2:
            block_a = st.text_input(
                "גוש (אופציונלי)",
                placeholder="לדוגמה: 6660",
                help="מספר גוש קדסטרלי — משפר את דיוק החיפוש",
                key="block_a",
            )
            parcel_a = st.text_input(
                "חלקה (אופציונלי)",
                placeholder="לדוגמה: 1",
                help="מספר חלקה — משפר את דיוק החיפוש",
                key="parcel_a",
            )

        submitted_a = st.form_submit_button(
            "▶ אסוף נתונים",
            type="primary",
            use_container_width=True,
        )

    if submitted_a:
        if not city_a.strip():
            st.error("יש להזין שם עיר.")
            st.stop()
        if not address_a.strip():
            st.error("יש להזין כתובת.")
            st.stop()

        bar = st.progress(0, text="מתחיל בגביית נתונים…")
        container = st.empty()
        msgs: list[str] = []

        try:
            from excel_interface.collection_engine import CollectionEngine
            engine = CollectionEngine()
            summary = engine.run(
                city=city_a.strip(),
                address=address_a.strip(),
                block=block_a.strip() or None,
                parcel=parcel_a.strip() or None,
                progress_callback=_make_progress_callback(bar, container, msgs),
            )
            bar.progress(100, text="✅ האיסוף הושלם!")
            st.session_state.active_result = summary
            st.session_state.history.append(summary)
        except FileNotFoundError as exc:
            st.error(f"שגיאה: {exc}")
            st.stop()
        except Exception as exc:
            st.error(f"שגיאה לא צפויה: {exc}")
            st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# TAB B — Cadastral (גוש + חלקות)
# ─────────────────────────────────────────────────────────────────────────────

with tab_b:
    st.markdown(
        "הזן עיר, גוש ורשימת חלקות. הסוכן יריץ איסוף נפרד לכל חלקה וישמור קובץ Excel נפרד."
    )

    with st.form("form_cadastral", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            city_b = st.text_input(
                "🏙️ שם עיר *",
                placeholder="לדוגמה: תל אביב",
                key="city_b",
            )
            neighborhood_b = st.text_input(
                "🏘️ שכונה (אופציונלי)",
                placeholder="לדוגמה: נווה צדק",
                help="מסייע לגיאוקודינג — לא חובה",
                key="neighborhood_b",
            )

        with col2:
            block_b = st.text_input(
                "🗺️ גוש *",
                placeholder="לדוגמה: 6660",
                key="block_b",
            )
            parcels_raw = st.text_area(
                "📋 חלקות * (מספר בשורה, או מופרדות בפסיקים)",
                placeholder="לדוגמה:\n1\n2\n5\n\nאו: 1, 2, 5",
                height=130,
                help="הזן חלקה אחת בכל שורה, או מופרדות בפסיקים",
                key="parcels_raw",
            )

        submitted_b = st.form_submit_button(
            "▶ אסוף נתונים לכל החלקות",
            type="primary",
            use_container_width=True,
        )

    if submitted_b:
        # Validate
        if not city_b.strip():
            st.error("יש להזין שם עיר.")
            st.stop()
        if not block_b.strip():
            st.error("יש להזין מספר גוש.")
            st.stop()
        if not parcels_raw.strip():
            st.error("יש להזין לפחות חלקה אחת.")
            st.stop()

        # Parse parcels — split on newlines and/or commas, strip whitespace
        raw_items = parcels_raw.replace(",", "\n").splitlines()
        parcels_b = [p.strip() for p in raw_items if p.strip()]
        if not parcels_b:
            st.error("לא נמצאו מספרי חלקות תקינים.")
            st.stop()

        st.info(f"מריץ איסוף עבור **{len(parcels_b)} חלקות** בגוש **{block_b.strip()}** — {city_b.strip()}")

        bar = st.progress(0, text="מתחיל…")
        container = st.empty()
        msgs: list[str] = []

        try:
            from excel_interface.collection_engine import CollectionEngine
            engine = CollectionEngine()
            summaries = engine.run_cadastral(
                city=city_b.strip(),
                block=block_b.strip(),
                parcels=parcels_b,
                neighborhood=neighborhood_b.strip(),
                progress_callback=_make_progress_callback(bar, container, msgs),
            )
            bar.progress(100, text="✅ האיסוף הושלם!")
            st.session_state.active_result = summaries
            st.session_state.history.append(summaries)
        except FileNotFoundError as exc:
            st.error(f"שגיאה: {exc}")
            st.stop()
        except Exception as exc:
            st.error(f"שגיאה לא צפויה: {exc}")
            st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# TAB C — Multi-street (רשימת רחובות)
# ─────────────────────────────────────────────────────────────────────────────

with tab_c:
    st.markdown(
        "הזן עיר ורשימת רחובות. הסוכן יריץ איסוף נפרד לכל רחוב וישמור קובץ Excel נפרד."
    )

    with st.form("form_streets", clear_on_submit=False):
        city_c = st.text_input(
            "🏙️ שם עיר *",
            placeholder="לדוגמה: תל אביב",
            key="city_c",
        )
        streets_raw = st.text_area(
            "🛣️ רשימת רחובות * (שם רחוב בשורה — ניתן להוסיף טווח מספרים אחרי רווח)",
            placeholder=(
                "לדוגמה:\n"
                "דיזנגוף\n"
                "אבן גבירול 1-50\n"
                "בן יהודה\n"
                "רוטשילד 1-100"
            ),
            height=160,
            help=(
                "שורה אחת לכל רחוב. "
                "ניתן להוסיף טווח מספרים (לדוגמה: אבן גבירול 1-50) — "
                "המידע ישמש כרמז לגיאוקודינג."
            ),
            key="streets_raw",
        )

        submitted_c = st.form_submit_button(
            "▶ אסוף נתונים לכל הרחובות",
            type="primary",
            use_container_width=True,
        )

    if submitted_c:
        if not city_c.strip():
            st.error("יש להזין שם עיר.")
            st.stop()
        if not streets_raw.strip():
            st.error("יש להזין לפחות רחוב אחד.")
            st.stop()

        # Parse streets — each non-empty line is "street [numbers]"
        streets_b: list[dict] = []
        for line in streets_raw.splitlines():
            line = line.strip()
            if not line:
                continue
            # Split at last whitespace group — if last token looks like a range "N-M" or a number, treat as numbers
            parts = line.rsplit(None, 1)
            if len(parts) == 2 and (parts[1].isdigit() or "-" in parts[1]):
                streets_b.append({"street": parts[0].strip(), "numbers": parts[1].strip()})
            else:
                streets_b.append({"street": line, "numbers": ""})

        if not streets_b:
            st.error("לא נמצאו רחובות תקינים.")
            st.stop()

        st.info(f"מריץ איסוף עבור **{len(streets_b)} רחובות** ב{city_c.strip()}")

        bar = st.progress(0, text="מתחיל…")
        container = st.empty()
        msgs: list[str] = []

        try:
            from excel_interface.collection_engine import CollectionEngine
            engine = CollectionEngine()
            summaries = engine.run_streets(
                city=city_c.strip(),
                streets=streets_b,
                progress_callback=_make_progress_callback(bar, container, msgs),
            )
            bar.progress(100, text="✅ האיסוף הושלם!")
            st.session_state.active_result = summaries
            st.session_state.history.append(summaries)
        except FileNotFoundError as exc:
            st.error(f"שגיאה: {exc}")
            st.stop()
        except Exception as exc:
            st.error(f"שגיאה לא צפויה: {exc}")
            st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# Results area — rendered below the tabs
# ══════════════════════════════════════════════════════════════════════════════

active = st.session_state.active_result
if active is None:
    st.stop()

st.divider()

if isinstance(active, list):
    # Multi-item result (cadastral or streets)
    first = active[0]
    mode_label = (
        f"גוש {active[0].address.split()[1] if 'גוש' in active[0].address else ''} — {active[0].city}"
        if first.search_mode == "cadastral"
        else first.city
    )
    _render_multi_summary(active, mode_label)
else:
    _render_single_summary(active)

st.divider()
if st.button("🔄 התחל מחדש", use_container_width=False):
    st.session_state.active_result = None
    st.rerun()
