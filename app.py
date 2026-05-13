import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Auto Roadmap Dashboard",
    page_icon="🗺️",
    layout="wide",
)

# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .stTabs [data-baseweb="tab"] { font-size: 15px; font-weight: 600; }
    .stat-card {
        background: #f8f9fb;
        border-radius: 10px;
        padding: 16px 20px;
        border-left: 4px solid #4f6ef7;
        margin-bottom: 8px;
    }
    .stat-label { color: #666; font-size: 13px; margin-bottom: 2px; }
    .stat-value { font-size: 22px; font-weight: 700; color: #1a1a2e; }
    .section-title { font-size: 17px; font-weight: 700; margin: 18px 0 8px 0; }
</style>
""", unsafe_allow_html=True)

# ─── Title ────────────────────────────────────────────────────────────────────
st.title("🗺️ Auto Roadmap Dashboard")
st.caption("Upload your workbook, browse the Roadmap, fill Requests, and export a completed Excel file.")

# ─── File Upload ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "**Upload Sample Auto.xlsx**",
    type=["xlsx"],
    help="Must contain 'RoadMap' and 'Requests' sheets.",
)

if not uploaded:
    st.info("👆 Upload your Excel file above to get started.")
    st.stop()

# ─── Load workbook ────────────────────────────────────────────────────────────
@st.cache_data
def load_data(file_bytes: bytes):
    wb = load_workbook(io.BytesIO(file_bytes))

    # ── RoadMap ──
    roadmap_ws = wb["RoadMap"]
    last_row = roadmap_ws.max_row
    headers = [roadmap_ws.cell(1, c).value for c in range(1, 5)]

    rows = []
    for r in range(2, last_row + 1):
        row = [roadmap_ws.cell(r, c).value for c in range(1, 5)]
        if any(v is not None for v in row):
            rows.append(row)

    df_roadmap = pd.DataFrame(rows, columns=headers)

    # ── Category → {WorkItem: Code} mapping ──
    categories: dict[str, dict[str, str]] = {}
    for _, row in df_roadmap.iterrows():
        cat  = row.iloc[1]   # Column B
        item = row.iloc[2]   # Column C
        code = row.iloc[3]   # Column D
        if cat and item:
            categories.setdefault(str(cat), {})[str(item)] = str(code) if code else ""

    # ── Existing Requests (cols J=10, K=11, L=12) ──
    req_ws = wb["Requests"]
    req_headers = [req_ws.cell(1, c).value for c in range(1, 13)]
    existing_reqs = []
    for r in range(2, req_ws.max_row + 1):
        row = [req_ws.cell(r, c).value for c in range(1, 13)]
        if any(v is not None for v in row):
            existing_reqs.append(row)

    df_requests = pd.DataFrame(existing_reqs, columns=req_headers) if existing_reqs else pd.DataFrame(columns=req_headers)

    return df_roadmap, categories, df_requests, last_row

file_bytes = uploaded.read()
df_roadmap, categories, df_existing_requests, last_row = load_data(file_bytes)

# ─── Session state ────────────────────────────────────────────────────────────
if "new_requests" not in st.session_state:
    st.session_state.new_requests = []

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_roadmap, tab_form, tab_export = st.tabs([
    "📋  RoadMap Viewer",
    "📝  Requests Form",
    "📤  Export Excel",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — ROADMAP VIEWER
# ═══════════════════════════════════════════════════════════════════════════════
with tab_roadmap:
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(f'<div class="stat-card"><div class="stat-label">Total Items</div><div class="stat-value">{len(df_roadmap)}</div></div>', unsafe_allow_html=True)
    with col_b:
        st.markdown(f'<div class="stat-card"><div class="stat-label">Categories</div><div class="stat-value">{len(categories)}</div></div>', unsafe_allow_html=True)
    with col_c:
        total_work_items = sum(len(v) for v in categories.values())
        st.markdown(f'<div class="stat-card"><div class="stat-label">Work Items</div><div class="stat-value">{total_work_items}</div></div>', unsafe_allow_html=True)

    st.markdown("---")

    # Search / filter
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        search = st.text_input("🔍 Search items", placeholder="Type to filter…")
    with col_s2:
        cat_filter = st.selectbox("Filter by Category", ["All"] + sorted(categories.keys()))

    filtered = df_roadmap.copy()
    if cat_filter != "All":
        filtered = filtered[filtered.iloc[:, 1] == cat_filter]
    if search:
        mask = filtered.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)
        filtered = filtered[mask]

    st.dataframe(filtered, use_container_width=True, height=420)

    st.caption(f"Showing **{len(filtered)}** of **{len(df_roadmap)}** rows")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — REQUESTS FORM
# ═══════════════════════════════════════════════════════════════════════════════
with tab_form:
    st.markdown('<div class="section-title">Add a New Request</div>', unsafe_allow_html=True)

    with st.container(border=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            selected_cat = st.selectbox(
                "Category (Column J)",
                options=sorted(categories.keys()),
                key="form_cat",
            )

        with col2:
            work_items = sorted(categories.get(selected_cat, {}).keys())
            selected_item = st.selectbox(
                "Work Item (Column K)",
                options=work_items,
                key="form_item",
            )

        with col3:
            auto_code = categories.get(selected_cat, {}).get(selected_item, "")
            st.text_input(
                "Code (Column L) — auto-filled",
                value=auto_code,
                disabled=True,
                key="form_code",
            )

        if st.button("➕ Add Request", use_container_width=True, type="primary"):
            if selected_cat and selected_item:
                st.session_state.new_requests.append({
                    "Category": selected_cat,
                    "Work Item": selected_item,
                    "Code": auto_code,
                })
                st.success(f"Added: **{selected_cat}** → **{selected_item}** → `{auto_code}`")
            else:
                st.warning("Please select both a Category and a Work Item.")

    # ── Existing requests from file ─────────────────────────────────────────
    if not df_existing_requests.empty:
        st.markdown('<div class="section-title">Existing Requests (from file)</div>', unsafe_allow_html=True)
        j_col = df_existing_requests.columns[9] if len(df_existing_requests.columns) > 9 else "J"
        k_col = df_existing_requests.columns[10] if len(df_existing_requests.columns) > 10 else "K"
        l_col = df_existing_requests.columns[11] if len(df_existing_requests.columns) > 11 else "L"
        st.dataframe(
            df_existing_requests[[j_col, k_col, l_col]].dropna(how="all"),
            use_container_width=True,
        )

    # ── New requests added this session ────────────────────────────────────
    st.markdown('<div class="section-title">New Requests (this session)</div>', unsafe_allow_html=True)

    if st.session_state.new_requests:
        df_new = pd.DataFrame(st.session_state.new_requests)
        st.dataframe(df_new, use_container_width=True)

        col_del, col_clr = st.columns([1, 5])
        with col_del:
            del_idx = st.number_input("Row # to delete", min_value=1, max_value=len(st.session_state.new_requests), step=1)
        with col_clr:
            st.write("")
            st.write("")
            bcol1, bcol2 = st.columns(2)
            with bcol1:
                if st.button("🗑 Delete Row"):
                    st.session_state.new_requests.pop(int(del_idx) - 1)
                    st.rerun()
            with bcol2:
                if st.button("🧹 Clear All"):
                    st.session_state.new_requests = []
                    st.rerun()
    else:
        st.info("No new requests added yet. Use the form above.")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — EXPORT
# ═══════════════════════════════════════════════════════════════════════════════
with tab_export:
    st.markdown('<div class="section-title">Export Summary</div>', unsafe_allow_html=True)

    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown(f'<div class="stat-card"><div class="stat-label">Existing Requests</div><div class="stat-value">{len(df_existing_requests)}</div></div>', unsafe_allow_html=True)
    with col_e2:
        st.markdown(f'<div class="stat-card"><div class="stat-label">New Requests (this session)</div><div class="stat-value">{len(st.session_state.new_requests)}</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.write("Click **Generate & Download** to apply all transformations and export the completed Excel file.")
    st.markdown("""
    The exported file will include:
    - ✅ **RoadmapTable** — formatted Excel table on the RoadMap sheet
    - ✅ **Helper sheet** — hidden sheet with named ranges for all categories
    - ✅ **Data Validation** — dropdown lists for Category (J) and Work Item (K) on Requests sheet
    - ✅ **Auto Code formula** — `IFERROR(INDEX/MATCH)` formula in column L
    - ✅ **Your new requests** — written into the Requests sheet starting at the next empty row
    """)

    if st.button("⚙️ Generate & Download Excel", type="primary", use_container_width=True):
        with st.spinner("Building your Excel file…"):

            wb_out = load_workbook(io.BytesIO(file_bytes))
            roadmap_ws = wb_out["RoadMap"]
            requests_ws = wb_out["Requests"]

            # ── RoadMap table ─────────────────────────────────────────────
            table_name = "RoadmapTable"
            if table_name not in roadmap_ws.tables:
                tbl = Table(displayName=table_name, ref=f"A1:D{last_row}")
                tbl.tableStyleInfo = TableStyleInfo(
                    name="TableStyleMedium2",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False,
                )
                roadmap_ws.add_table(tbl)

            # ── Helper sheet ──────────────────────────────────────────────
            helper = wb_out["Helper"] if "Helper" in wb_out.sheetnames else wb_out.create_sheet("Helper")
            helper.sheet_state = "hidden"

            cats: dict[str, list[str]] = {}
            for r in range(2, last_row + 1):
                cat  = roadmap_ws[f"B{r}"].value
                item = roadmap_ws[f"C{r}"].value
                if cat and item:
                    cats.setdefault(cat, []).append(item)

            helper["A1"] = "Categories"
            for i, cat in enumerate(sorted(cats.keys()), start=2):
                helper[f"A{i}"] = cat

            cat_end = len(cats) + 1
            wb_out.defined_names.add(
                DefinedName("CategoryList", attr_text=f"Helper!$A$2:$A${cat_end}")
            )

            col_num = 2
            for cat, items in cats.items():
                safe_name = cat.replace(" ", "").replace("-", "")
                helper.cell(row=1, column=col_num).value = safe_name
                unique_items = sorted(set(items))
                for idx, item in enumerate(unique_items, start=2):
                    helper.cell(row=idx, column=col_num).value = item
                col_letter = get_column_letter(col_num)
                rng = f"Helper!${col_letter}$2:${col_letter}${len(unique_items)+1}"
                wb_out.defined_names.add(DefinedName(safe_name, attr_text=rng))
                col_num += 1

            # ── Write new requests into sheet ─────────────────────────────
            # Find first empty row in Requests (col J)
            next_row = 2
            while requests_ws[f"J{next_row}"].value is not None:
                next_row += 1

            for req in st.session_state.new_requests:
                requests_ws[f"J{next_row}"] = req["Category"]
                requests_ws[f"K{next_row}"] = req["Work Item"]
                requests_ws[f"L{next_row}"] = req["Code"]
                next_row += 1

            # ── Validation ────────────────────────────────────────────────
            dv_cat = DataValidation(type="list", formula1="=CategoryList")
            requests_ws.add_data_validation(dv_cat)
            dv_cat.add("J2:J1000")

            dv_work = DataValidation(
                type="list",
                formula1='=INDIRECT(SUBSTITUTE(SUBSTITUTE($J2," ",""),"-",""))',
            )
            requests_ws.add_data_validation(dv_work)
            dv_work.add("K2:K1000")

            # ── Auto Code formula ─────────────────────────────────────────
            for row in range(2, 1001):
                requests_ws[f"L{row}"] = (
                    f'=IFERROR(INDEX(RoadMap!$D:$D,MATCH(K{row},RoadMap!$C:$C,0)),"")'
                )

            # ── Save to buffer ────────────────────────────────────────────
            output = io.BytesIO()
            wb_out.save(output)
            output.seek(0)

        st.success("✅ File ready!")
        st.download_button(
            label="📥 Download Sample_Auto_Completed.xlsx",
            data=output,
            file_name="Sample_Auto_Completed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
