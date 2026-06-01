import streamlit as st
import pandas as pd
from datetime import date
import calendar
from PIL import Image, ImageDraw

# =========================================================
# PAGE CONFIGURATION
# =========================================================
try:
    raw_image = Image.open("assets/SpoonLogo.png").convert("RGBA")
    min_dim = min(raw_image.size)
    left = (raw_image.width - min_dim) / 2
    top = (raw_image.height - min_dim) / 2
    right = (raw_image.width + min_dim) / 2
    bottom = (raw_image.height + min_dim) / 2
    squared_image = raw_image.crop((left, top, right, bottom))
    target_size = 120
    resized_image = squared_image.resize((target_size, target_size), Image.Resampling.LANCZOS)
    mask = Image.new("L", (target_size, target_size), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0, target_size, target_size), fill=255)
    page_icon_setting = Image.new("RGBA", (target_size, target_size), (0, 0, 0, 0))
    page_icon_setting.paste(resized_image, (0, 0), mask=mask)
except:
    page_icon_setting = "🔄"

st.set_page_config(
    page_title="Timesheet Reconciliation Dashboard",
    layout="wide",
    page_icon=page_icon_setting
)

st.title("🔄⚙️ Reconciliation Tool - Excel Viewer")
st.markdown("---")

# =========================================================
# SESSION STATE & DEFAULT CALCULATIONS
# =========================================================
today = date.today()
default_start = date(today.year, today.month, 1)
_, last_day = calendar.monthrange(today.year, today.month)
default_end = date(today.year, today.month, last_day)

if "aligned_names" not in st.session_state:
    st.session_state.aligned_names = []

if "mapping_dict" not in st.session_state:
    st.session_state.mapping_dict = {}

if "isolated_person" not in st.session_state:
    st.session_state.isolated_person = None

if "project_managers" not in st.session_state:
    st.session_state.project_managers = []

if "initialized_mapping" not in st.session_state:
    st.session_state.initialized_mapping = False

if "show_tables" not in st.session_state:
    st.session_state.show_tables = False

if "snapshot_start_date" not in st.session_state:
    st.session_state.snapshot_start_date = default_start
if "snapshot_end_date" not in st.session_state:
    st.session_state.snapshot_end_date = default_end
if "snapshot_pms" not in st.session_state:
    st.session_state.snapshot_pms = []

if "start_date_input" not in st.session_state:
    st.session_state["start_date_input"] = default_start
if "end_date_input" not in st.session_state:
    st.session_state["end_date_input"] = default_end


def reset_filters_callback():
    st.session_state["start_date_input"] = default_start
    st.session_state["end_date_input"] = default_end
    st.session_state.project_managers = []
    st.session_state.aligned_names = []
    st.session_state.mapping_dict = {}
    st.session_state.isolated_person = None
    st.session_state.initialized_mapping = False
    st.session_state.show_tables = False
    st.session_state.snapshot_start_date = default_start
    st.session_state.snapshot_end_date = default_end
    st.session_state.snapshot_pms = []


# =========================================================
# DROPDOWN FILTER SECTION (REPLIABLE / HIDEABLE)
# =========================================================
with st.expander("🛠️ Configuration & Filters", expanded=True):
    top_row_col1, top_row_col2, top_row_col3 = st.columns(3)

    with top_row_col1:
        st.markdown("##### 📁 Document Inputs")
        file_beeline = st.file_uploader("Beeline File", type=["xlsx", "xls"], label_visibility="collapsed")
        file_timesheet = st.file_uploader("Timesheet File", type=["xlsx", "xls"], label_visibility="collapsed")
        file_mapping = st.file_uploader("Name Mapping File", type=["xlsx", "xls"], label_visibility="collapsed")

    with top_row_col2:
        st.markdown("##### 📅 Date Range")
        start_date = st.date_input("Start Date", key="start_date_input")
        end_date = st.date_input("End Date", key="end_date_input")

    with top_row_col3:
        st.markdown("##### 👨‍💼 Project Management")

        def add_project_manager():
            pm_clean = st.session_state.pm_input.strip()
            if (
                    pm_clean != ""
                    and pm_clean not in st.session_state.project_managers
            ):
                st.session_state.project_managers.append(pm_clean)
            st.session_state.pm_input = ""

        st.text_input(
            "Type Manager name and press Enter",
            key="pm_input",
            on_change=add_project_manager
        )

        if st.session_state.project_managers:
            pm_cols = st.columns(3)
            pm_to_remove = None
            for i, pm in enumerate(st.session_state.project_managers):
                with pm_cols[i % 3]:
                    if st.button(f"❌ {pm}", key=f"remove_pm_{i}", use_container_width=True):
                        pm_to_remove = pm
            if pm_to_remove:
                st.session_state.project_managers.remove(pm_to_remove)
                st.rerun()

st.markdown("---")

# =========================================================
# ACTION BUTTON TRACK (RUN / CLEAR)
# =========================================================
btn_col1, btn_col2 = st.columns([2, 8])

with btn_col1:
    btn_afficher = st.button("Display Data Tables", type="primary", use_container_width=True)

with btn_col2:
    st.button("🔄 Reset Filters", type="secondary", on_click=reset_filters_callback)


# =========================================================
# EXCEL READ FUNCTION
# =========================================================
def read_excel_safely(uploaded_file, key_suffix, is_beeline=False):
    if uploaded_file is None:
        return None

    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        sheet_names = excel_obj.sheet_names
        default_index = 0
        if is_beeline:
            for i, sheet in enumerate(sheet_names):
                if "BL_Detailed_Approval" in sheet:
                    default_index = i
                    break

        if len(sheet_names) > 1:
            selected_sheet = st.selectbox(
                f"Worksheet for '{uploaded_file.name}':",
                sheet_names,
                index=default_index,
                key=f"selector_{key_suffix}"
            )
        else:
            selected_sheet = sheet_names[0]

        df = pd.read_excel(excel_obj, sheet_name=selected_sheet, dtype=str)
        df.columns = [str(col).strip() for col in df.columns]

        if is_beeline:
            target_col = None
            if "Billing" in df.columns:
                target_col = "Billing"
            elif "md" in df.columns:
                target_col = "md"

            if target_col:
                idx = df.columns.get_loc(target_col)
                df = df.iloc[:, :idx + 1]

        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"❌ Error while reading {uploaded_file.name}: {e}")
        return None


# =========================================================
# TOTAL HOURS CALCULATION (BY INDEX)
# =========================================================
def calculate_total_by_index(df, col_index):
    if df is None or df.empty:
        return 0

    if len(df.columns) > col_index:
        try:
            total = pd.to_numeric(
                df.iloc[:, col_index],
                errors="coerce"
            ).fillna(0).sum()
            return round(total, 2)
        except:
            return 0
    return 0


# =========================================================
# DATA PROCESSING & LOGIC CORE
# =========================================================
has_all_files = file_beeline is not None and file_timesheet is not None and file_mapping is not None
has_at_least_one_pm = len(st.session_state.project_managers) > 0

if btn_afficher:
    if not has_all_files:
        st.error("⚠️ **Error:** All 3 Excel files (Beeline, Timesheet, and Name Mapping) are mandatory.")
        st.session_state.show_tables = False
    elif not has_at_least_one_pm:
        st.error("⚠️ **Error:** You must add at least one Project Manager before displaying the results.")
        st.session_state.show_tables = False
    else:
        st.session_state.snapshot_start_date = start_date
        st.session_state.snapshot_end_date = end_date
        st.session_state.snapshot_pms = list(st.session_state.project_managers)
        st.session_state.show_tables = True
        st.session_state.initialized_mapping = False
        st.session_state.isolated_person = None  # Réinitialise la vue isolée au clic sur Display Data Tables

if st.session_state.show_tables:
    if not has_all_files:
        st.warning("⚠️ Required files have been removed.")
        st.session_state.show_tables = False
    else:
        df_beeline = read_excel_safely(file_beeline, "beeline", is_beeline=True)
        df_timesheet = read_excel_safely(file_timesheet, "timesheet", is_beeline=False)
        pm_match_found = True

        if df_beeline is not None:
            if len(df_beeline.columns) > 0:
                bl_dates = pd.to_datetime(df_beeline.iloc[:, 0], errors='coerce', format='mixed').dt.date
                df_beeline = df_beeline[(bl_dates >= st.session_state.snapshot_start_date) & (
                        bl_dates <= st.session_state.snapshot_end_date)]

            if len(st.session_state.snapshot_pms) > 0 and len(df_beeline.columns) > 14:
                df_beeline_filtered_pm = df_beeline[
                    df_beeline.iloc[:, 14].str.strip().isin(st.session_state.snapshot_pms)]
                if df_beeline_filtered_pm.empty and not df_beeline.empty:
                    pm_match_found = False
                    df_beeline = df_beeline_filtered_pm
                else:
                    df_beeline = df_beeline_filtered_pm

        if df_timesheet is not None:
            if len(df_timesheet.columns) > 2:
                ts_dates = pd.to_datetime(df_timesheet.iloc[:, 2], errors='coerce', format='mixed').dt.date
                df_timesheet = df_timesheet[(ts_dates >= st.session_state.snapshot_start_date) & (
                        ts_dates <= st.session_state.snapshot_end_date)]

            if len(df_timesheet.columns) > 8:
                df_timesheet = df_timesheet[df_timesheet.iloc[:, 8].str.strip() == "No"]

            if not pm_match_found:
                df_timesheet = pd.DataFrame(columns=df_timesheet.columns)

        if file_mapping is not None and not st.session_state.initialized_mapping:
            if not pm_match_found:
                st.session_state.aligned_names = []
                st.session_state.mapping_dict = {}
                st.session_state.initialized_mapping = True
            else:
                try:
                    mapping_excel = pd.ExcelFile(file_mapping)
                    mapping_sheet = mapping_excel.sheet_names[0]
                    mapping_df = pd.read_excel(mapping_excel, sheet_name=mapping_sheet, dtype=str).fillna("")

                    active_beeline_names = set(
                        df_beeline.iloc[:, 13].str.strip().tolist()) if df_beeline is not None and len(
                        df_beeline.columns) > 13 else set()
                    active_timesheet_names = set(
                        df_timesheet.iloc[:, 1].str.strip().tolist()) if df_timesheet is not None and len(
                        df_timesheet.columns) > 1 else set()

                    valid_aligned_names = []
                    mapping_dict = {}

                    for idx, row in mapping_df.iterrows():
                        ts_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
                        bl_name = str(row.iloc[1]).strip() if len(row) > 1 else ""
                        aligned_official = str(row.iloc[2]).strip() if len(row) > 2 else ""

                        if not aligned_official:
                            aligned_official = ts_name if ts_name else bl_name

                        if aligned_official:
                            row_all_names = [ts_name, bl_name, aligned_official]
                            has_beeline_match = any(name in active_beeline_names for name in row_all_names if name)
                            has_timesheet_match = any(name in active_timesheet_names for name in row_all_names if name)

                            if has_beeline_match or has_timesheet_match:
                                if aligned_official not in valid_aligned_names:
                                    valid_aligned_names.append(aligned_official)
                                mapping_dict[aligned_official] = {
                                    "timesheet": ts_name,
                                    "beeline": bl_name
                                }

                    st.session_state.aligned_names = valid_aligned_names
                    st.session_state.mapping_dict = mapping_dict
                    st.session_state.initialized_mapping = True
                except Exception as e:
                    st.error(f"Error while processing Mapping File: {e}")

        # -------------------------------------------------
        # ALIGNED NAMES CARD CONTAINER DISPLAY SECTION
        # -------------------------------------------------
        st.markdown("---")

        with st.expander("👤 Consultants (Active File Filters)", expanded=True):
            if not pm_match_found:
                st.warning(
                    "⚠️ No data found matching the selected Project Manager(s) in Beeline. All outputs are hidden.")
            else:
                if st.session_state.isolated_person is not None:
                    st.info(f"Isolated view currently active for: **{st.session_state.isolated_person}**")
                    if st.button("👁️ Reset View (Show All Users)", type="secondary"):
                        st.session_state.isolated_person = None
                        st.rerun()

                if len(st.session_state.aligned_names) > 0:
                    name_to_remove = None
                    display_names = [n for n in st.session_state.aligned_names if
                                     st.session_state.isolated_person is None or n == st.session_state.isolated_person]

                    cards_per_row = 6
                    for i in range(0, len(display_names), cards_per_row):
                        chunk = display_names[i:i + cards_per_row]
                        grid_cols = st.columns(cards_per_row)

                        for idx, name in enumerate(chunk):
                            orig_idx = st.session_state.aligned_names.index(name)

                            with grid_cols[idx]:
                                with st.container(border=True):
                                    st.markdown(f"👤 **{name}**")

                                    if st.session_state.isolated_person is None:
                                        card_btn1, card_btn2 = st.columns(2)
                                        with card_btn1:
                                            if st.button("👁️ View", key=f"isolate_card_{orig_idx}",
                                                         use_container_width=True):
                                                st.session_state.isolated_person = name
                                                st.rerun()
                                        with card_btn2:
                                            if st.button("❌", key=f"remove_card_{orig_idx}", use_container_width=True):
                                                name_to_remove = name
                                    else:
                                        st.button("Isolated 🔒", key=f"disabled_isolate_{orig_idx}", disabled=True,
                                                  use_container_width=True)

                    if name_to_remove:
                        st.session_state.aligned_names.remove(name_to_remove)
                        if st.session_state.isolated_person == name_to_remove:
                            st.session_state.isolated_person = None
                        st.rerun()
                else:
                    st.warning("No aligned names match the criteria.")

        # -------------------------------------------------
        # FINAL DATA TABLES FILTERING
        # -------------------------------------------------
        if pm_match_found:
            final_ts_filter = []
            final_bl_filter = []
            list_to_process = st.session_state.aligned_names if st.session_state.isolated_person is None else [
                st.session_state.isolated_person]

            for name in list_to_process:
                details = st.session_state.mapping_dict.get(name, {"timesheet": name, "beeline": name})
                if details["timesheet"]:
                    final_ts_filter.append(details["timesheet"])
                if details["beeline"]:
                    final_bl_filter.append(details["beeline"])
                final_ts_filter.append(name)
                final_bl_filter.append(name)

            if len(list_to_process) > 0:
                if df_beeline is not None and len(df_beeline.columns) > 13:
                    df_beeline = df_beeline[df_beeline.iloc[:, 13].str.strip().isin(final_bl_filter)]
                if df_timesheet is not None and len(df_timesheet.columns) > 1:
                    df_timesheet = df_timesheet[df_timesheet.iloc[:, 1].str.strip().isin(final_ts_filter)]
            else:
                df_beeline = pd.DataFrame(columns=df_beeline.columns) if df_beeline is not None else None
                df_timesheet = pd.DataFrame(columns=df_timesheet.columns) if df_timesheet is not None else None

        total_beeline = calculate_total_by_index(df_beeline, 17)
        total_timesheet = calculate_total_by_index(df_timesheet, 10)

        # -------------------------------------------------
        # TWO-COLUMN DATA VIEW AFFORDANCE
        # -------------------------------------------------
        st.markdown("---")
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("🐝 Beeline Data")
            if df_beeline is not None:
                st.metric("🧮 Total Billable Hours", total_beeline)
                st.dataframe(df_beeline, width="stretch", height=600)
                st.success(f"✅ Displayed Rows: {len(df_beeline)}")

        with col2:
            st.subheader("📅 Timesheet Data")
            if df_timesheet is not None:
                st.metric("🧮 Total Workload (hr)", total_timesheet)
                st.dataframe(df_timesheet, width="stretch", height=600)
                st.success(f"✅ Displayed Rows: {len(df_timesheet)}")