import streamlit as st
import pandas as pd
from datetime import date
import calendar
from PIL import Image, ImageDraw
import warnings
import io  # NOUVEAU: Importation pour l'export Excel

# =========================================================
# WARNINGS OPTIMIZATION
# =========================================================
# Masquer les avertissements de styles d'openpyxl dans la console
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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

if "aligned_names" not in st.session_state: st.session_state.aligned_names = []
if "mapping_dict" not in st.session_state: st.session_state.mapping_dict = {}
if "isolated_person" not in st.session_state: st.session_state.isolated_person = None
if "selected_consultants" not in st.session_state: st.session_state.selected_consultants = []
if "temp_checkboxes" not in st.session_state: st.session_state.temp_checkboxes = {}
if "temp_project_checkboxes" not in st.session_state: st.session_state.temp_project_checkboxes = {}
if "project_managers" not in st.session_state: st.session_state.project_managers = []
if "initialized_mapping" not in st.session_state: st.session_state.initialized_mapping = False
if "show_tables" not in st.session_state: st.session_state.show_tables = False
if "multiselect_enabled" not in st.session_state: st.session_state.multiselect_enabled = False

if "snapshot_start_date" not in st.session_state: st.session_state.snapshot_start_date = default_start
if "snapshot_end_date" not in st.session_state: st.session_state.snapshot_end_date = default_end
if "snapshot_pms" not in st.session_state: st.session_state.snapshot_pms = []

if "start_date_input" not in st.session_state:
    st.session_state["start_date_input"] = default_start
if "end_date_input" not in st.session_state:
    st.session_state["end_date_input"] = default_end

if "warning_only_mode" not in st.session_state:
    st.session_state.warning_only_mode = False

if "col_bl_date" not in st.session_state: st.session_state.col_bl_date = ""
if "col_bl_name" not in st.session_state: st.session_state.col_bl_name = ""
if "col_bl_hours" not in st.session_state: st.session_state.col_bl_hours = ""
if "col_bl_pm" not in st.session_state: st.session_state.col_bl_pm = ""
if "col_bl_project" not in st.session_state: st.session_state.col_bl_project = ""
if "col_bl_desc" not in st.session_state: st.session_state.col_bl_desc = ""
if "filter_bl_project_val" not in st.session_state: st.session_state.filter_bl_project_val = ["[All Projects]"]

if "col_ts_date" not in st.session_state: st.session_state.col_ts_date = ""
if "col_ts_name" not in st.session_state: st.session_state.col_ts_name = ""
if "col_ts_hours" not in st.session_state: st.session_state.col_ts_hours = ""
if "col_ts_lock" not in st.session_state: st.session_state.col_ts_lock = ""
if "col_ts_project" not in st.session_state: st.session_state.col_ts_project = ""
if "col_ts_desc" not in st.session_state: st.session_state.col_ts_desc = ""
if "filter_ts_project_val" not in st.session_state: st.session_state.filter_ts_project_val = "[All Projects]"


def reset_filters_callback():
    st.session_state["start_date_input"] = default_start
    st.session_state["end_date_input"] = default_end
    st.session_state.project_managers = []
    st.session_state.aligned_names = []
    st.session_state.mapping_dict = {}
    st.session_state.isolated_person = None
    st.session_state.selected_consultants = []
    st.session_state.temp_checkboxes = {}
    st.session_state.temp_project_checkboxes = {}
    st.session_state.initialized_mapping = False
    st.session_state.show_tables = False
    st.session_state.snapshot_start_date = default_start
    st.session_state.snapshot_end_date = default_end
    st.session_state.snapshot_pms = []
    st.session_state.warning_only_mode = False
    st.session_state.multiselect_enabled = False
    st.session_state.filter_ts_project_val = "[All Projects]"
    st.session_state.filter_bl_project_val = ["[All Projects]"]


# =========================================================
# DROPDOWN FILTER SECTION (REPLIABLE / HIDEABLE)
# =========================================================
with st.expander("Monitored Inputs & Columns Mapping Configuration", expanded=True):
    top_row_col1, top_row_col2, top_row_col3 = st.columns(3)

    cols_beeline_list = []
    cols_timesheet_list = []
    unique_ts_projects = ["[All Projects]"]
    unique_bl_projects = ["[All Projects]"]

    with top_row_col1:
        st.markdown("##### 📁 Document Inputs")

        file_beeline = st.file_uploader("Upload Beeline Export File (.xlsx, .xls)", type=["xlsx", "xls"])
        sheet_beeline = None
        if file_beeline is not None:
            try:
                xl_beeline = pd.ExcelFile(file_beeline)
                sheets = xl_beeline.sheet_names
                sheet_beeline = st.selectbox(f"Worksheet for '{file_beeline.name}':", sheets, index=0,
                                             key="selector_beeline") if len(sheets) > 1 else sheets[0]

                df_bl_init = pd.read_excel(xl_beeline, sheet_name=sheet_beeline, dtype=str).fillna("")
                df_bl_init.columns = [str(c).strip() for c in df_bl_init.columns]
                cols_beeline_list = [c for c in df_bl_init.columns if c and not c.startswith("Unnamed:")]
            except Exception as e:
                st.error(f"Error inspecting Beeline tabs: {e}")

        file_timesheet = st.file_uploader("Upload OneEye Timesheet File (.xlsx, .xls)", type=["xlsx", "xls"])
        sheet_timesheet = None
        if file_timesheet is not None:
            try:
                xl_timesheet = pd.ExcelFile(file_timesheet)
                sheets = xl_timesheet.sheet_names
                sheet_timesheet = st.selectbox(f"Worksheet for '{file_timesheet.name}':", sheets, index=0,
                                               key="selector_timesheet") if len(sheets) > 1 else sheets[0]

                df_ts_init = pd.read_excel(xl_timesheet, sheet_name=sheet_timesheet, dtype=str).fillna("")
                df_ts_init.columns = [str(c).strip() for c in df_ts_init.columns]
                cols_timesheet_list = [c for c in df_ts_init.columns if c and not c.startswith("Unnamed:")]
            except Exception as e:
                st.error(f"Error inspecting Timesheet tabs: {e}")

        file_mapping = st.file_uploader("Upload Name Mapping Reference File (.xlsx, .xls)", type=["xlsx", "xls"])
        sheet_mapping = None
        if file_mapping is not None:
            try:
                xl_mapping = pd.ExcelFile(file_mapping)
                sheets = xl_mapping.sheet_names
                sheet_mapping = st.selectbox(f"Worksheet for '{file_mapping.name}':", sheets, index=0,
                                             key="selector_mapping") if len(sheets) > 1 else sheets[0]
            except Exception as e:
                st.error(f"Error inspecting Mapping tabs: {e}")

    with top_row_col2:
        st.markdown("##### 📅 Date Range")
        start_date = st.date_input("Start Date", key="start_date_input")
        end_date = st.date_input("End Date", key="end_date_input")

    with top_row_col3:
        st.markdown("##### 👨‍💼 Project Management")


        def add_project_manager():
            pm_clean = st.session_state.pm_input.strip()
            if pm_clean != "" and pm_clean not in st.session_state.project_managers:
                st.session_state.project_managers.append(pm_clean)
            st.session_state.pm_input = ""


        st.text_input("Type Manager name and press Enter", key="pm_input", on_change=add_project_manager)

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

    if file_beeline is not None or file_timesheet is not None:
        st.markdown("---")
        st.markdown("##### ⚙️ Columns Custom Mapping Selection")
        map_col1, map_col2 = st.columns(2)

        with map_col1:
            st.markdown("**🐝 Beeline Sheet Fields Mapping**")
            if cols_beeline_list:
                def_bl_date = cols_beeline_list.index("Work Date") if "Work Date" in cols_beeline_list else 0
                def_bl_name = cols_beeline_list.index(
                    "First Name Last Name") if "First Name Last Name" in cols_beeline_list else 0
                def_bl_hrs = cols_beeline_list.index("Billable Hours") if "Billable Hours" in cols_beeline_list else 0
                def_bl_pm = cols_beeline_list.index("Project Manager") if "Project Manager" in cols_beeline_list else 0
                def_bl_proj = cols_beeline_list.index(
                    "Clarity Project - Clarity Project Name") if "Clarity Project - Clarity Project Name" in cols_beeline_list else (
                    5 if len(cols_beeline_list) > 5 else 0)
                def_bl_desc = cols_beeline_list.index(
                    "Timesheet Detail Comments") if "Timesheet Detail Comments" in cols_beeline_list else (
                    8 if len(cols_beeline_list) > 8 else 0)

                sel_bl_date = st.selectbox("Work Date Column (Beeline)", cols_beeline_list, index=def_bl_date)
                sel_bl_name = st.selectbox("Consultant Name Column (Beeline)", cols_beeline_list, index=def_bl_name)
                sel_bl_hours = st.selectbox("Billable Hours Column (Beeline)", cols_beeline_list, index=def_bl_hrs)
                sel_bl_pm = st.selectbox("Project Manager Column (Beeline)", cols_beeline_list, index=def_bl_pm)
                sel_bl_project = st.selectbox("Project Column (Beeline)", cols_beeline_list, index=def_bl_proj)
                sel_bl_desc = st.selectbox("Description Column (Beeline)", cols_beeline_list, index=def_bl_desc)

                if sel_bl_project in df_bl_init.columns and sel_bl_pm in df_bl_init.columns:
                    if st.session_state.project_managers:
                        cleaned_pms = [pm.replace(chr(34), "").strip() for pm in st.session_state.project_managers if
                                       pm.strip()]
                        regex_pattern = "|".join(cleaned_pms)

                        if regex_pattern:
                            df_filtered_pm_init = df_bl_init[
                                df_bl_init[sel_bl_pm].str.contains(regex_pattern, case=False, na=False)]
                            bl_projs_extracted = sorted(df_filtered_pm_init[sel_bl_project].str.strip().unique())
                        else:
                            bl_projs_extracted = sorted(df_bl_init[sel_bl_project].str.strip().unique())
                    else:
                        bl_projs_extracted = sorted(df_bl_init[sel_bl_project].str.strip().unique())

                    if not bl_projs_extracted:
                        bl_projs_extracted = []

                    unique_bl_projects = ["[All Projects]"] + [p for p in bl_projs_extracted if p != ""]

                    st.markdown("<br><b>🎯 Select Project Filters (Beeline) :</b>", unsafe_allow_html=True)

                    for p_item in unique_bl_projects:
                        p_key = f"proj_bl_state_{p_item}"
                        if p_key not in st.session_state.temp_project_checkboxes:
                            st.session_state.temp_project_checkboxes[
                                p_key] = p_item in st.session_state.filter_bl_project_val

                    proj_cards_cols = st.columns(2)
                    for idx_p, p_item in enumerate(unique_bl_projects):
                        p_key = f"proj_bl_state_{p_item}"
                        with proj_cards_cols[idx_p % 2]:
                            current_p_state = st.session_state.temp_project_checkboxes.get(p_key, False)
                            new_p_state = st.checkbox(f"{p_item}", key=f"widget_proj_bl_{idx_p}", value=current_p_state)
                            if new_p_state != current_p_state:
                                st.session_state.temp_project_checkboxes[p_key] = new_p_state
                                st.rerun()

                    active_selected_projects = [p for p in unique_bl_projects if
                                                st.session_state.temp_project_checkboxes.get(f"proj_bl_state_{p}",
                                                                                             False)]
                    if not active_selected_projects:
                        active_selected_projects = ["[All Projects]"]

                    sel_bl_project_val = active_selected_projects
                else:
                    sel_bl_project_val = ["[All Projects]"]
            else:
                st.info("Upload Beeline File to map its columns")

        with map_col2:
            st.markdown("**📅 OneEye Timesheet Fields Mapping**")
            if cols_timesheet_list:
                def_ts_date = cols_timesheet_list.index("Date") if "Date" in cols_timesheet_list else 0
                def_ts_name = cols_timesheet_list.index(
                    "Owner: Full Name") if "Owner: Full Name" in cols_timesheet_list else (
                    cols_timesheet_list.index("User") if "User" in cols_timesheet_list else 0)
                def_ts_hrs = cols_timesheet_list.index(
                    "Entered Workload (hr)") if "Entered Workload (hr)" in cols_timesheet_list else (
                    cols_timesheet_list.index("Workload (hr)") if "Workload (hr)" in cols_timesheet_list else 0)
                def_ts_lock = cols_timesheet_list.index(
                    "Non Billable ?") if "Non Billable ?" in cols_timesheet_list else (cols_timesheet_list.index(
                    "Non Billable Task or Request ?") if "Non Billable Task or Request ?" in cols_timesheet_list else 0)
                def_ts_proj = cols_timesheet_list.index(
                    "Project: Project Name") if "Project: Project Name" in cols_timesheet_list else (
                    cols_timesheet_list.index(
                        "Reporting Line Name") if "Reporting Line Name" in cols_timesheet_list else 0)
                def_ts_desc = cols_timesheet_list.index("Description") if "Description" in cols_timesheet_list else (
                    cols_timesheet_list.index("Comments") if "Comments" in cols_timesheet_list else (
                        12 if len(cols_timesheet_list) > 12 else 0))

                sel_ts_date = st.selectbox("Date Column (Timesheet)", cols_timesheet_list, index=def_ts_date)
                sel_ts_name = st.selectbox("Consultant Name Column (Timesheet)", cols_timesheet_list, index=def_ts_name)
                sel_ts_hours = st.selectbox("Workload Hours Column (Timesheet)", cols_timesheet_list, index=def_ts_hrs)
                sel_ts_lock = st.selectbox("Non Billable Filtering Column (Timesheet)", cols_timesheet_list,
                                           index=def_ts_lock)
                sel_ts_project = st.selectbox("Project Column (Timesheet)", cols_timesheet_list, index=def_ts_proj)
                sel_ts_desc = st.selectbox("Description Column (Timesheet)", cols_timesheet_list, index=def_ts_desc)

                if sel_ts_project in df_ts_init.columns:
                    projs_extracted = sorted(df_ts_init[sel_ts_project].str.strip().unique())
                    unique_ts_projects.extend([p for p in projs_extracted if p != ""])

                    try:
                        def_proj_val_idx = unique_ts_projects.index(st.session_state.filter_ts_project_val)
                    except:
                        def_proj_val_idx = 0

                    sel_ts_project_val = st.selectbox("🎯 Select Project Filter (Timesheet)", unique_ts_projects,
                                                      index=def_proj_val_idx)
                else:
                    sel_ts_project_val = "[All Projects]"
            else:
                st.info("Upload Timesheet File to map its columns")

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
# EXCEL READ/WRITE FUNCTIONS
# =========================================================
def read_excel_safely(uploaded_file, selected_sheet):
    if uploaded_file is None or selected_sheet is None:
        return None
    try:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, dtype=str)
        df.columns = [str(col).strip() for col in df.columns]
        df = df.loc[:, ~df.columns.str.startswith('Unnamed:')]
        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"❌ Error while reading {uploaded_file.name} ({selected_sheet}): {e}")
        return None


def calculate_total_by_column_name(df, column_name):
    if df is None or df.empty or not column_name or column_name not in df.columns:
        return 0
    try:
        clean_series = df[column_name].astype(str).str.replace(',', '.')
        total = pd.to_numeric(clean_series, errors="coerce").fillna(0).sum()
        return round(total, 2)
    except:
        return 0


# NOUVEAU: Fonction pour convertir un DataFrame en Excel (en mémoire)
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Exported Data')
    processed_data = output.getvalue()
    return processed_data


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

        st.session_state.col_bl_date = sel_bl_date
        st.session_state.col_bl_name = sel_bl_name
        st.session_state.col_bl_hours = sel_bl_hours
        st.session_state.col_bl_pm = sel_bl_pm
        st.session_state.col_bl_project = sel_bl_project
        st.session_state.col_bl_desc = sel_bl_desc
        st.session_state.filter_bl_project_val = sel_bl_project_val

        st.session_state.col_ts_date = sel_ts_date
        st.session_state.col_ts_name = sel_ts_name
        st.session_state.col_ts_hours = sel_ts_hours
        st.session_state.col_ts_lock = sel_ts_lock
        st.session_state.col_ts_project = sel_ts_project
        st.session_state.col_ts_desc = sel_ts_desc
        st.session_state.filter_ts_project_val = sel_ts_project_val

        st.session_state.show_tables = True
        st.session_state.initialized_mapping = False
        st.session_state.isolated_person = None
        st.session_state.selected_consultants = []
        st.session_state.temp_checkboxes = {}
        st.session_state.warning_only_mode = False
        st.session_state.multiselect_enabled = False

if st.session_state.show_tables:
    if not has_all_files:
        st.warning("⚠️ Required files have been removed.")
        st.session_state.show_tables = False
    else:
        df_beeline_raw = read_excel_safely(file_beeline, sheet_beeline)
        df_timesheet_raw = read_excel_safely(file_timesheet, sheet_timesheet)

        validation_passed = True
        pm_match_found = True

        if df_beeline_raw is not None:
            required_bl = [st.session_state.col_bl_date, st.session_state.col_bl_name, st.session_state.col_bl_hours,
                           st.session_state.col_bl_pm, st.session_state.col_bl_project, st.session_state.col_bl_desc]
            if not all(c in df_beeline_raw.columns for c in required_bl):
                st.error(f"⚠️ **Mapping Format Error:** Selected columns not found in Beeline sheet headers.")
                validation_passed = False
                st.session_state.show_tables = False
            else:
                try:
                    df_beeline_raw[st.session_state.col_bl_date] = pd.to_datetime(
                        df_beeline_raw[st.session_state.col_bl_date], errors='coerce', dayfirst=True).dt.date
                    df_beeline = df_beeline_raw[
                        (df_beeline_raw[st.session_state.col_bl_date] >= st.session_state.snapshot_start_date) & (
                                df_beeline_raw[st.session_state.col_bl_date] <= st.session_state.snapshot_end_date)]
                except:
                    st.error(
                        f"⚠️ **Format Error:** Column '{st.session_state.col_bl_date}' does not contain valid dates.")
                    validation_passed = False
                    st.session_state.show_tables = False

                if validation_passed and len(st.session_state.snapshot_pms) > 0:
                    cleaned_snapshot_pms = [p.strip() for p in st.session_state.snapshot_pms if p.strip()]
                    regex_bl_run = "|".join(cleaned_snapshot_pms)
                    df_beeline = df_beeline[
                        df_beeline[st.session_state.col_bl_pm].str.contains(regex_bl_run, case=False, na=False)]

                    if st.session_state.filter_bl_project_val and "[All Projects]" not in st.session_state.filter_bl_project_val:
                        df_beeline = df_beeline[df_beeline[st.session_state.col_bl_project].str.strip().isin(
                            st.session_state.filter_bl_project_val)]
                    if df_beeline.empty:
                        pm_match_found = False

        if validation_passed and df_timesheet_raw is not None:
            required_ts = [st.session_state.col_ts_date, st.session_state.col_ts_name, st.session_state.col_ts_hours,
                           st.session_state.col_ts_lock, st.session_state.col_ts_project, st.session_state.col_ts_desc]
            if not all(c in df_timesheet_raw.columns for c in required_ts):
                st.error(f"⚠️ **Mapping Format Error:** Selected columns not found in Timesheet sheet headers.")
                validation_passed = False
                st.session_state.show_tables = False
            else:
                try:
                    df_timesheet_raw[st.session_state.col_ts_date] = pd.to_datetime(
                        df_timesheet_raw[st.session_state.col_ts_date], errors='coerce', dayfirst=True).dt.date
                    df_timesheet = df_timesheet_raw[
                        (df_timesheet_raw[st.session_state.col_ts_date] >= st.session_state.snapshot_start_date) & (
                                df_timesheet_raw[st.session_state.col_ts_date] <= st.session_state.snapshot_end_date)]
                except:
                    st.error(
                        f"⚠️ **Format Error:** Column '{st.session_state.col_ts_date}' does not contain valid dates.")
                    validation_passed = False
                    st.session_state.show_tables = False

                if validation_passed:
                    target_val = "false" if "Non Billable ?" in df_timesheet.columns else "no"
                    df_timesheet = df_timesheet[
                        df_timesheet[st.session_state.col_ts_lock].astype(str).str.strip().str.lower() == target_val]
                    if st.session_state.filter_ts_project_val != "[All Projects]":
                        df_timesheet = df_timesheet[df_timesheet[
                                                        st.session_state.col_ts_project].str.strip() == st.session_state.filter_ts_project_val]
                    if not pm_match_found:
                        df_timesheet = pd.DataFrame(columns=df_timesheet.columns)

        warning_persons_list = []

        if validation_passed and file_mapping is not None and not st.session_state.initialized_mapping:
            if not pm_match_found:
                st.session_state.aligned_names = []
                st.session_state.mapping_dict = {}
                st.session_state.initialized_mapping = True
            else:
                try:
                    mapping_df = pd.read_excel(file_mapping, sheet_name=sheet_mapping, dtype=str).fillna("")
                    expected_headers = ["Timesheet", "Beeline", "Aligned Name"]
                    actual_headers = list(mapping_df.columns)

                    if not (len(actual_headers) >= 3 and all(
                            expected in actual_headers[:3] for expected in expected_headers)):
                        st.error(
                            f"⚠️ **Format Error:** The mapping file '{file_mapping.name}' requires headers: **Timesheet**, **Beeline**, **Aligned Name**.")
                        validation_passed = False
                        st.session_state.show_tables = False
                    else:
                        active_beeline_names = set(df_beeline[
                                                       st.session_state.col_bl_name].str.strip().tolist()) if df_beeline is not None else set()
                        active_timesheet_names = set(df_timesheet[
                                                         st.session_state.col_ts_name].str.strip().tolist()) if df_timesheet is not None else set()

                        valid_aligned_names = []
                        mapping_dict = {}

                        for idx, row in mapping_df.iterrows():
                            try:
                                ts_name = str(row["Timesheet"]).strip()
                                bl_name = str(row["Beeline"]).strip()
                                aligned_official = str(row["Aligned Name"]).strip()

                                if not aligned_official: aligned_official = ts_name if ts_name else bl_name

                                if aligned_official:
                                    row_all_names = [ts_name, bl_name, aligned_official]
                                    has_beeline_match = any(
                                        name in active_beeline_names for name in row_all_names if name)
                                    has_timesheet_match = any(
                                        name in active_timesheet_names for name in row_all_names if name)

                                    if has_beeline_match or has_timesheet_match:
                                        if aligned_official not in valid_aligned_names:
                                            valid_aligned_names.append(aligned_official)
                                        mapping_dict[aligned_official] = {
                                            "timesheet": ts_name,
                                            "beeline": bl_name
                                        }
                            except Exception:
                                continue

                        mapped_bl_names = [v["beeline"] for v in mapping_dict.values() if v["beeline"]]
                        mapped_ts_names = [v["timesheet"] for v in mapping_dict.values() if v["timesheet"]]

                        for name in active_beeline_names:
                            if name and name not in mapped_bl_names and name not in valid_aligned_names:
                                valid_aligned_names.append(name)
                                mapping_dict[name] = {"timesheet": name, "beeline": name}

                        for name in active_timesheet_names:
                            if name and name not in mapped_ts_names and name not in valid_aligned_names:
                                valid_aligned_names.append(name)
                                mapping_dict[name] = {"timesheet": name, "beeline": name}

                        st.session_state.aligned_names = valid_aligned_names
                        st.session_state.mapping_dict = mapping_dict
                        st.session_state.initialized_mapping = True
                except Exception as e:
                    st.error(f"❌ Error while processing Mapping File '{file_mapping.name}': {e}")
                    validation_passed = False
                    st.session_state.show_tables = False

        if validation_passed and pm_match_found:
            for name in st.session_state.aligned_names:
                details = st.session_state.mapping_dict.get(name, {"timesheet": name, "beeline": name})
                cond_ts = [details["timesheet"], name]
                cond_bl = [details["beeline"], name]

                sub_bl = df_beeline[df_beeline[st.session_state.col_bl_name].str.strip().isin(
                    cond_bl)] if df_beeline is not None else None
                sub_ts = df_timesheet[df_timesheet[st.session_state.col_ts_name].str.strip().isin(
                    cond_ts)] if df_timesheet is not None else None

                hrs_bl = calculate_total_by_column_name(sub_bl, st.session_state.col_bl_hours)
                hrs_ts = calculate_total_by_column_name(sub_ts, st.session_state.col_ts_hours)
                if hrs_bl != hrs_ts or hrs_bl == 0:
                    warning_persons_list.append(name)

        # -------------------------------------------------
        # ALIGNED NAMES CARD CONTAINER DISPLAY SECTION
        # -------------------------------------------------
        if validation_passed:
            st.markdown("---")

            with st.expander("👤 Consultants (Active File Filters)", expanded=True):
                header_col1, header_col2 = st.columns([8, 4])
                with header_col1:
                    if st.session_state.isolated_person is not None:
                        st.info(f"Isolated view currently active for: **{st.session_state.isolated_person}**")
                    elif len(st.session_state.selected_consultants) > 0:
                        st.info(
                            f"🎯 **Target View active:** Showing grid matrix for **{len(st.session_state.selected_consultants)}** isolated consultant profile(s).")

                with header_col2:
                    if st.session_state.isolated_person is not None or len(st.session_state.selected_consultants) > 0:
                        if st.button("👁️ Show All Consultants", key="reset_warning_btn", type="primary",
                                     use_container_width=True):
                            st.session_state.warning_only_mode = False
                            st.session_state.isolated_person = None
                            st.session_state.selected_consultants = []
                            st.session_state.temp_checkboxes = {}
                            st.session_state.multiselect_enabled = False
                            st.rerun()
                    elif not st.session_state.multiselect_enabled:
                        if st.button("⚙️ Multi-Selection Mode", key="activate_multiselect_btn", type="secondary",
                                     use_container_width=True):
                            st.session_state.multiselect_enabled = True
                            st.rerun()

                if st.session_state.isolated_person is not None:
                    display_names = [st.session_state.isolated_person]
                elif len(st.session_state.selected_consultants) > 0:
                    display_names = st.session_state.selected_consultants
                else:
                    display_names = st.session_state.aligned_names

                if len(display_names) > 0:
                    name_to_remove = None
                    cards_per_row = 6

                    if st.session_state.multiselect_enabled and len(st.session_state.selected_consultants) == 0:
                        with st.form(key="multiselect_consultants_form", clear_on_submit=False):

                            for i in range(0, len(display_names), cards_per_row):
                                chunk = display_names[i:i + cards_per_row]
                                grid_cols = st.columns(cards_per_row)

                                for idx, name in enumerate(chunk):
                                    orig_idx = st.session_state.aligned_names.index(name)
                                    details = st.session_state.mapping_dict.get(name,
                                                                                {"timesheet": name, "beeline": name})
                                    cond_ts = [details["timesheet"], name]
                                    cond_bl = [details["beeline"], name]

                                    sub_bl = df_beeline[df_beeline[st.session_state.col_bl_name].str.strip().isin(
                                        cond_bl)] if df_beeline is not None else None
                                    sub_ts = df_timesheet[df_timesheet[st.session_state.col_ts_name].str.strip().isin(
                                        cond_ts)] if df_timesheet is not None else None

                                    hrs_bl = calculate_total_by_column_name(sub_bl, st.session_state.col_bl_hours)
                                    hrs_ts = calculate_total_by_column_name(sub_ts, st.session_state.col_ts_hours)
                                    is_equivalent = (name not in warning_persons_list)

                                    with grid_cols[idx]:
                                        st.checkbox("Select", key=f"form_check_{orig_idx}", value=False,
                                                    label_visibility="collapsed")

                                        if not is_equivalent:
                                            st.markdown(
                                                f"""<div style="background-color: #FFEBEE; border: 1.5px solid #FF1744; border-radius: 5px; padding: 10px; margin-bottom: 10px;"><span style="color: #B71C1C; font-weight: bold;">⚠️ {name}</span><div style="font-size: 11px; color: #555; margin-top: 5px;">Beeline: {hrs_bl}h | TS: {hrs_ts}h</div></div>""",
                                                unsafe_allow_html=True)
                                        else:
                                            st.markdown(
                                                f"""<div style="background-color: #E8F5E9; border: 1.5px solid #00E676; border-radius: 5px; padding: 10px; margin-bottom: 10px;"><span style="color: #1B5E20; font-weight: bold;">👤 {name}</span><div style="font-size: 11px; color: #555; margin-top: 5px;">Heures validées: {hrs_bl}h</div></div>""",
                                                unsafe_allow_html=True)

                            st.markdown("<br>", unsafe_allow_html=True)
                            submit_btn = st.form_submit_button("👁️ View Selected Consultants", use_container_width=True,
                                                               type="primary")

                            if submit_btn:
                                verified_selections = []
                                for idx_name, current_name in enumerate(st.session_state.aligned_names):
                                    if st.session_state.get(f"form_check_{idx_name}", False):
                                        verified_selections.append(current_name)

                                if len(verified_selections) > 0:
                                    st.session_state.selected_consultants = verified_selections
                                    st.session_state.isolated_person = None
                                    st.rerun()
                                else:
                                    st.error("⚠️ Please check at least one consultant before validating.")
                    else:
                        for i in range(0, len(display_names), cards_per_row):
                            chunk = display_names[i:i + cards_per_row]
                            grid_cols = st.columns(cards_per_row)

                            for idx, name in enumerate(chunk):
                                orig_idx = st.session_state.aligned_names.index(name)
                                details = st.session_state.mapping_dict.get(name, {"timesheet": name, "beeline": name})
                                cond_ts = [details["timesheet"], name]
                                cond_bl = [details["beeline"], name]

                                sub_bl = df_beeline[df_beeline[st.session_state.col_bl_name].str.strip().isin(
                                    cond_bl)] if df_beeline is not None else None
                                sub_ts = df_timesheet[df_timesheet[st.session_state.col_ts_name].str.strip().isin(
                                    cond_ts)] if df_timesheet is not None else None

                                hrs_bl = calculate_total_by_column_name(sub_bl, st.session_state.col_bl_hours)
                                hrs_ts = calculate_total_by_column_name(sub_ts, st.session_state.col_ts_hours)
                                is_equivalent = (name not in warning_persons_list)

                                with grid_cols[idx]:
                                    if not is_equivalent:
                                        st.markdown(
                                            f"""<div style="background-color: #FFEBEE; border: 1.5px solid #FF1744; border-radius: 5px; padding: 10px; margin-bottom: 10px;"><span style="color: #B71C1C; font-weight: bold;">⚠️ {name}</span><div style="font-size: 11px; color: #555; margin-top: 5px;">Beeline: {hrs_bl}h | TS: {hrs_ts}h</div></div>""",
                                            unsafe_allow_html=True)
                                    else:
                                        st.markdown(
                                            f"""<div style="background-color: #E8F5E9; border: 1.5px solid #00E676; border-radius: 5px; padding: 10px; margin-bottom: 10px;"><span style="color: #1B5E20; font-weight: bold;">👤 {name}</span><div style="font-size: 11px; color: #555; margin-top: 5px;">Heures validées: {hrs_bl}h</div></div>""",
                                            unsafe_allow_html=True)

                                    if st.session_state.isolated_person is None and len(
                                            st.session_state.selected_consultants) == 0:
                                        if st.button("👁️ View", key=f"isolate_card_{orig_idx}",
                                                     use_container_width=True):
                                            st.session_state.isolated_person = name
                                            st.session_state.warning_only_mode = False
                                            st.rerun()

                        if name_to_remove:
                            st.session_state.aligned_names.remove(name_to_remove)
                            if st.session_state.isolated_person == name_to_remove: st.session_state.isolated_person = None
                            st.rerun()

            # -------------------------------------------------
            # FINAL DATA TABLES FILTERING & GAP CALCULATION
            # -------------------------------------------------
            df_bl_display = pd.DataFrame()
            df_ts_display = pd.DataFrame()

            if pm_match_found:
                final_ts_filter = []
                final_bl_filter = []

                if st.session_state.isolated_person is not None:
                    list_to_process = [st.session_state.isolated_person]
                elif len(st.session_state.selected_consultants) > 0:
                    list_to_process = st.session_state.selected_consultants
                else:
                    list_to_process = st.session_state.aligned_names

                for name in list_to_process:
                    details = st.session_state.mapping_dict.get(name, {"timesheet": name, "beeline": name})
                    if details["timesheet"]: final_ts_filter.append(details["timesheet"])
                    if details["beeline"]: final_bl_filter.append(details["beeline"])
                    final_ts_filter.append(name)
                    final_bl_filter.append(name)

                df_bl_display = df_beeline.copy() if df_beeline is not None else pd.DataFrame()
                df_ts_display = df_timesheet.copy() if df_timesheet is not None else pd.DataFrame()

                if len(list_to_process) > 0:
                    if not df_bl_display.empty:
                        df_bl_display = df_bl_display[
                            df_bl_display[st.session_state.col_bl_name].str.strip().isin(final_bl_filter)]
                    if not df_ts_display.empty:
                        df_ts_display = df_ts_display[
                            df_ts_display[st.session_state.col_ts_name].str.strip().isin(final_ts_filter)]
                else:
                    df_bl_display = pd.DataFrame(
                        columns=df_beeline.columns) if df_beeline is not None else pd.DataFrame()
                    df_ts_display = pd.DataFrame(
                        columns=df_timesheet.columns) if df_timesheet is not None else pd.DataFrame()

            total_beeline = calculate_total_by_column_name(df_bl_display, st.session_state.col_bl_hours)
            total_timesheet = calculate_total_by_column_name(df_ts_display, st.session_state.col_ts_hours)
            global_gap = round(total_beeline - total_timesheet, 2)

            # -------------------------------------------------
            # METRICS DISPLAY SECTION (TOP BAR)
            # -------------------------------------------------
            st.markdown("---")
            metric_col1, metric_col2, metric_col3 = st.columns(3)

            with metric_col1:
                st.metric("🐝 Beeline Total Hours", f"{total_beeline} hrs")
            with metric_col2:
                st.metric("📅 Timesheet Total Hours", f"{total_timesheet} hrs")
            with metric_col3:
                gap_delta = f"⚠️ Discrepancy" if global_gap != 0 else "✅ Perfect Match"
                st.metric("📊 Gap (Beeline - TS)", f"{global_gap} hrs", delta=gap_delta,
                          delta_color="inverse" if global_gap != 0 else "normal")

            # -------------------------------------------------
            # TWO-COLUMN DATA VIEW AFFORDANCE
            # -------------------------------------------------
            st.markdown("---")

            if st.session_state.isolated_person is not None:
                st.info(
                    f"👁️ **Active View:** Tables and cards are currently isolated on profile: **{st.session_state.isolated_person}**.")
            elif len(st.session_state.selected_consultants) > 0:
                st.info(
                    f"🎯 **Active View:** Aligned Matrix is isolated on **{len(st.session_state.selected_consultants)}** custom selected consultants.")

            info_msg = []
            if st.session_state.filter_bl_project_val and "[All Projects]" not in st.session_state.filter_bl_project_val:
                info_msg.append(f"Beeline Projects: **{', '.join(st.session_state.filter_bl_project_val)}**")
            if st.session_state.filter_ts_project_val != "[All Projects]":
                info_msg.append(f"Timesheet Project: **{st.session_state.filter_ts_project_val}**")
            if info_msg: st.info(f"🎯 **Active Target Filters:** " + " | ".join(info_msg))

            # =========================================================
            # ADVANCED ALIGNED RECONCILIATION VIEW (SIDE-BY-SIDE)
            # =========================================================
            st.subheader("📊 Aligned Reconciliation Matrix Viewer")
            st.markdown(
                "This matrix automatically groups, pairs, and tracks discrepancies side-by-side using unified **Dates** and **Consultant Profiles** to streamline validation workflows.")

            if not df_bl_display.empty or not df_ts_display.empty:
                df_bl_group = df_bl_display.copy()
                rev_bl_map = {v["beeline"]: k for k, v in st.session_state.mapping_dict.items() if v["beeline"]}
                df_bl_group["_AlignedName"] = df_bl_group[st.session_state.col_bl_name].str.strip().map(
                    rev_bl_map).fillna(df_bl_group[st.session_state.col_bl_name].str.strip())
                df_bl_group["_Hours"] = pd.to_numeric(
                    df_bl_group[st.session_state.col_bl_hours].astype(str).str.replace(',', '.'),
                    errors="coerce").fillna(0)

                df_bl_agg = df_bl_group.groupby(
                    [st.session_state.col_bl_date, "_AlignedName", st.session_state.col_bl_project,
                     st.session_state.col_bl_desc], as_index=False)["_Hours"].sum()
                df_bl_agg.columns = ["Match_Date", "Match_Name", "Beeline_Project", "Beeline_Description",
                                     "Beeline_Hours"]
                df_bl_agg = df_bl_agg.sort_values(
                    by=["Match_Date", "Match_Name", "Beeline_Project", "Beeline_Description"]).reset_index(drop=True)
                df_bl_agg["_Group_Idx"] = df_bl_agg.groupby(["Match_Date", "Match_Name"]).cumcount()

                df_ts_group = df_ts_display.copy()
                rev_ts_map = {v["timesheet"]: k for k, v in st.session_state.mapping_dict.items() if v["timesheet"]}
                df_ts_group["_AlignedName"] = df_ts_group[st.session_state.col_ts_name].str.strip().map(
                    rev_ts_map).fillna(df_ts_group[st.session_state.col_ts_name].str.strip())
                df_ts_group["_Hours"] = pd.to_numeric(
                    df_ts_group[st.session_state.col_ts_hours].astype(str).str.replace(',', '.'),
                    errors="coerce").fillna(0)

                df_ts_agg = df_ts_group.groupby(
                    [st.session_state.col_ts_date, "_AlignedName", st.session_state.col_ts_project,
                     st.session_state.col_ts_desc], as_index=False)["_Hours"].sum()
                df_ts_agg.columns = ["Match_Date", "Match_Name", "Timesheet_Project", "Timesheet_Description",
                                     "Timesheet_Hours"]
                df_ts_agg = df_ts_agg.sort_values(
                    by=["Match_Date", "Match_Name", "Timesheet_Project", "Timesheet_Description"]).reset_index(
                    drop=True)
                df_ts_agg["_Group_Idx"] = df_ts_agg.groupby(["Match_Date", "Match_Name"]).cumcount()

                df_recon = pd.merge(df_bl_agg, df_ts_agg, on=["Match_Date", "Match_Name", "_Group_Idx"], how="outer")
                df_recon["Beeline_Hours"] = df_recon["Beeline_Hours"].fillna(0)
                df_recon["Timesheet_Hours"] = df_recon["Timesheet_Hours"].fillna(0)
                df_recon["Beeline_Project"] = df_recon["Beeline_Project"].fillna("-")
                df_recon["Beeline_Description"] = df_recon["Beeline_Description"].fillna("-")
                df_recon["Timesheet_Project"] = df_recon["Timesheet_Project"].fillna("-")
                df_recon["Timesheet_Description"] = df_recon["Timesheet_Description"].fillna("-")
                df_recon = df_recon.sort_values(by=["Match_Date", "Match_Name", "_Group_Idx"]).reset_index(drop=True)

                df_recon["Variance_Hours"] = round(df_recon["Beeline_Hours"] - df_recon["Timesheet_Hours"], 2)
                df_recon = df_recon[
                    ["Match_Date", "Match_Name", "Beeline_Project", "Beeline_Description", "Beeline_Hours",
                     "Timesheet_Project", "Timesheet_Description", "Timesheet_Hours", "Variance_Hours"]]
                df_recon.columns = ["Reconciliation Date", "Consultant Profile", "Beeline Project",
                                    "Beeline Description", "Beeline Hours Total", "Timesheet Project",
                                    "Timesheet Description", "Timesheet Hours Total", "Variance Discrepancy"]


                def color_rows(row):
                    b_hrs = float(row["Beeline Hours Total"])
                    t_hrs = float(row["Timesheet Hours Total"])
                    if b_hrs == 0 or t_hrs == 0:
                        return ['background-color: #FFCDD2; color: #B71C1C; font-weight: 500;'] * len(row)
                    elif b_hrs == t_hrs:
                        return ['background-color: #C8E6C9; color: #1B5E20;'] * len(row)
                    else:
                        return ['background-color: #FFE0B2; color: #E65100; font-weight: 500;'] * len(row)


                styled_recon = df_recon.style.apply(color_rows, axis=1).format(
                    {"Beeline Hours Total": "{:.2f} h", "Timesheet Hours Total": "{:.2f} h",
                     "Variance Discrepancy": "{:+.2f} h"})

                st.dataframe(styled_recon, use_container_width=True, height=450,
                             column_config={"Beeline Project": st.column_config.TextColumn(width="medium"),
                                            "Beeline Description": st.column_config.TextColumn(width="large"),
                                            "Timesheet Project": st.column_config.TextColumn(width="medium"),
                                            "Timesheet Description": st.column_config.TextColumn(width="large")})
                st.caption(
                    "🎨 **Color Key Legend:** Green = Perfect Balance | Orange = Divergent Hours Discrepancy | Red = Missing Corresponding Entry (No Match Found)")

                # NOUVEAU: Bouton d'export pour la Matrice
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 Download Matrix as Excel",
                    data=convert_df_to_excel(df_recon),
                    file_name=f"Reconciliation_Matrix_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_matrix"
                )

            else:
                st.info("No data available to build the reconciliation matrix.")

            # =========================================================
            # HIDEABLE HIERARCHICAL RAW STREAM SECTION
            # =========================================================
            st.markdown("---")
            with st.expander("📁 Inspect Separated Raw Data Streams (Beeline / Timesheet Exports)", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("🐝 Beeline Raw Data Stream")
                    if not df_bl_display.empty:
                        st.dataframe(df_bl_display, width="stretch", height=400)
                        st.success(f"✅ Displayed Rows: {len(df_bl_display)}")

                        # NOUVEAU: Bouton d'export pour Beeline
                        st.download_button(
                            label="📥 Download Beeline Data",
                            data=convert_df_to_excel(df_bl_display),
                            file_name=f"Beeline_Raw_Data_{date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_beeline"
                        )
                    else:
                        st.info("No unfiltered Beeline stream data rows to show.")

                with col2:
                    st.subheader("📅 Timesheet Raw Data Stream")
                    if not df_ts_display.empty:
                        st.dataframe(df_ts_display, width="stretch", height=400)
                        st.success(f"✅ Displayed Rows: {len(df_ts_display)}")

                        # NOUVEAU: Bouton d'export pour Timesheet
                        st.download_button(
                            label="📥 Download Timesheet Data",
                            data=convert_df_to_excel(df_ts_display),
                            file_name=f"Timesheet_Raw_Data_{date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_timesheet"
                        )
                    else:
                        st.info("No unfiltered Timesheet stream data rows to show.")