import streamlit as st
import pandas as pd
import re
from io import BytesIO
from zipfile import ZipFile
from openpyxl import load_workbook
import os

st.set_page_config(page_title="Quant Scoring Reckoner", layout="wide")

# ======================================================
# 🧮 Helper Functions
# ======================================================

def extract_number_keep_none(obj):
    if pd.isna(obj):
        return None
    s = str(obj).strip()
    nums = re.findall(r'-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+\.\d+|-?\d+', s)
    if not nums:
        return None
    try:
        return float(nums[-1].replace(',', ''))
    except:
        return None

def clean_excel(filepath, sheet_name):
    excel_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine="openpyxl")
    excel_raw = excel_raw.dropna(how="all")

    header_candidates = excel_raw[excel_raw.astype(str).apply(
        lambda r: r.str.contains("project_code", case=False, na=False)
    ).any(axis=1)]
    header_row = header_candidates.index[0] if not header_candidates.empty else 0

    df = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row, engine="openpyxl")
    df.columns = df.columns.astype(str).str.strip()

    if "project_code" in df.columns:
        df["project_code"] = df["project_code"].astype(str).str.strip().str.upper()

    for col in ["score_1_response", "score_2_response", "score_1_average", "score_2_average"]:
        if col in df.columns:
            df[col + "_num"] = df[col].apply(extract_number_keep_none)
        else:
            df[col + "_num"] = 0

    return df

# ======================================================
# 📊 Data Loaders with Optimized Resource Caching
# ======================================================

@st.cache_resource(ttl=3600)
def load_analysis_data():
    try:
        xl = pd.ExcelFile("source.xlsx", engine="openpyxl")
        sheet_names = xl.sheet_names

        if "indicator_analysis_table" in sheet_names:
            df_main = clean_excel("source.xlsx", "indicator_analysis_table")
        else:
            candidate = next((s for s in sheet_names if "indicator" in s.lower() and "analysis" in s.lower()), None)
            df_main = clean_excel("source.xlsx", candidate) if candidate else pd.DataFrame()

        if "short_term_long_term" in sheet_names:
            df_stlt = xl.parse("short_term_long_term")
        else:
            candidate = next((s for s in sheet_names if "short" in s.lower() and "long" in s.lower()), None)
            df_stlt = xl.parse(candidate) if candidate else pd.DataFrame()

        qual_candidate = next((
            s for s in sheet_names
            if s.strip().lower() in ("study closure", "study_closure")
               or ("study" in s.lower() and "closure" in s.lower())
        ), None)

        df_qual = pd.DataFrame()
        if qual_candidate:
            df_qual = xl.parse(qual_candidate)
            df_qual.columns = df_qual.columns.astype(str).str.strip()

        return df_main, df_stlt, df_qual, sheet_names

    except FileNotFoundError:
        st.error("⚠️ Source file 'source.xlsx' not found. Please add it to the app folder.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), []

    except Exception as e:
        st.error(f"⚠️ Error loading 'source.xlsx': {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), []

# ======================================================
# 🧾 KPI Calculations
# ======================================================

def excel_style_calculations(df):
    if df.empty:
        return None

    s1r, s2r = df["score_1_response_num"].sum(), df["score_2_response_num"].sum()
    a1, a2 = df["score_1_average_num"].mean(), df["score_2_average_num"].mean()
    s1w, s2w = a1 * s1r, a2 * s2r
    total_responses = s1r + s2r
    total_weight = s1w + s2w
    total_score = total_weight / total_responses if total_responses else 0

    return {
        "Total Responses": int(total_responses),
        "Priority / Timeliness / Measures of Sustainability": round(a1, 2),
        "Sufficiency / Quality / Current Status": round(a2, 2),
        "Score1_averageXsum": int(round(s1w)),
        "Score2_averageXsum": int(round(s2w)),
        "score1weight+score2weight": int(round(total_weight)),
        "Total Quant Score - Relevance / Efficiency / Sustainability": round(total_score, 2),
    }

# ======================================================
# 🖥️ Streamlit Dashboard (MULTI-PAGE)
# ======================================================

page = st.sidebar.selectbox("Pages", ["Quant Scoring Reckoner", "Short Term Long Term", "Qualitative Score"])

df_main, df_stlt, df_qual, sheet_names = load_analysis_data()

# --- Page: Quant Scoring Reckoner ---
if page == "Quant Scoring Reckoner":
    st.title("📊 Quant Scoring + Reckoner Dashboard")
    st.caption("Optimized, Modular CSR Evaluation Framework")

    if df_main.empty:
        st.info("Indicator analysis data not found.")
        st.stop()

    st.sidebar.header("🔍 Filters")

    filter_cols = ["project_code", "Tool_level4", "score_type", "Intervention_level3", "Activity_level1"]
    filtered_df = df_main.copy()

    for col in filter_cols:
        if col in df_main.columns:
            options = sorted(df_main[col].dropna().unique().tolist())
            if col == "score_type":
                options = [opt for opt in options if opt not in ["Impact", "Effectiveness"]]
            selected = st.sidebar.multiselect(col, options, default=options)
            if len(selected) != len(options):
                filtered_df = filtered_df[filtered_df[col].isin(selected)]

    if st.sidebar.button("🧹 Clear All Filters"):
        st.rerun()

    calc = excel_style_calculations(filtered_df)
    if calc:
        st.markdown("## 📈 Key Performance Indicators (KPIs)")
        kpi_cols = st.columns(3)
        kpi_cols[0].metric("Total Responses", f"{calc['Total Responses']:,}")
        kpi_cols[1].metric("Priority / Timeliness / Measures of Sustainability", f"{calc['Priority / Timeliness / Measures of Sustainability']:.2f}")
        kpi_cols[2].metric("Sufficiency / Quality / Current Status", f"{calc['Sufficiency / Quality / Current Status']:.2f}")

        st.divider()
        kpi_cols2 = st.columns(3)
        kpi_cols2[0].metric("Score1_averageXsum", f"{calc['Score1_averageXsum']:,}")
        kpi_cols2[1].metric("Score2_averageXsum", f"{calc['Score2_averageXsum']:,}")
        kpi_cols2[2].metric("score1weight+score2weight", f"{calc['score1weight+score2weight']:,}")

        st.metric("Total Quant Score", f"{calc['Total Quant Score - Relevance / Efficiency / Sustainability']:.2f}")

    st.divider()
    st.subheader("⬇️ Download Data")

    buf = BytesIO()
    filtered_df.to_excel(buf, index=False)
    st.download_button("📥 Download Filtered Raw Data",
                       data=buf.getvalue(),
                       file_name="Filtered_Raw_Data.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    def generate_data_package():
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w") as zf:
            out_excel = BytesIO()
            with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
                df_main.to_excel(writer, index=False, sheet_name="indicator_analysis_table")
                df_stlt.to_excel(writer, index=False, sheet_name="short_term_long_term")
            zf.writestr("source_data.xlsx", out_excel.getvalue())

            possible_folders = ["static_data", "static data"]
            csv_folder = next((f for f in possible_folders if os.path.exists(f)), None)

            if csv_folder:
                for csv_file in os.listdir(csv_folder):
                    if csv_file.lower().endswith(".csv"):
                        file_path = os.path.join(csv_folder, csv_file)
                        with open(file_path, "rb") as f:
                            zf.writestr(f"{csv_folder}/{csv_file}", f.read())
            else:
                st.warning("⚠️ CSV folder not found. Ensure it's named 'static_data' or 'static data'.")

        zip_buffer.seek(0)
        return zip_buffer.getvalue()

    st.download_button(
        "📦 Download All Data (ZIP)",
        data=generate_data_package(),
        file_name="Complete_Data_Package.zip",
        mime="application/zip",
    )

# --- Page: Short Term Long Term ---
elif page == "Short Term Long Term":
    st.title("📋 Short Term / Long Term")
    if df_stlt.empty:
        st.info("No 'short_term_long_term' sheet found.")
    else:
        # Normalize column names
        df_stlt.columns = df_stlt.columns.astype(str).str.strip()

        # Coerce (if present) to numeric
        if "response_count" in df_stlt.columns:
            df_stlt["response_count"] = pd.to_numeric(df_stlt["response_count"], errors="coerce")
        else:
            df_stlt["response_count"] = pd.NA

        if "average" in df_stlt.columns:
            df_stlt["average"] = pd.to_numeric(df_stlt["average"], errors="coerce")
        else:
            df_stlt["average"] = pd.NA

        # Sum of response_count (ignore NaN)
        total_response_count = int(df_stlt["response_count"].dropna().sum()) if df_stlt["response_count"].dropna().size else 0

        # Mean of the 'average' column (ignore NaN)
        avg_of_avg = float(df_stlt["average"].dropna().mean()) if df_stlt["average"].dropna().size else None

        # Display metrics (weighted average removed)
        col1, col2 = st.columns([1, 1])
        col1.metric("Total Response Count (sum)", f"{total_response_count:,}")
        if avg_of_avg is not None:
            col2.metric("Mean of 'average' column", f"{avg_of_avg:.2f}")
        else:
            col2.info("Mean of 'average': N/A")

        st.dataframe(df_stlt.reset_index(drop=True), use_container_width=True)

        buf2 = BytesIO()
        df_stlt.to_excel(buf2, index=False)
        st.download_button("📥 Download Short Term Long Term",
                           data=buf2.getvalue(),
                           file_name="short_term_long_term.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Page: Qualitative Score ---
else:
    st.title("📝 Qualitative Score")

    if df_qual.empty:
        st.info("No 'study closure' sheet found in source.xlsx.")
    else:
        def find_col(df, *candidates_lower):
            for cand in candidates_lower:
                for actual in df.columns:
                    if actual.strip().lower() == cand.strip().lower():
                        return actual
                for actual in df.columns:
                    if cand.strip().lower() in actual.strip().lower():
                        return actual
            return None

        col_project = find_col(df_qual, "data.project_code", "project_code", "data.projectcode", "project")
        col_value = find_col(df_qual, "value", "Value", "val")
        col_tool = find_col(df_qual, "tool", "tool_name")
        col_indicator = find_col(df_qual, "indicator", "indicator_name")
        col_parameter = find_col(df_qual, "parameter", "parameter_name")
        col_intervention = find_col(df_qual, "intervention", "intervention_name")

        df_qual[col_value] = pd.to_numeric(df_qual[col_value], errors="coerce")

        qual_filtered = df_qual.copy()
        filter_cols_qual = []
        if col_project: filter_cols_qual.append((col_project, "data.project_code"))
        if col_tool: filter_cols_qual.append((col_tool, "tool"))
        if col_indicator: filter_cols_qual.append((col_indicator, "indicator"))
        if col_parameter: filter_cols_qual.append((col_parameter, "parameter"))
        if col_intervention: filter_cols_qual.append((col_intervention, "intervention"))

        st.sidebar.subheader("Qualitative Filters")

        for actual_col, label in filter_cols_qual:
            options = sorted(qual_filtered[actual_col].dropna().unique().tolist())
            selected = st.sidebar.multiselect(f"{label}", options, default=options)
            if len(selected) != len(options):
                qual_filtered = qual_filtered[qual_filtered[actual_col].isin(selected)]

        total_responses_qual = len(qual_filtered)
        avg_value = float(qual_filtered[col_value].mean()) if not qual_filtered.empty else None

        col1, col2 = st.columns(2)
        col1.metric("Total Responses (Qualitative)", f"{total_responses_qual:,}")
        if avg_value is not None:
            col2.metric("Average Value", f"{avg_value:.2f}")
        else:
            col2.info("Average Value: N/A")

        st.dataframe(qual_filtered.reset_index(drop=True), use_container_width=True)

        qual_buf = BytesIO()
        qual_filtered.to_excel(qual_buf, index=False)
        st.download_button("📥 Download Filtered Qualitative Data",
                           data=qual_buf.getvalue(),
                           file_name="Filtered_Qualitative_Data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.caption("If you updated 'source.xlsx', refresh the app.")
