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
    """Extract numeric value from messy Excel cell."""
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
    """Read specific Excel sheet efficiently."""
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
    """
    Load source.xlsx efficiently with 1-hour caching.
    This prevents reloading large Excel files on every rerun.
    """
    try:
        df_main = clean_excel("source.xlsx", "indicator_analysis_table")
        df_stlt = pd.read_excel("source.xlsx", sheet_name="short_term_long_term", engine="openpyxl")
        st.success(f"✅ Loaded indicator_analysis_table ({len(df_main):,} rows) and short_term_long_term.")
        return df_main, df_stlt
    except Exception as e:
        st.error(f"⚠️ Error loading 'source.xlsx': {e}")
        return pd.DataFrame(), pd.DataFrame()


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
# 🖥️ Streamlit Dashboard
# ======================================================

st.title("📊 Quant Scoring Reckoner Dashboard")
st.caption("Optimized, Modular CSR Evaluation Framework")

df_main, df_stlt = load_analysis_data()
if df_main.empty:
    st.stop()

# --- Sidebar Filters ---
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
    st.rerun()  # ✅ updated API

# --- KPI Calculations ---
calc = excel_style_calculations(filtered_df)
if calc:
    st.markdown("## 📈 Key Performance Indicators (KPIs)")
    kpi_cols = st.columns(3)
    kpi_cols[0].metric("Total Responses", f"{calc['Total Responses']:,}")
    kpi_cols[1].metric("Priority / Timeliness / Measures of Sustainability",
                       f"{calc['Priority / Timeliness / Measures of Sustainability']:.2f}")
    kpi_cols[2].metric("Sufficiency / Quality / Current Status",
                       f"{calc['Sufficiency / Quality / Current Status']:.2f}")

    st.divider()
    kpi_cols2 = st.columns(3)
    kpi_cols2[0].metric("Score1_averageXsum", f"{calc['Score1_averageXsum']:,}")
    kpi_cols2[1].metric("Score2_averageXsum", f"{calc['Score2_averageXsum']:,}")
    kpi_cols2[2].metric("score1weight+score2weight", f"{calc['score1weight+score2weight']:,}")

    st.metric("Total Quant Score",
              f"{calc['Total Quant Score - Relevance / Efficiency / Sustainability']:.2f}")

# ======================================================
# 📦 Download Options
# ======================================================

st.divider()
st.subheader("⬇️ Download Data")

# Filtered Excel
buf = BytesIO()
filtered_df.to_excel(buf, index=False)
st.download_button("📥 Download Filtered Raw Data",
                   data=buf.getvalue(),
                   file_name="Filtered_Raw_Data.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Smart Combined Download (Excel + CSVs) ---
st.markdown("### 📦 Download Complete Dataset (Excel + CSVs)")
st.caption("Includes indicator_analysis_table, short_term_long_term, and all supporting CSVs.")

def generate_data_package():
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zf:
        # Add both Excel sheets
        out_excel = BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
            df_main.to_excel(writer, index=False, sheet_name="indicator_analysis_table")
            df_stlt.to_excel(writer, index=False, sheet_name="short_term_long_term")
        zf.writestr("source_data.xlsx", out_excel.getvalue())

        # Detect CSV folder automatically
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

st.caption("⚡ Optimized with 1-hour resource caching and automatic CSV folder detection.")
