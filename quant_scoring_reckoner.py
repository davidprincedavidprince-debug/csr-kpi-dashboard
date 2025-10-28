import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook

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
    """Read only the required sheet efficiently and clean it."""
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
# 📊 Data Loader (Lightweight)
# ======================================================

@st.cache_data(show_spinner=False)
def load_analysis_data():
    try:
        df = clean_excel("source.xlsx", "indicator_analysis_table")
        wb = load_workbook("source.xlsx", read_only=True)
        all_sheets = wb.sheetnames
        wb.close()
        st.success(f"✅ Loaded sheet 'indicator_analysis_table' ({len(df):,} rows).")
        return df, all_sheets
    except FileNotFoundError:
        st.error("⚠️ 'source.xlsx' not found.")
        return pd.DataFrame(), []
    except Exception as e:
        st.error(f"⚠️ Error loading Excel: {e}")
        return pd.DataFrame(), []


# ======================================================
# 🧾 KPI Calculation
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
# 🖥️ Streamlit UI
# ======================================================

st.title("📊 Quant Scoring Reckoner Dashboard")
st.caption("Centralized Quantitative Scoring and Evaluation Framework")

# --- Load Main Data ---
df, sheet_names = load_analysis_data()
if df.empty:
    st.stop()

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filters")

filter_cols = ["project_code", "Tool_level4", "score_type", "Intervention_level3", "Activity_level1"]
filtered_df = df.copy()

for col in filter_cols:
    if col in df.columns:
        options = sorted(df[col].dropna().unique().tolist())
        # Remove Impact & Effectiveness
        if col == "score_type":
            options = [opt for opt in options if opt not in ["Impact", "Effectiveness"]]
        selected = st.sidebar.multiselect(col, options, default=options)
        if len(selected) != len(options):
            filtered_df = filtered_df[filtered_df[col].isin(selected)]

if st.sidebar.button("🧹 Clear All Filters"):
    st.experimental_rerun()

# --- KPIs ---
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

    # Filtered Data
    buf = BytesIO()
    filtered_df.to_excel(buf, index=False)
    st.download_button("📥 Download Filtered Raw Data",
                       data=buf.getvalue(),
                       file_name="Filtered_Raw_Data.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Lazy Full Source Download
    def generate_full_source():
        wb = load_workbook("source.xlsx", read_only=True)
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for s in wb.sheetnames:
                df_temp = pd.read_excel("source.xlsx", sheet_name=s, engine="openpyxl")
                df_temp.to_excel(writer, index=False, sheet_name=s)
        wb.close()
        return out.getvalue()

    st.download_button(
        "📦 Download Full Source File (All Sheets — including POE, SDLE, NRM, Health & Hygiene)",
        data=generate_full_source(),
        file_name="Full_Source_File.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.warning("⚠️ No data available for the selected filters.")

st.caption("Data auto-loaded from 'source.xlsx' (includes all sheets dynamically).")
