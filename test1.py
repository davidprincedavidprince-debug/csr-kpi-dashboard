import pandas as pd
import re
import streamlit as st

# ============================================================
# 🔧 HELPER FUNCTIONS
# ============================================================

def extract_number_keep_none(obj):
    """Extract numeric value from messy Excel cell, return None if invalid."""
    if pd.isna(obj):
        return None
    s = str(obj).strip()
    nums = re.findall(r'-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+\.\d+|-?\d+', s)
    if not nums:
        return None
    n = nums[-1].replace(',', '')
    try:
        return float(n)
    except:
        return None


def clean_excel(filepath):
    """Read and clean the Excel file, handling blank rows before headers."""
    excel_raw = pd.read_excel(filepath, header=None)
    excel_raw = excel_raw.dropna(how='all')
    header_candidates = excel_raw[excel_raw.astype(str).apply(
        lambda r: r.str.contains('project_code', case=False, na=False)
    ).any(axis=1)]
    header_row = header_candidates.index[0] if not header_candidates.empty else 0

    df = pd.read_excel(filepath, header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    df['project_code'] = df['project_code'].astype(str).str.strip().str.upper()

    cols = ['score_1_response', 'score_2_response', 'score_1_average', 'score_2_average']
    for c in cols:
        if c in df.columns:
            df[c + '_num'] = df[c].apply(extract_number_keep_none)
        else:
            df[c + '_num'] = pd.NA

    return df


# ============================================================
# 📊 STATIC SOURCE DATA LOADER
# ============================================================

@st.cache_data
def load_static_source():
    """Load the static Excel file 'source.xlsx' from the same folder."""
    try:
        df = clean_excel("source.xlsx")  # static file path
        st.success("✅ Source file 'source.xlsx' loaded successfully.")
        return df
    except FileNotFoundError:
        st.error("⚠️ Source file 'source.xlsx' not found in the app folder.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"⚠️ Error reading 'source.xlsx': {e}")
        return pd.DataFrame()


# ============================================================
# 🧮 CALCULATION FUNCTION
# ============================================================

def calculate_project_metrics(df, filters=None):
    if df.empty:
        return pd.DataFrame()

    if filters:
        for key, values in filters.items():
            if values and key in df.columns:
                df = df[df[key].isin(values)]

    projects = sorted(df['project_code'].dropna().unique().tolist())
    results = []

    for project in projects:
        dproj = df[df['project_code'] == project].copy()
        row_count = len(dproj)

        sum_score_1_response = dproj['score_1_response_num'].fillna(0).sum()
        sum_score_2_response = dproj['score_2_response_num'].fillna(0).sum()
        avg_score_1_average = dproj['score_1_average_num'].dropna().mean()
        avg_score_2_average = dproj['score_2_average_num'].dropna().mean()

        score1_averageXsum = (avg_score_1_average or 0.0) * sum_score_1_response
        score2_averageXsum = (avg_score_2_average or 0.0) * sum_score_2_response
        total_responses = sum_score_1_response + sum_score_2_response
        total_weight = score1_averageXsum + score2_averageXsum
        total_score = total_weight / total_responses if total_responses else 0.0

        results.append({
            'project_code': project,
            'Row_Count': row_count,
            'Total_Responses': int(total_responses),
            'Priority/Timeliness /Measures of Sustainability': round(avg_score_1_average, 2),
            'Sufficiency/Quality/Current Status': round(avg_score_2_average, 2),
            'Score1_averageXsum': round(score1_averageXsum),
            'Score2_averageXsum': round(score2_averageXsum),
            'score1weight+score2weight': round(total_weight),
            'Total Quant Score - Relevance /Effeciency /Sustainability': round(total_score, 2)
        })

    return pd.DataFrame(results)


# ============================================================
# 🖥️ STREAMLIT DASHBOARD
# ============================================================

st.set_page_config(page_title="CSR KPI Dashboard", layout="wide")

st.title("📊 CSR KPI Dashboard")
st.markdown("#### Scoring Reckoner Quant Dashboard")

# Load the static data
df = load_static_source()

if df.empty:
    st.stop()

# --- Filter Section ---
st.sidebar.header("🔍 Filters")

filter_fields = {
    "project_code": st.sidebar.multiselect("Project Code", sorted(df['project_code'].unique())),
    "Tool_level4": st.sidebar.multiselect("Tool Level 4", sorted(df['Tool_level4'].dropna().unique())),
    "score_type": st.sidebar.multiselect("Score Type", sorted(df['score_type'].dropna().unique())),
    "Intervention_level3": st.sidebar.multiselect("Intervention Level 3", sorted(df['Intervention_level3'].dropna().unique())),
    "Activity_level1": st.sidebar.multiselect("Activity Level 1", sorted(df['Activity_level1'].dropna().unique()))
}

# Clear Filters button
if st.sidebar.button("🧹 Clear All Filters"):
    st.experimental_rerun()

# --- Calculations ---
summary_df = calculate_project_metrics(df, filters=filter_fields)

# --- Display Results ---
st.dataframe(summary_df, use_container_width=True)

# --- Download Option ---
csv = summary_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Results as CSV", csv, "CSR_KPI_Summary.csv", "text/csv")

st.caption("Data auto-loaded from static file 'source.xlsx'.")
