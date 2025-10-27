import streamlit as st
import pandas as pd
import re
from io import BytesIO

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
    n = nums[-1].replace(',', '')
    try:
        return float(n)
    except:
        return None


def clean_excel(filepath):
    """Clean Excel by detecting header row and extracting numeric fields."""
    excel_raw = pd.read_excel(filepath, header=None)
    excel_raw = excel_raw.dropna(how="all")

    header_candidates = excel_raw[excel_raw.astype(str).apply(
        lambda r: r.str.contains("project_code", case=False, na=False)
    ).any(axis=1)]
    header_row = header_candidates.index[0] if not header_candidates.empty else 0

    df = pd.read_excel(filepath, header=header_row, sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()
    df["project_code"] = df["project_code"].astype(str).str.strip().str.upper()

    cols = ["score_1_response", "score_2_response", "score_1_average", "score_2_average"]
    for c in cols:
        if c in df.columns:
            df[c + "_num"] = df[c].apply(extract_number_keep_none)
        else:
            df[c + "_num"] = pd.NA

    return df


# ======================================================
# 📊 Static Data Loader
# ======================================================

@st.cache_data
def load_static_source():
    """Load static Excel file 'source.xlsx' automatically."""
    try:
        # Load all sheets at once
        xl = pd.ExcelFile("source.xlsx")
        all_sheets = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}

        # Extract the first sheet for main dashboard logic
        df_main = clean_excel("source.xlsx")
        st.success("✅ Source file 'source.xlsx' loaded successfully with multiple sheets.")
        return df_main, all_sheets
    except FileNotFoundError:
        st.error("⚠️ Source file 'source.xlsx' not found. Please add it to the app folder.")
        return pd.DataFrame(), {}
    except Exception as e:
        st.error(f"⚠️ Error reading 'source.xlsx': {e}")
        return pd.DataFrame(), {}


# ======================================================
# 🧾 Excel-Style Calculations
# ======================================================

def excel_style_calculations(df):
    if df.empty:
        return None

    sum_score_1_response = df["score_1_response_num"].fillna(0).sum()
    sum_score_2_response = df["score_2_response_num"].fillna(0).sum()
    avg_score_1_average = df["score_1_average_num"].dropna().mean() or 0.0
    avg_score_2_average = df["score_2_average_num"].dropna().mean() or 0.0

    score1_averageXsum = avg_score_1_average * sum_score_1_response
    score2_averageXsum = avg_score_2_average * sum_score_2_response
    total_responses = sum_score_1_response + sum_score_2_response
    total_weight = score1_averageXsum + score2_averageXsum
    total_score = total_weight / total_responses if total_responses else 0.0

    return {
        "Total Responses": int(total_responses),
        "Priority / Timeliness / Measures of Sustainability": round(avg_score_1_average, 2),
        "Sufficiency / Quality / Current Status": round(avg_score_2_average, 2),
        "Score1_averageXsum": int(round(score1_averageXsum)),
        "Score2_averageXsum": int(round(score2_averageXsum)),
        "score1weight+score2weight": int(round(total_weight)),
        "Total Quant Score - Relevance / Efficiency / Sustainability": round(total_score, 2),
    }


# ======================================================
# 🖥️ Streamlit UI — Quant Scoring Reckoner
# ======================================================

st.title("📊 Quant Scoring Reckoner Dashboard")
st.caption("Centralized Quantitative Scoring and Evaluation Framework")

# --- Infographic ---
st.markdown("""
<style>
.infotable {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
  text-align: center;
}
.infotable th {
  background-color: #1b263b;
  color: white;
  padding: 8px;
  border: 1px solid #555;
}
.infotable td {
  border: 1px solid #ccc;
  padding: 6px;
}
.infotable .col1 { background-color: #a7c7e7; font-weight: bold; }
.infotable .score1 { background-color: #d5e8d4; }
.infotable .score2 { background-color: #f8cecc; }
.infotable .param1 { background-color: #fff2cc; }
.infotable .param2 { background-color: #fff2cc; }
.infotable .remark { background-color: #f4cccc; }
</style>

<h4>🧭 Scoring Reckoner Quant — CSR Evaluation Framework</h4>

<table class="infotable">
<thead>
<tr>
  <th>Indicator</th>
  <th>Score 1</th>
  <th>Score 2</th>
  <th>Parameter 1</th>
  <th>Parameter 2</th>
  <th>Remark</th>
</tr>
</thead>
<tbody>
<tr><td class="col1">Relevance</td><td class="score1">Priority</td><td class="score2">Sufficiency</td><td class="param1">Beneficiary Need Alignment</td><td></td><td class="remark">Average of score 1 and score 2</td></tr>
<tr><td class="col1">Efficiency</td><td class="score1">Timeliness</td><td class="score2">Quality</td><td class="param1">Timeliness</td><td class="param2">Quality</td><td></td></tr>
<tr><td class="col1">Effectiveness</td><td class="score1">Short Term Result</td><td></td><td class="param1">Short Term Results</td><td></td><td></td></tr>
<tr><td class="col1">Sustainability</td><td class="score1">Measures of Sustainability</td><td class="score2">Current Status</td><td class="param1">Sustainability</td><td></td><td class="remark">Average of score 1 and score 2</td></tr>
<tr><td class="col1">Impact</td><td class="score1">Long-term Outcome</td><td></td><td class="param1">Impact</td><td></td><td></td></tr>
</tbody>
</table>
<hr>
""", unsafe_allow_html=True)

# --- Load Data ---
df, all_sheets = load_static_source()
if df.empty:
    st.stop()

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filters")

filter_cols = ["project_code", "Tool_level4", "score_type", "Intervention_level3", "Activity_level1"]
filtered_df = df.copy()

for col in filter_cols:
    if col in df.columns:
        options = sorted(df[col].dropna().unique().tolist())
        if col == "score_type":
            options = sorted(set(options + ["Effectiveness", "Impact"]))
        selected = st.sidebar.multiselect(f"{col}", options, default=options)
        filtered_df = filtered_df[filtered_df[col].isin(selected)]

if st.sidebar.button("🧹 Clear All Filters"):
    st.experimental_rerun()

# --- KPI Calculations ---
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

    st.markdown("### 🧮 Total Quant Score - Relevance / Efficiency / Sustainability")
    st.metric("Total Quant Score", f"{calc['Total Quant Score - Relevance / Efficiency / Sustainability']:.2f}")

    st.divider()
    st.subheader("⬇️ Download Data")

    raw_buf = BytesIO()
    filtered_df.to_excel(raw_buf, index=False)
    st.download_button(
        "📥 Download Filtered Raw Data",
        data=raw_buf.getvalue(),
        file_name="Filtered_Raw_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # --- New: Download Full Source File (All Sheets) ---
    source_buf = BytesIO()
    with pd.ExcelWriter(source_buf, engine="openpyxl") as writer:
        for sheet_name, sheet_data in all_sheets.items():
            sheet_data.to_excel(writer, index=False, sheet_name=sheet_name)
    st.download_button(
        "📦 Download Full Source File (All Sheets)",
        data=source_buf.getvalue(),
        file_name="Full_Source_File.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    summary_df = pd.DataFrame([calc])
    calc_buf = BytesIO()
    summary_df.to_excel(calc_buf, index=False)
    st.download_button(
        "📊 Download Calculated Summary",
        data=calc_buf.getvalue(),
        file_name="Calculated_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.warning("⚠️ No data available for the selected filters.")

st.caption("Data auto-loaded from 'source.xlsx' (includes all sheets).")
