import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Short-Term / Long-Term Dashboard", layout="wide")

st.title("🧠 Short-Term / Long-Term Dashboard")
st.caption("Independent analysis of effectiveness and impact data from short_term_long_term sheet.")

# ======================================================
# 📊 Data Loader
# ======================================================
@st.cache_data
def load_short_term_long_term():
    """Load 'short_term_long_term' sheet from source.xlsx"""
    try:
        df = pd.read_excel("source.xlsx", sheet_name="short_term_long_term")
        df.columns = df.columns.astype(str).str.strip()

        rename_map = {
            "Project Code": "project_code",
            "Tool (Level 4)": "Tool_level4",
            "Intervention (Level 3)": "Intervention_level3",
            "Score_type": "score_type",
            "average": "average",
            "response_count": "response_count"
        }
        df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

        # Normalize text fields
        for col in ["project_code", "Tool_level4", "Intervention_level3", "score_type"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.title()

        # Ensure numeric fields
        for col in ["average", "response_count"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        st.success("✅ short_term_long_term sheet loaded successfully.")
        return df
    except FileNotFoundError:
        st.error("⚠️ 'source.xlsx' file not found in the app folder.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"⚠️ Error loading data: {e}")
        return pd.DataFrame()


# ======================================================
# 🧮 KPI Calculations
# ======================================================
def calculate_short_term_metrics(df):
    if df.empty:
        return None

    total_responses = df["response_count"].fillna(0).sum()
    avg_score = df["average"].dropna().mean()
    return {
        "Total Responses": int(total_responses),
        "Average Score": round(avg_score, 2)
    }


# ======================================================
# 📈 Page Layout
# ======================================================

df = load_short_term_long_term()
if df.empty:
    st.stop()

st.sidebar.header("🔍 Filters")

# Sidebar Filters (Independent)
filter_cols = ["project_code", "score_type", "Tool_level4", "Intervention_level3"]
filtered_df = df.copy()

for col in filter_cols:
    if col in df.columns:
        options = sorted(df[col].dropna().unique().tolist())
        selected = st.sidebar.multiselect(f"{col}", options, default=options)
        filtered_df = filtered_df[filtered_df[col].isin(selected)]

# Clear filters
if st.sidebar.button("🧹 Clear All Filters"):
    st.experimental_rerun()

# --- KPIs ---
metrics = calculate_short_term_metrics(filtered_df)

if metrics:
    st.markdown("## 📊 Effectiveness & Impact Metrics")
    cols = st.columns(2)
    cols[0].metric("Total Responses", f"{metrics['Total Responses']:,}")
    cols[1].metric("Average Score", f"{metrics['Average Score']:.2f}")

    st.divider()
    st.markdown("### 🔍 Filtered Data View")
    st.dataframe(filtered_df, use_container_width=True)

    # --- Download Filtered Data ---
    buf = BytesIO()
    filtered_df.to_excel(buf, index=False)
    st.download_button(
        "⬇️ Download Filtered Data",
        data=buf.getvalue(),
        file_name="Short_Term_Long_Term_Standalone.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.caption(f"ℹ️ Total rows displayed: {len(filtered_df)} | Source: 'short_term_long_term' sheet")
else:
    st.warning("⚠️ No data available after applying filters.")
