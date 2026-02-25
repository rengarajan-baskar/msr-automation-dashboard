import streamlit as st
import os
import pandas as pd
import plotly.express as px
from msr_automator import normalize_columns, coalesce_root_cause, build_pivot, safe_get_col, load_config

st.set_page_config(page_title="MSR Automation", layout="wide")

st.title("📊 Monthly Status Review (MSR) Automation")
st.caption("Upload your Excel tracker and generate summarized MSR reports automatically.")

# --- Load Configuration ---
config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
cfg = load_config(config_path)
aliases = cfg.get("column_aliases", {})
rc_rules = cfg.get("root_cause_rules", {})
out_cfg = cfg.get("output", {})

uploaded_file = st.file_uploader("Upload your Excel Tracker (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.info("Processing file... please wait ⏳")

    # --- Read Excel ---
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df = normalize_columns(df, aliases)

    # --- Convert date columns ---
    for date_col in ["opened_at", "resolved_at"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    # --- Determine root causes ---
    rc_col = safe_get_col(df, "root_cause")
    df["RootCauseFinal"] = df.apply(lambda r: coalesce_root_cause(r, rc_col, rc_rules), axis=1)

    # --- Compute MTTR ---
    if "opened_at" in df.columns and "resolved_at" in df.columns:
        df["ResolutionTimeHours"] = (df["resolved_at"] - df["opened_at"]).dt.total_seconds() / 3600
    else:
        df["ResolutionTimeHours"] = None

    # --- FILTERS SECTION ---
    st.markdown("### 🔍 Filters")

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        selected_type = st.multiselect(
            "Filter by Ticket Type",
            sorted(df["type"].dropna().unique()),
            default=sorted(df["type"].dropna().unique())
        )

    with col2:
        selected_priority = st.multiselect(
            "Filter by Priority",
            sorted(df["priority"].dropna().unique()),
            default=sorted(df["priority"].dropna().unique())
        )

    with col3:
        selected_state = st.multiselect(
            "Filter by State",
            sorted(df["state"].dropna().unique()),
            default=sorted(df["state"].dropna().unique())
        )

    with col4:
        if "opened_at" in df.columns:
            min_open, max_open = df["opened_at"].min(), df["opened_at"].max()
            date_range_open = st.date_input(
                "Opened Date Range",
                value=(min_open, max_open),
                min_value=min_open,
                max_value=max_open
            )
        else:
            date_range_open = None

    with col5:
        if "resolved_at" in df.columns:
            min_res, max_res = df["resolved_at"].min(), df["resolved_at"].max()
            date_range_res = st.date_input(
                "Resolved Date Range",
                value=(min_res, max_res),
                min_value=min_res,
                max_value=max_res
            )
        else:
            date_range_res = None

    # --- Apply filters ---
    df_filtered = df.copy()
    df_filtered = df_filtered[
        (df_filtered["type"].isin(selected_type)) &
        (df_filtered["priority"].isin(selected_priority)) &
        (df_filtered["state"].isin(selected_state))
    ]

    if date_range_open and "opened_at" in df.columns:
        start_open, end_open = pd.to_datetime(date_range_open[0]), pd.to_datetime(date_range_open[1])
        df_filtered = df_filtered[
            (df_filtered["opened_at"] >= start_open) & (df_filtered["opened_at"] <= end_open)
        ]

    if date_range_res and "resolved_at" in df.columns:
        start_res, end_res = pd.to_datetime(date_range_res[0]), pd.to_datetime(date_range_res[1])
        df_filtered = df_filtered[
            (df_filtered["resolved_at"].isna()) |  # Include unresolved tickets
            ((df_filtered["resolved_at"] >= start_res) & (df_filtered["resolved_at"] <= end_res))
        ]

    st.success(f"Showing {len(df_filtered)} filtered records out of {len(df)} total tickets.")

    # --- Build pivot summaries ---
    wanted = [
        "type","priority","state","category","subcategory",
        "RootCauseFinal","assignment_group","customer","sla_breached"
    ]
    pivots = {col: build_pivot(df_filtered, col) for col in wanted}

    # --- Show summaries interactively ---
    tab_titles = ["Overview", "MTTR", "Trends"] + list(pivots.keys())
    tabs = st.tabs(tab_titles)

    with tabs[0]:
        st.metric(label="Total Tickets", value=len(df_filtered))
        st.dataframe(df_filtered.head(10), use_container_width=True)

    # --- MTTR Summary ---
    with tabs[1]:
        st.subheader("⏱ Mean Time To Resolve (MTTR)")
        if "ResolutionTimeHours" in df_filtered.columns and not df_filtered["ResolutionTimeHours"].isna().all():
            mttr_summary = (
                df_filtered.groupby("type")["ResolutionTimeHours"]
                .mean()
                .reset_index()
                .rename(columns={"ResolutionTimeHours": "Avg_Resolution_Hours"})
            )
            mttr_summary["Avg_Resolution_Days"] = (mttr_summary["Avg_Resolution_Hours"] / 24).round(2)
            st.dataframe(mttr_summary, use_container_width=True)

            chart_mttr = px.bar(
                mttr_summary,
                x="type",
                y="Avg_Resolution_Days",
                title="Average Resolution Time (Days) per Ticket Type",
                color="type",
                text_auto=True
            )
            st.plotly_chart(chart_mttr, use_container_width=True)
        else:
            st.warning("No resolved tickets found to calculate MTTR.")

    # --- Trends Tab ---
    with tabs[2]:
        st.subheader("📈 Monthly Ticket Trends")

        if "opened_at" in df_filtered.columns:
            df_filtered["MonthOpened"] = df_filtered["opened_at"].dt.to_period("M").astype(str)
            trend_opened = (
                df_filtered.groupby("MonthOpened").size().reset_index(name="Opened_Tickets")
            )
            chart_open = px.line(
                trend_opened,
                x="MonthOpened",
                y="Opened_Tickets",
                title="Tickets Opened Over Time",
                markers=True
            )
            st.plotly_chart(chart_open, use_container_width=True)

        if "resolved_at" in df_filtered.columns:
            df_filtered["MonthResolved"] = df_filtered["resolved_at"].dt.to_period("M").astype(str)
            trend_resolved = (
                df_filtered.groupby("MonthResolved").size().reset_index(name="Resolved_Tickets")
            )
            chart_res = px.line(
                trend_resolved,
                x="MonthResolved",
                y="Resolved_Tickets",
                title="Tickets Resolved Over Time",
                markers=True
            )
            st.plotly_chart(chart_res, use_container_width=True)

    # --- Remaining Tabs (Pivot Summaries) ---
    for i, (col, pv) in enumerate(pivots.items(), start=3):
        with tabs[i]:
            st.subheader(f"{col.title()} Summary")
            st.dataframe(pv, use_container_width=True)

    # --- Allow download ---
    out_dir = os.path.join(os.path.dirname(__file__), "out")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "MSR_Summary.xlsx")

    with pd.ExcelWriter(
        out_file,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}}
    ) as writer:
        for col, pv in pivots.items():
            pv.to_excel(writer, index=False, sheet_name=col[:31])
        df_filtered.to_excel(writer, index=False, sheet_name="Details")

    with open(out_file, "rb") as f:
        st.download_button(
            label="⬇️ Download Filtered MSR Summary",
            data=f,
            file_name="MSR_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ MSR Summary generated successfully!")

else:
    st.warning("Please upload your Excel file to begin.")
