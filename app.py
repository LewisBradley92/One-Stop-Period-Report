import datetime as dt
import io
import os

import pandas as pd
import streamlit as st

from report_builder import build_report, make_client_workbook, period_dates, filter_audits

st.set_page_config(page_title="One Stop Periodic Report Builder", page_icon="📊", layout="wide")

st.title("One Stop Periodic Test Purchase Report")
st.write(
    "Generate the One Stop report for a selected period. The app replaces the **This month** data, "
    "appends new audits to **YTD**, and preserves the workbook's existing column positions and formatting."
)

st.info(
    "Rules applied: eligible One Stop site IDs from the Master Schedule; only Alcohol, E-Cig and Cigarettes; "
    "the product must match that site's assignment for the selected period; aborted audits are excluded. "
    "One Stop periods run Sunday to Saturday."
)

with st.expander("Files needed each run", expanded=True):
    st.markdown(
        "- **Latest One Stop report workbook** — use the original workbook for the first run, then upload the most recently generated workbook each subsequent period so YTD history is retained.\n"
        "- **Audit export CSV** — must include `site_internal_id`, `internal_id`, `item_to_order`, and `date_of_visit`.\n"
        "- **Master Schedule** — the current list of eligible One Stop sites and each site's period/category assignment.\n"
        "- **Period calendar** — Tesco sister-company calendar; the app shifts period boundaries to Sunday–Saturday for One Stop."
    )

left, right = st.columns(2)
with left:
    template_file = st.file_uploader("1. Upload latest One Stop report workbook (.xlsx)", type=["xlsx"], key="template")
    csv_file = st.file_uploader("2. Upload audit export (.csv)", type=["csv"], key="audit")
with right:
    master_file = st.file_uploader("3. Upload One Stop Master Schedule (.xlsx)", type=["xlsx"], key="master")
    calendar_file = st.file_uploader("4. Upload period calendar (.xlsx)", type=["xlsx"], key="calendar")

if all([template_file, csv_file, master_file, calendar_file]):
    calendar_bytes = calendar_file.getvalue()
    # Parse available financial years from the calendar without assuming a fixed year range.
    try:
        cal_df = pd.read_excel(io.BytesIO(calendar_bytes), sheet_name=0, dtype=object)
        year_options = [str(c).strip() for c in cal_df.columns if len(str(c).strip()) == 5 and str(c).strip()[2] == "-"]
        if not year_options:
            year_options = ["26-27"]
    except Exception:
        year_options = ["26-27"]

    now = dt.date.today()
    default_fy = next((fy for fy in year_options if fy.startswith(str(now.year)[-2:])), year_options[0])
    if default_fy not in year_options:
        default_fy = year_options[0]

    st.subheader("Select reporting period")
    c1, c2 = st.columns([1, 2])
    with c1:
        financial_year = st.selectbox(
            "Financial year",
            options=year_options,
            index=year_options.index(default_fy),
        )
    with c2:
        period = st.selectbox("One Stop period", options=list(range(1, 13)), index=7, format_func=lambda p: f"Period {p}")

    try:
        start_date, end_date = period_dates(calendar_bytes, financial_year, period)
        st.caption(f"Selected One Stop period: **{start_date:%d %b %Y} – {end_date:%d %b %Y}** (inclusive)")
    except Exception as exc:
        st.error(f"Could not read the selected period dates: {exc}")
        st.stop()

    if st.button("Check allocation before building", type="secondary"):
        try:
            _, _, check = filter_audits(
                csv_file.getvalue(), master_file.getvalue(), calendar_bytes, financial_year, period
            )
            st.session_state["allocation_check"] = check
        except Exception as exc:
            st.error(f"Allocation check failed: {exc}")

    check = st.session_state.get("allocation_check")
    if check and check.get("period") == period and check.get("financial_year") == financial_year:
        st.subheader("Allocation check")
        metrics = st.columns(4)
        metrics[0].metric("Source rows", check["source_rows"])
        metrics[1].metric("Eligible rows", check["included_rows"])
        metrics[2].metric("Aborted rows in source", check["aborted_source_rows"])
        metrics[3].metric("Sites/category/period", check["scheduled_category_rows"])
        st.write(
            f"Selected period: {check['start_date']:%d/%m/%Y} to {check['end_date']:%d/%m/%Y}. "
            f"Rows from listed sites and allowed categories before period-category matching: "
            f"{check['listed_site_rows']:,}; rows matching the selected period's site assignments: "
            f"{check['scheduled_category_rows']:,}."
        )
        if check["category_counts"]:
            st.write("Included audit categories:")
            st.dataframe(
                pd.DataFrame(
                    [{"Category": category, "Audits": count} for category, count in check["category_counts"].items()]
                ),
                hide_index=True,
                use_container_width=True,
            )

    if st.button("Build and download One Stop report", type="primary"):
        try:
            output_bytes, summary = build_report(
                template_file.getvalue(),
                csv_file.getvalue(),
                master_file.getvalue(),
                calendar_bytes,
                financial_year,
                period,
            )
            st.success("The report workbook has been generated.")
            st.subheader("Build summary")
            cols = st.columns(4)
            cols[0].metric("This month rows", summary["this_month_rows"])
            cols[1].metric("New YTD rows", summary["ytd_appended"])
            cols[2].metric("YTD duplicates skipped", summary["ytd_duplicates_skipped"])
            cols[3].metric("YTD total rows", summary["ytd_rows_after_append"])
            st.write(
                f"Filters included **{summary['included_rows']:,}** audit(s) for "
                f"**{summary['start_date']:%d/%m/%Y} – {summary['end_date']:%d/%m/%Y}**. "
                "Aborted audits are excluded. Excel is set to recalculate formulas when the workbook is opened."
            )
            client_bytes = make_client_workbook(output_bytes)
            st.download_button(
                "Download internal LIVE workbook",
                data=output_bytes,
                file_name=f"One-Stop Test Purchase Report P{period} - LIVE.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
            )
            st.download_button(
                "Download client-facing workbook (values only)",
                data=client_bytes,
                file_name=f"One-Stop Test Purchase Report P{period}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            st.caption(
                "The LIVE workbook retains the full internal workbook. The client-facing workbook contains only "
                "Summary Data through Product Performance, with formulas converted to calculated values."
            )
        except Exception as exc:
            st.error(f"Could not build the report: {exc}")
            st.exception(exc)
else:
    st.warning("Upload all four files to select a period and build the report.")
