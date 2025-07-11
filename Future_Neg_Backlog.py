import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("📊 Future Negative Backlog Analyzer")

# File uploaders
billing_file = st.file_uploader("Upload Billing Plan Excel File", type=["xlsx"])
backlog_file = st.file_uploader("Upload Backlog Excel File", type=["xlsx"])
engagement_file = st.file_uploader("Upload Engagement Manager Excel File", type=["xlsx"])

if billing_file and backlog_file and engagement_file:
    billing_df = pd.read_excel(billing_file, engine="openpyxl")
    backlog_df = pd.read_excel(backlog_file, engine="openpyxl")
    engagement_df = pd.read_excel(engagement_file, engine="openpyxl")

    # Summarize billing
    billing_summary = billing_df.groupby(
        ["WBS Element", "Sales Organization", "Sales Order"], as_index=False
    )["Billing Value"].sum()

    # Summarize backlog
    backlog_summary = backlog_df.groupby(
        ["WBS Element", "Sales Organization", "Sales Order"], as_index=False
    )[["Remaining Backlog", "Measurement customer Name 1"]].first()

    # Merge billing and backlog
    merged_df = pd.merge(
        billing_summary,
        backlog_summary,
        on=["WBS Element", "Sales Organization", "Sales Order"],
        how="left"
    )

    # Calculate Delta Backlog
    merged_df["Delta Backlog"] = (merged_df["Remaining Backlog"] - merged_df["Billing Value"]).round(2)

    # Merge with engagement manager
    engagement_df = engagement_df[["Sales Document", "Eng Mgr - First name", "Eng Mgr - Last name"]]
    merged_df = pd.merge(
        merged_df,
        engagement_df,
        left_on="Sales Order",
        right_on="Sales Document",
        how="left"
    ).drop(columns=["Sales Document"])

    # Reorder columns
    ordered_columns = [
        "Sales Organization", "Sales Order", "Measurement customer Name 1", "WBS Element",
        "Billing Value", "Remaining Backlog", "Delta Backlog",
        "Eng Mgr - First name", "Eng Mgr - Last name"
    ]
    merged_df = merged_df[ordered_columns]

    # Save to Excel with formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="Report")
    output.seek(0)

    # Highlight negative Delta Backlog
    wb = load_workbook(output)
    ws = wb.active
    header = [cell.value for cell in ws[1]]
    delta_col_idx = header.index("Delta Backlog") + 1
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for row in ws.iter_rows(min_row=2, min_col=delta_col_idx, max_col=delta_col_idx):
        for cell in row:
            if isinstance(cell.value, (int, float)) and cell.value < 0:
                cell.fill = yellow_fill

    # Save updated workbook to BytesIO
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    st.success("✅ Processing complete. Download your report below.")
    st.download_button(
        label="📥 Download Excel Report",
        data=final_output,
        file_name="Future_negative_backlog.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
