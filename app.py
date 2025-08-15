import io
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
from data_processing import (
    load_timeworkbook,
    load_craft_order,
    load_address_book,
    prepare_report_data,
)

st.set_page_config(page_title="Work Order Daily Report", layout="wide")

st.title("Work Order Reporting App")
st.caption("Upload a work order export and generate a craft-group report for a selected production date.")

with st.sidebar:
    st.header("1) Upload Files")
    time_file = st.file_uploader(
        label="Time on Work Order export (.xlsx)",
        type=["xlsx"],
        key="time_file",
        help="Export should match the format of your 'Time on Work Order' report."
    )
    craft_file = st.file_uploader(
        label="Craft Group Order (.xlsx) — optional",
        type=["xlsx"],
        key="craft_file",
        help="If omitted, the sample order in /sample_data will be used."
    )
    addr_file = st.file_uploader(
        label="Address Book (.xlsx) — optional",
        type=["xlsx"],
        key="addr_file",
        help="If omitted, the sample address book in /sample_data will be used."
    )
    st.markdown("---")
    st.header("2) Options")
    show_detail = st.checkbox("Show detailed work order rows", value=True)
    show_person_summary = st.checkbox("Show per-person summary", value=True)

# Load reference files (with fallback to sample data)
craft_df = load_craft_order(craft_file) if craft_file else load_craft_order("sample_data/Craft Group Order.xlsx")
addr_df = load_address_book(addr_file) if addr_file else load_address_book("sample_data/MS_Address_Book_Sorted_craft removed.xlsx")

if time_file is None:
    st.info("⬆️ Upload a **Time on Work Order** export to begin.")
    st.stop()

# Load the work order export
time_df = load_timeworkbook(time_file)

# Date selector with only dates present in the file (mm/dd/yyyy)
if "Production Date" not in time_df.columns:
    st.error("The uploaded file does not contain a 'Production Date' column after parsing. Please verify the export format.")
    st.stop()

dates = sorted(pd.to_datetime(time_df["Production Date"]).dt.date.unique())
if len(dates) == 0:
    st.warning("No valid production dates detected in the uploaded file.")
    st.stop()

date_labels = [datetime.strftime(pd.to_datetime(d), "%m/%d/%Y") for d in dates]
label_to_date = dict(zip(date_labels, dates))

selected_label = st.selectbox("Select Production Date", options=date_labels, index=len(date_labels)-1)
selected_date = label_to_date[selected_label]

# Prepare report data
report = prepare_report_data(time_df, addr_df, craft_df, selected_date)

st.markdown(f"### Report for {selected_label}")

if report["unmapped_people"]:
    with st.expander("⚠️ Unmapped personnel (not found in Address Book)", expanded=False):
        st.dataframe(pd.DataFrame(report["unmapped_people"], columns=["AddressBookNumber", "Name"]).sort_values("Name"))

# Render by craft group in the defined order
for craft_name, payload in report["groups"]:
    st.markdown(f"#### {craft_name}")
    if show_person_summary and not payload['person_summary'].empty:
        st.subheader("Per-Person Summary (Hours & WO Count)")
        st.dataframe(payload["person_summary"], use_container_width=True)
    if show_detail and not payload['detail'].empty:
        st.subheader("Detail")
        st.dataframe(payload["detail"], use_container_width=True)
    st.markdown("---")

# Downloadable CSV of the full filtered detail
if not report["full_detail"].empty:
    csv = report["full_detail"].to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered detail (CSV)", data=csv, file_name=f"workorder_detail_{selected_label.replace('/', '-')}.csv", mime="text/csv")
