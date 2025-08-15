import pandas as pd
import numpy as np
from datetime import datetime
from typing import Tuple, Dict, Any, List

REQUIRED_TIME_COLUMNS = [
    "AddressBookNumber", "Name", "Production Date", "OrderNumber", "Sum of Hours.",
    "Hours Estimated", "Status", "Type", "PMFrequency", "Description",
    "Department", "Location", "Equipment", "PM Number", "PM"
]

def _find_header_row(df_raw: pd.DataFrame) -> int:
    """
    The Time on Work Order export often has 'Applied filters' text above the true header row.
    This function finds the row index where the first column equals 'AddressBookNumber'.
    """
    first_col = df_raw.columns[0]
    # Search value 'AddressBookNumber' anywhere in the first column
    mask = df_raw[first_col].astype(str).str.strip() == "AddressBookNumber"
    matches = df_raw.index[mask].tolist()
    if not matches:
        # Fallback: try to find the row that contains all/most required headers
        for i in range(min(10, len(df_raw))):
            row_vals = df_raw.iloc[i].astype(str).str.strip().tolist()
            if "AddressBookNumber" in row_vals and "Production Date" in row_vals:
                return i
        raise ValueError("Could not locate header row containing 'AddressBookNumber'.")
    return matches[0]

def load_timeworkbook(file_like) -> pd.DataFrame:
    """
    Load and normalize the 'Time on Work Order' export.
    - Detects header row dynamically
    - Ensures required columns exist
    - Normalizes datatypes
    """
    df_raw = pd.read_excel(file_like, header=None, dtype=str)
    header_row = _find_header_row(df_raw)
    df = pd.read_excel(file_like, header=header_row)

    # Drop completely empty columns
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    # Ensure required columns present (some may be missing depending on export options)
    missing = [c for c in REQUIRED_TIME_COLUMNS if c not in df.columns]
    if missing:
        # Only warn; still proceed with the columns we have
        # Add any missing columns as empty to keep downstream logic simple
        for c in missing:
            df[c] = pd.NA

    # Normalize types
    # AddressBookNumber may be numeric-like; keep as string to join reliably
    df["AddressBookNumber"] = df["AddressBookNumber"].astype(str).str.strip()
    if "Production Date" in df.columns:
        df["Production Date"] = pd.to_datetime(df["Production Date"], errors="coerce").dt.date

    # Rename for convenience
    df = df.rename(columns={"Sum of Hours.": "Hours"})
    return df

def load_craft_order(file_like) -> pd.DataFrame:
    df = pd.read_excel(file_like)
    # Expect a single column named 'Craft Description'
    if "Craft Description" not in df.columns:
        # Try to find the column with "Craft" in name
        cand = [c for c in df.columns if "craft" in c.lower()]
        if cand:
            df = df.rename(columns={cand[0]: "Craft Description"})
        else:
            raise ValueError("Craft Group Order file must contain a 'Craft Description' column.")
    df["Craft Description"] = df["Craft Description"].astype(str).str.strip()
    # Drop blanks
    df = df[df["Craft Description"].str.len() > 0].reset_index(drop=True)
    return df

def load_address_book(file_like) -> pd.DataFrame:
    df = pd.read_excel(file_like, dtype={"AddressBookNumber": str})
    # Standardize expected columns
    rename_map = {}
    if "Craft Description" not in df.columns:
        # Attempt heuristic rename if needed
        cand = [c for c in df.columns if "craft" in c.lower()]
        if cand:
            rename_map[cand[0]] = "Craft Description"
    if "AddressBookNumber" not in df.columns:
        cand = [c for c in df.columns if "address" in c.lower() and "number" in c.lower()]
        if cand:
            rename_map[cand[0]] = "AddressBookNumber"
    if rename_map:
        df = df.rename(columns=rename_map)

    for col in ["AddressBookNumber", "Name", "Craft Description"]:
        if col not in df.columns:
            raise ValueError(f"Address Book must contain '{col}' column.")

    df["AddressBookNumber"] = df["AddressBookNumber"].astype(str).str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()
    df["Craft Description"] = df["Craft Description"].astype(str).str.strip()
    return df[["AddressBookNumber", "Name", "Craft Description"]]

def _categorical_craft_order(df: pd.DataFrame, order_df: pd.DataFrame) -> pd.DataFrame:
    order = order_df["Craft Description"].tolist()
    # Ensure unique while preserving order
    seen = set()
    order_unique = []
    for c in order:
        if c not in seen:
            order_unique.append(c)
            seen.add(c)
    cat = pd.CategoricalDtype(order_unique + ["Unassigned"], ordered=True)
    df["Craft Description"] = df["Craft Description"].fillna("Unassigned")
    df["Craft Description"] = df["Craft Description"].astype(pd.CategoricalDtype(cat.categories, ordered=True))
    return df

def prepare_report_data(time_df: pd.DataFrame,
                        addr_df: pd.DataFrame,
                        craft_order_df: pd.DataFrame,
                        selected_date) -> Dict[str, Any]:
    # Filter by selected date
    f = time_df.copy()
    f = f[f["Production Date"] == selected_date].copy()

    # Merge in craft from address book using AddressBookNumber first (most reliable)
    f["AddressBookNumber"] = f["AddressBookNumber"].astype(str).str.strip()
    addr_df["AddressBookNumber"] = addr_df["AddressBookNumber"].astype(str).str.strip()
    merged = f.merge(addr_df[["AddressBookNumber", "Craft Description", "Name"]].rename(columns={"Name": "AB_Name"}),
                     on="AddressBookNumber", how="left")
    # Prefer the Name from the time export, but if missing, fallback to Address Book name
    merged["Name"] = merged["Name"].fillna(merged["AB_Name"])
    merged = merged.drop(columns=["AB_Name"])

    # Identify unmapped people (no craft group)
    unmapped = []
    mask_unmapped = merged["Craft Description"].isna() | (merged["Craft Description"].astype(str).str.len() == 0)
    if mask_unmapped.any():
        unmapped = merged.loc[mask_unmapped, ["AddressBookNumber", "Name"]].drop_duplicates().to_dict("records")
        merged.loc[mask_unmapped, "Craft Description"] = "Unassigned"

    # Apply craft order
    merged = _categorical_craft_order(merged, craft_order_df)

    # Build per-person summary and detail views by craft
    detail_cols = [c for c in ["Name", "AddressBookNumber", "OrderNumber", "Description", "Hours",
                               "Department", "Location", "Equipment", "PM Number", "Status", "Type", "PMFrequency"]
                   if c in merged.columns]

    groups_payload: List = []
    for craft in merged["Craft Description"].cat.categories:
        g = merged[merged["Craft Description"] == craft]
        if g.empty:
            continue

        # Per-person summary
        person_summary = (
            g.groupby(["Name", "AddressBookNumber"], dropna=False)
             .agg(Hours=("Hours", "sum"),
                  WorkOrders=("OrderNumber", "nunique"))
             .reset_index()
             .sort_values(["Hours", "Name"], ascending=[False, True])
        )
        # Ensure numeric Hours
        if "Hours" in person_summary.columns:
            person_summary["Hours"] = pd.to_numeric(person_summary["Hours"], errors="coerce").fillna(0.0).round(2)

        detail = g[detail_cols].copy()
        # Sort detail by Name then OrderNumber
        if "Name" in detail.columns and "OrderNumber" in detail.columns:
            detail = detail.sort_values(["Name", "OrderNumber"])

        groups_payload.append((str(craft), {"person_summary": person_summary, "detail": detail}))

    full_detail = merged[detail_cols + ["Craft Description"]].copy()
    full_detail = full_detail.sort_values(["Craft Description", "Name", "OrderNumber"])

    return {
        "groups": groups_payload,
        "full_detail": full_detail,
        "unmapped_people": unmapped,
    }
