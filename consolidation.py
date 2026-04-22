# consolidation_app.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from openpyxl import load_workbook

# ----------------------------------------------------------------------
# Page config
# ----------------------------------------------------------------------
st.set_page_config(page_title="Inventory Consolidation Tool", layout="wide")
st.title("📦 Inventory Consolidation Tool")
st.markdown("Upload your Excel file (`MIRA - ALL SOURCE DATA.xlsx`) to generate a consolidated CSV with stock data by item code and supplier columns.")

# ----------------------------------------------------------------------
# Helper functions (copied from your script, adapted for Streamlit)
# ----------------------------------------------------------------------
def normalize_name(name):
    name = str(name).strip().lower()
    name = re.sub(r'[^a-z0-9]', '', name)
    return name

def clean_sheet_name(name):
    name = re.sub(r'\(.*?\)', '', name).strip()
    name = name.replace(' ', '_')
    name = re.sub(r'[^a-zA-Z0-9_]', '', name)
    return name

def clean_numeric_column(series):
    cleaned = series.astype(str).str.replace(r'[^\d.-]', '', regex=True)
    cleaned = cleaned.replace('', np.nan)
    return pd.to_numeric(cleaned, errors='coerce')

def find_column(df, possible_names):
    for col in df.columns:
        for name in possible_names:
            if col.strip().lower() == name.lower():
                return col
    return None

def find_description_column(df):
    """Find a column that likely contains product description."""
    patterns = [
        r'kinshasa.*description',
        r'product.*description',
        r'erp.*description',
        r'item name',
        r'description'
    ]
    for col in df.columns:
        col_lower = col.strip().lower()
        for pat in patterns:
            if re.search(pat, col_lower):
                return col
    return None

def standardize_excel_in_memory(uploaded_file):
    """Read uploaded Excel, standardize sheets, return a dictionary of DataFrames."""
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names
    standardized_sheets = {}
    for original_name in sheet_names:
        df = pd.read_excel(xl, sheet_name=original_name, header=0)
        if df.empty:
            continue
        if original_name == "Compiled Data":
            standardized_sheets[original_name] = df
            continue
        item_col = find_column(df, ['item code', 'item_code', 'itemcode', 'code'])
        if item_col is None:
            st.warning(f"Sheet '{original_name}' has no item code column. Skipping.")
            continue
        phys_col = find_column(df, ['physical_stock', 'physical stock', 'physical stock (pcs)'])
        pend_col = find_column(df, ['pending_orders', 'pending orders', 'pending order'])
        trans_col = find_column(df, ['total_qty_of_in_transit', 'total qty of in transit', 'total in transit'])
        desc_col = find_description_column(df)
        new_df = pd.DataFrame()
        new_df["Item Code"] = df[item_col]
        if desc_col:
            new_df["Description"] = df[desc_col].astype(str).str.strip()
        else:
            new_df["Description"] = ""
        def add_stock_column(new_name, actual_col):
            if actual_col is not None:
                cleaned = clean_numeric_column(df[actual_col])
            else:
                cleaned = pd.Series([np.nan] * len(df), dtype='float64')
            new_df[new_name] = cleaned
        add_stock_column("PHYSICAL_STOCK", phys_col)
        add_stock_column("PENDING_ORDERS", pend_col)
        add_stock_column("TOTAL_QTY_OF_IN_TRANSIT", trans_col)
        new_df = new_df.dropna(subset=["Item Code"])
        new_sheet_name = clean_sheet_name(original_name)
        # Avoid duplicate names
        if new_sheet_name in standardized_sheets:
            new_sheet_name = f"{new_sheet_name}_{len(standardized_sheets)}"
        standardized_sheets[new_sheet_name] = new_df
    return standardized_sheets

# ----------------------------------------------------------------------
# Consolidation function
# ----------------------------------------------------------------------
def consolidate(sheets_dict, consolidation_rule='max', supplier_priority=None):
    if supplier_priority is None:
        supplier_priority = ['BIOMATRIX', 'LINCOLN', 'SCOTT', 'ZEST', 'INTAS']
    master_df = sheets_dict["Compiled Data"].copy()
    required_cols = ["Supplier Name", "Item Code", "Item Name (Local)", "Category"]
    for col in required_cols:
        if col not in master_df.columns:
            st.error(f"Missing required column in 'Compiled Data': {col}")
            return None
    master_df = master_df[required_cols].copy()

    stock_dict = {}          # (norm_supplier, item_code) -> (phys, pend, trans, sheet_name, desc)
    description_map = {}     # (sheet, item_code) -> description

    for sheet_name, df in sheets_dict.items():
        if sheet_name == "Compiled Data":
            continue
        if "Item Code" not in df.columns:
            continue
        norm_supplier = normalize_name(sheet_name)
        phys_col = "PHYSICAL_STOCK" if "PHYSICAL_STOCK" in df.columns else None
        pend_col = "PENDING_ORDERS" if "PENDING_ORDERS" in df.columns else None
        trans_col = "TOTAL_QTY_OF_IN_TRANSIT" if "TOTAL_QTY_OF_IN_TRANSIT" in df.columns else None
        desc_col = "Description" if "Description" in df.columns else None
        if phys_col is None and pend_col is None and trans_col is None:
            continue
        for _, row in df.iterrows():
            item_code = row["Item Code"]
            if pd.isna(item_code):
                continue
            item_code = str(item_code).strip()
            phys = row.get(phys_col) if phys_col else None
            pend = row.get(pend_col) if pend_col else None
            trans = row.get(trans_col) if trans_col else None
            desc = row.get(desc_col) if desc_col else ""
            if desc and desc != "nan":
                description_map[(sheet_name, item_code)] = desc
            key = (norm_supplier, item_code)
            if key not in stock_dict:
                stock_dict[key] = (phys, pend, trans, sheet_name, desc)

    # Build detail DataFrame
    detail_rows = []
    for (norm_supplier, item_code), (phys, pend, trans, sheet_name, desc) in stock_dict.items():
        detail_rows.append({
            "Supplier Name": sheet_name,
            "Item Code": item_code,
            "PHYSICAL_STOCK": phys,
            "PENDING_ORDERS": pend,
            "IN_TRANSIT_TOTAL": trans,
            "Description": desc
        })
    detail_df = pd.DataFrame(detail_rows)

    # Aggregate numeric columns
    if consolidation_rule == 'sum':
        agg_df = detail_df.groupby("Item Code", as_index=False).agg({
            "PHYSICAL_STOCK": "sum",
            "PENDING_ORDERS": "sum",
            "IN_TRANSIT_TOTAL": "sum"
        })
    elif consolidation_rule == 'max':
        agg_df = detail_df.groupby("Item Code", as_index=False).agg({
            "PHYSICAL_STOCK": "max",
            "PENDING_ORDERS": "max",
            "IN_TRANSIT_TOTAL": "max"
        })
    elif consolidation_rule == 'mean':
        agg_df = detail_df.groupby("Item Code", as_index=False).agg({
            "PHYSICAL_STOCK": "mean",
            "PENDING_ORDERS": "mean",
            "IN_TRANSIT_TOTAL": "mean"
        })
    elif consolidation_rule == 'first':
        def first_by_priority(group):
            group = group.copy()
            group["priority"] = group["Supplier Name"].apply(
                lambda x: supplier_priority.index(x) if x in supplier_priority else len(supplier_priority)
            )
            group = group.sort_values("priority")
            return group.iloc[0]
        agg_df = detail_df.groupby("Item Code").apply(first_by_priority).reset_index(drop=True)
        agg_df = agg_df[["Item Code", "PHYSICAL_STOCK", "PENDING_ORDERS", "IN_TRANSIT_TOTAL"]]
    else:
        st.error(f"Unknown consolidation rule: {consolidation_rule}")
        return None

    # Supplier list per item
    supplier_list = detail_df.groupby("Item Code")["Supplier Name"].apply(lambda x: list(x.unique())).reset_index()
    supplier_list.columns = ["Item Code", "Supplier Names"]

    # Merge with master
    result_df = agg_df.merge(supplier_list, on="Item Code", how="outer")
    result_df = result_df.merge(master_df[["Item Code", "Item Name (Local)", "Category"]].drop_duplicates("Item Code"),
                                on="Item Code", how="left")

    # Fill missing item names from description_map
    item_desc_map = {}
    for (sheet, item_code), desc in description_map.items():
        if item_code not in item_desc_map:
            item_desc_map[item_code] = desc
    def fill_item_name(row):
        orig = row["Item Name (Local)"]
        if pd.isna(orig) or str(orig).strip() in ["", "#N/A", "nan"]:
            return item_desc_map.get(row["Item Code"], orig)
        return orig
    result_df["Item Name (Local)"] = result_df.apply(fill_item_name, axis=1)

    # Create supplier columns
    max_suppliers = result_df["Supplier Names"].apply(len).max() if len(result_df) > 0 else 0
    for i in range(1, max_suppliers + 1):
        result_df[f"Supplier Name {i}"] = result_df["Supplier Names"].apply(lambda x: x[i-1] if len(x) >= i else "")
    result_df["Without Supplier Name"] = result_df["Supplier Names"].apply(lambda x: "" if len(x) > 0 else "No Supplier Data")
    result_df = result_df.drop(columns=["Supplier Names"])

    # Reorder columns
    col_order = ["Item Name (Local)", "Item Code", "PHYSICAL_STOCK", "PENDING_ORDERS", "IN_TRANSIT_TOTAL"]
    for i in range(1, max_suppliers + 1):
        col_order.append(f"Supplier Name {i}")
    col_order.append("Without Supplier Name")
    col_order.append("Category")
    result_df = result_df[[c for c in col_order if c in result_df.columns]]

    return result_df

# ----------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload Excel file (xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing file... This may take a few seconds."):
        try:
            sheets = standardize_excel_in_memory(uploaded_file)
            if "Compiled Data" not in sheets:
                st.error("The uploaded file does not contain a sheet named 'Compiled Data'.")
                st.stop()
            result_df = consolidate(sheets, consolidation_rule='max')
            if result_df is not None:
                st.success("Consolidation completed successfully!")
                
                # Preview
                st.subheader("Preview of consolidated data (first 10 rows)")
                st.dataframe(result_df.head(10), use_container_width=True)
                
                # Download CSV
                csv_data = result_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Consolidated CSV",
                    data=csv_data,
                    file_name="final_stock_with_suppliers.csv",
                    mime="text/csv"
                )
                
                # Verification report
                total_items = len(result_df)
                items_with_stock = result_df[result_df["PHYSICAL_STOCK"].notna() | 
                                             result_df["PENDING_ORDERS"].notna() | 
                                             result_df["IN_TRANSIT_TOTAL"].notna()]
                items_without_stock = result_df[result_df["Without Supplier Name"] == "No Supplier Data"]
                max_supp = max([len([c for c in result_df.columns if c.startswith("Supplier Name")])])
                st.subheader("Verification Report")
                st.write(f"- Total items in master: {total_items}")
                st.write(f"- Items with at least one supplier stock record: {len(items_with_stock)}")
                st.write(f"- Items with no stock data: {len(items_without_stock)}")
                st.write(f"- Maximum number of suppliers for a single item: {max_supp}")
                
                # Provide option to download report
                report_text = f"""CONSOLIDATION VERIFICATION REPORT
Consolidation rule: max
Total items: {total_items}
Items with stock data: {len(items_with_stock)}
Items without stock data: {len(items_without_stock)}
Maximum suppliers per item: {max_supp}
"""
                st.download_button("📄 Download Verification Report", report_text, file_name="verification_report.txt")
        except Exception as e:
            st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to begin.")

st.markdown("---")
st.caption("This tool replicates the consolidation logic from your local script. Upload the exact same Excel structure (sheet 'Compiled Data' and supplier sheets).")