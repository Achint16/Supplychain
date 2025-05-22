import streamlit as st
import pandas as pd
from io import BytesIO

# ---------- UTILITIES ----------

def rename_columns(df):
    return df.rename(columns={
        'Column1': 'SiteCode',
        'Column2': 'LocationCode',
        'Column3': 'Product',
        'Column4': 'Description',
        'Column5': 'Date',
        'Column6': 'Qty'
    })

def revert_column_names(df):
    return df.rename(columns={
        'SiteCode': 'Column1',
        'LocationCode': 'Column2',
        'Product': 'Column3',
        'Description': 'Column4',
        'Date': 'Column5',
        'Qty': 'Column6'
    })

def generate_pivot(df):
    df['Date'] = df['Date'].astype(str).str.zfill(8)
    df['Date2'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
    df['Month'] = df['Date2'].dt.to_period('M').dt.to_timestamp()
    df['Site'] = df['SiteCode'] + '-' + df['LocationCode']
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)

    desc_lookup = df[['Site', 'Product', 'Description']].drop_duplicates(subset=['Site', 'Product'])

    pivot = pd.pivot_table(
        df,
        index=['Site', 'Product'],
        columns='Month',
        values='Qty',
        aggfunc='sum'
    ).reset_index()

    pivot = pd.merge(pivot, desc_lookup, how='left', on=['Site', 'Product'])
    month_cols = [col for col in pivot.columns if isinstance(col, pd.Timestamp)]
    pivot = pivot[['Site', 'Product', 'Description'] + month_cols]
    pivot.columns = [col.strftime('%Y-%m') if isinstance(col, pd.Timestamp) else col for col in pivot.columns]
    return pivot

def pivot_to_sage_format(pivot_df):
    pivot_long = pivot_df.melt(id_vars=['Site', 'Product', 'Description'], var_name='Month', value_name='Qty')
    pivot_long['Month'] = pd.to_datetime(pivot_long['Month'], errors='coerce')
    pivot_long = pivot_long.dropna(subset=['Month'])
    pivot_long['Date'] = pivot_long['Month'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')
    pivot_long['SiteCode'] = pivot_long['Site'].str.split('-').str[0]
    pivot_long['LocationCode'] = pivot_long['Site'].str.split('-').str[1]
    pivot_long['Qty'] = pd.to_numeric(pivot_long['Qty'], errors='coerce')
    return pivot_long[['SiteCode', 'LocationCode', 'Product', 'Description', 'Date', 'Qty']]

# ---------- STREAMLIT APP ----------

st.set_page_config(page_title="Sales Forecast Adjustment", layout="wide")
st.title("Sales Forecast Adjustment - Sage X3")

step = st.radio("Step", ["Upload Sage X3 File", "Generate Pivot", "Upload Modified Pivot", "Download Final Sage X3 Format"])

if step == "Upload Sage X3 File":
    uploaded_file = st.file_uploader("Upload Original CSV from Sage X3", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file, encoding='latin1', header=0, usecols=lambda col: col != 'Column7')
        df = rename_columns(df)
        st.session_state['df_original'] = df
        st.write("Preview:", df.head())

elif step == "Generate Pivot":
    if 'df_original' in st.session_state:
        pivot = generate_pivot(st.session_state['df_original'])
        st.session_state['pivot'] = pivot
        st.dataframe(pivot.fillna('').head(), use_container_width=True)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            pivot.to_excel(writer, index=False)
        st.download_button("Download Pivot for Sales Team", data=buffer.getvalue(), file_name="pivot_table.xlsx")
    else:
        st.warning("Please upload the original Sage X3 file first.")

elif step == "Upload Modified Pivot":
    uploaded_pivot = st.file_uploader("Upload Modified Pivot Table", type=["xlsx"])
    if uploaded_pivot:
        df = pd.read_excel(uploaded_pivot)
        st.session_state['pivot_modified'] = df
        st.success("Modified pivot uploaded.")
        st.dataframe(df.head(), use_container_width=True)

elif step == "Download Final Sage X3 Format":
    if 'pivot_modified' in st.session_state and 'pivot' in st.session_state:
        updated_flat = pivot_to_sage_format(st.session_state['pivot_modified'])
        original_flat = pivot_to_sage_format(st.session_state['pivot'])

        updated_flat['MonthStart'] = updated_flat['Date']
        original_flat['ParsedDate'] = pd.to_datetime(original_flat['Date'], format='%Y%m%d', errors='coerce')
        original_flat['MonthStart'] = original_flat['ParsedDate'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')

        updated_flat['key'] = updated_flat[['SiteCode', 'LocationCode', 'Product', 'MonthStart']].agg('-'.join, axis=1)
        original_flat['key'] = original_flat[['SiteCode', 'LocationCode', 'Product', 'MonthStart']].agg('-'.join, axis=1)

        merged = updated_flat.merge(original_flat[['key', 'Qty']], on='key', how='left', suffixes=('', '_orig'))
        merged['Qty'] = pd.to_numeric(merged['Qty'], errors='coerce')
        merged['Qty_orig'] = pd.to_numeric(merged['Qty_orig'], errors='coerce')

        df_changed = merged[(merged['Qty_orig'].isna()) | (abs(merged['Qty'] - merged['Qty_orig']) > 1e-6)].drop(columns=['Qty_orig'])

        df_original = st.session_state['df_original'].copy()
        df_original['Date'] = df_original['Date'].astype(str).str.zfill(8)
        df_original['Qty'] = pd.to_numeric(df_original['Qty'], errors='coerce')
        df_original['ParsedDate'] = pd.to_datetime(df_original['Date'], format='%Y%m%d', errors='coerce')
        df_original['MonthStart'] = df_original['ParsedDate'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')
        df_original['key'] = df_original[['SiteCode', 'LocationCode', 'Product', 'MonthStart']].agg('-'.join, axis=1)
        df_changed['key'] = df_changed[['SiteCode', 'LocationCode', 'Product', 'MonthStart']].agg('-'.join, axis=1)

        # Remove unchanged rows
        unchanged_keys = set(df_original['key']) - set(df_changed['key'])
        df_original_clean = df_original[df_original['key'].isin(unchanged_keys)]

        # REMOVE deleted products
        products_modified = set(df_changed['Product'].unique())
        df_original_clean = df_original_clean[df_original_clean['Product'].isin(products_modified)]

        # Final output
        final_df = pd.concat([df_original_clean, df_changed.drop(columns='key')], ignore_index=True)
        final_df = final_df[final_df[['SiteCode', 'LocationCode', 'Product', 'Date', 'Qty']].notna().all(axis=1)]
        final_df = final_df[final_df['Product'].astype(str).str.strip() != '']
        final_df = final_df.drop(columns=['key', 'ParsedDate', 'MonthStart'], errors='ignore')

        df_export = revert_column_names(final_df)
        df_export = df_export[['Column1', 'Column2', 'Column3', 'Column4', 'Column5', 'Column6']]
        csv = df_export.to_csv(index=False, header=False).encode('utf-8')
        st.download_button("Download Final CSV for Sage X3", data=csv, file_name="final_output.csv")
    else:
        st.warning("Please upload both the original and modified pivot.")
