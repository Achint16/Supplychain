import streamlit as st
import pandas as pd
from io import BytesIO

def rename_columns(df):
    rename_map = {
        'Column1': 'SiteCode',
        'Column2': 'LocationCode',
        'Column3': 'Product',
        'Column4': 'Date',
        'Column5': 'Qty'
    }
    return df.rename(columns=rename_map)

def generate_pivot(df, pivot_type, group_by_site):
    df['Date'] = df['Date'].astype(str).str.zfill(8)
    df['Date2'] = pd.to_datetime(df['Date'], format='%Y%m%d')
    df['Month'] = df['Date2'].dt.to_period('M').dt.to_timestamp()
    df['Week_Monday'] = df['Date2'].apply(lambda x: x - pd.Timedelta(days=x.weekday()))
    df['Site'] = df['SiteCode'] + '-' + df['LocationCode']

    index_fields = ['Product']
    if group_by_site:
        index_fields = ['Site'] + index_fields

    if pivot_type == 'Month':
        pivot = pd.pivot_table(df, index=index_fields, columns='Month', values='Qty', aggfunc='sum', fill_value=0)
        pivot.columns = pivot.columns.strftime('%Y-%m')
    elif pivot_type == 'Week':
        pivot = pd.pivot_table(df, index=index_fields, columns='Week_Monday', values='Qty', aggfunc='sum', fill_value=0)
        pivot.columns = pivot.columns.strftime('%Y-%m-%d')
    elif pivot_type == 'Date':
        pivot = pd.pivot_table(df, index=index_fields, columns='Date2', values='Qty', aggfunc='sum', fill_value=0)
        pivot.columns = pivot.columns.strftime('%Y-%m-%d')
    else:
        return None

    return pivot.reset_index()

def pivot_to_sage_format(pivot_df):
    pivot_df.columns = [col.strftime('%Y-%m') if isinstance(col, pd.Timestamp) else col for col in pivot_df.columns]
    known_cols = [col for col in ['Site', 'Product'] if col in pivot_df.columns]
    pivot_flat = pivot_df.melt(id_vars=known_cols, var_name='Month', value_name='Qty')
    pivot_flat['Month'] = pd.to_datetime(pivot_flat['Month'], format='%Y-%m', errors='coerce')
    pivot_flat = pivot_flat.dropna(subset=['Month'])
    pivot_flat['Date'] = pivot_flat['Month'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')

    if 'Site' in pivot_flat.columns:
        pivot_flat['SiteCode'] = pivot_flat['Site'].str.split('-').str[0]
        pivot_flat['LocationCode'] = pivot_flat['Site'].str.split('-').str[1]
    else:
        pivot_flat['SiteCode'] = ''
        pivot_flat['LocationCode'] = ''

    return pivot_flat[['SiteCode', 'LocationCode', 'Product', 'Date', 'Qty']]

# -------------------------------
# Streamlit App UI Starts Here
# -------------------------------

st.set_page_config(page_title="Sales Forecast Adjustment", layout="wide")

st.markdown(
    "<h1 style='text-align:center; color:teal;'>Sales Forecast Adjustment</h1>",
    unsafe_allow_html=True
)

step = st.radio("Step", ["Upload CSV", "Generate Pivot", "Upload Updated Pivot", "Download Final Output"])

if step == "Upload CSV":
    uploaded_file = st.file_uploader("Upload Original CSV", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file, usecols=lambda col: col != 'Column6')
        df = rename_columns(df)
        st.session_state['df_original'] = df
        st.write("Original Data Preview:")
        st.dataframe(df.head(), use_container_width=True)

elif step == "Generate Pivot":
    if 'df_original' in st.session_state:
        pivot_type = st.selectbox("Choose Pivot Type", ["Month", "Week", "Date"])
        group_by_site = st.checkbox("Group by Site (SiteCode + LocationCode)?", value=True)
        pivot = generate_pivot(st.session_state['df_original'], pivot_type, group_by_site)
        st.session_state['pivot'] = pivot
        st.write("Pivot Table Preview:")
        st.dataframe(pivot.head(), use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            pivot.to_excel(writer, index=False)
        st.download_button("Download Pivot Table for Sales Team", data=buffer.getvalue(), file_name="pivot_table.xlsx")
    else:
        st.warning("Please upload a CSV first.")

elif step == "Upload Updated Pivot":
    uploaded_pivot = st.file_uploader("Upload Updated Pivot Table", type=["xlsx"])
    if uploaded_pivot:
        df = pd.read_excel(uploaded_pivot)
        st.session_state['pivot_updated'] = df
        st.success("✅ Pivot table uploaded and saved to session state.")
        st.write("Updated Pivot Preview:")
        st.dataframe(df.head(), use_container_width=True)

elif step == "Download Final Output":
    if 'pivot_updated' in st.session_state and 'df_original' in st.session_state:
        updated_df = pivot_to_sage_format(st.session_state['pivot_updated'])
        original_df = st.session_state['df_original'][['SiteCode', 'LocationCode', 'Product', 'Date', 'Qty']].copy()

        # Ensure consistent formatting
        updated_df['Date'] = updated_df['Date'].astype(str).str.zfill(8)
        original_df['Date'] = original_df['Date'].astype(str).str.zfill(8)

        # Aggregate to remove duplicate combinations
        updated_df = updated_df.groupby(['SiteCode', 'LocationCode', 'Product', 'Date'], as_index=False).agg({'Qty': 'sum'})

        if 'pivot' in st.session_state:
            generated_df = pivot_to_sage_format(st.session_state['pivot'])
            generated_df['Date'] = generated_df['Date'].astype(str).str.zfill(8)
            generated_df = generated_df.groupby(['SiteCode', 'LocationCode', 'Product', 'Date'], as_index=False).agg({'Qty': 'sum'})

            # Merge and apply precision-safe delta check
            updated_df['key'] = updated_df[['SiteCode', 'LocationCode', 'Product', 'Date']].agg('-'.join, axis=1)
            generated_df['key'] = generated_df[['SiteCode', 'LocationCode', 'Product', 'Date']].agg('-'.join, axis=1)

            merged = updated_df.merge(generated_df[['key', 'Qty']], on='key', how='left', suffixes=('', '_orig'))
            df_changed = merged[(merged['Qty_orig'].isna()) | (abs(merged['Qty'] - merged['Qty_orig']) > 1e-6)].drop(columns=['key', 'Qty_orig'])

            # Combine with original Sage data
            final_df = pd.concat([original_df, df_changed]).drop_duplicates(
                subset=['SiteCode', 'LocationCode', 'Product', 'Date'], keep='last'
            ).reset_index(drop=True)
        else:
            final_df = pd.concat([original_df, updated_df]).drop_duplicates(
                subset=['SiteCode', 'LocationCode', 'Product', 'Date'], keep='last'
            ).reset_index(drop=True)

        st.write("Final Output Preview:")
        st.dataframe(final_df.head(), use_container_width=True)

        csv = final_df.to_csv(index=False, header=False).encode('utf-8')
        st.download_button("Download Final CSV (No Header)", data=csv, file_name="final_output.csv")
    else:
        st.warning("Please complete the previous steps.")
