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

# ---------- PIVOT GENERATION ----------

def generate_pivot(df, pivot_type, group_by_site):
    df['Date'] = df['Date'].astype(str).str.zfill(8)
    df['Date2'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
    df['Month'] = df['Date2'].dt.to_period('M').dt.to_timestamp()
    df['Week_Monday'] = df['Date2'].apply(lambda x: x - pd.Timedelta(days=x.weekday()) if pd.notnull(x) else x)
    df['Site'] = df['SiteCode'] + '-' + df['LocationCode']
    
    # Clean descriptions — fill missing with blank (NOT 'null')
    df['Description'] = df['Description'].fillna('').astype(str)

    index_fields = ['Product']
    if group_by_site:
        index_fields = ['Site'] + index_fields

    if pivot_type == 'Month':
        pivot = pd.pivot_table(df, index=index_fields, columns='Month', values='Qty',
                               aggfunc='sum', fill_value=0, dropna=False)
        pivot.columns = pivot.columns.strftime('%Y-%m')
    elif pivot_type == 'Week':
        pivot = pd.pivot_table(df, index=index_fields, columns='Week_Monday', values='Qty',
                               aggfunc='sum', fill_value=0, dropna=False)
        pivot.columns = pivot.columns.strftime('%Y-%m-%d')
    elif pivot_type == 'Date':
        pivot = pd.pivot_table(df, index=index_fields, columns='Date2', values='Qty',
                               aggfunc='sum', fill_value=0, dropna=False)
        pivot.columns = pivot.columns.strftime('%Y-%m-%d')
    else:
        return None

    pivot = pivot.reset_index()

    # Add back descriptions from original only where available
    desc_map = df[df['Description'] != ''].groupby(['Site', 'Product'])['Description'].first().to_dict()
    pivot['Description'] = pivot.apply(lambda row: desc_map.get((row['Site'], row['Product']), ''), axis=1)

    # Reorder columns
    cols = pivot.columns.tolist()
    if 'Description' in cols:
        cols.insert(cols.index('Product') + 1, cols.pop(cols.index('Description')))
        pivot = pivot[cols]

    return pivot

# ---------- PIVOT TO SAGE FORMAT ----------

def pivot_to_sage_format(pivot_df):
    pivot_df.columns = [col.strftime('%Y-%m') if isinstance(col, pd.Timestamp) else col for col in pivot_df.columns]
    id_vars = [col for col in ['Site', 'Product', 'Description'] if col in pivot_df.columns]
    pivot_flat = pivot_df.melt(id_vars=id_vars, var_name='Month', value_name='Qty')
    pivot_flat['Month'] = pd.to_datetime(pivot_flat['Month'], format='%Y-%m', errors='coerce')
    pivot_flat['Date'] = pivot_flat['Month'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')

    if 'Site' in pivot_flat:
        pivot_flat['SiteCode'] = pivot_flat['Site'].str.split('-').str[0]
        pivot_flat['LocationCode'] = pivot_flat['Site'].str.split('-').str[1]
    else:
        pivot_flat['SiteCode'] = ''
        pivot_flat['LocationCode'] = ''

    # Return in Sage X3 column format
    return pivot_flat[['SiteCode', 'LocationCode', 'Product', 'Description', 'Date', 'Qty']]

# ---------- STREAMLIT UI ----------

st.set_page_config(page_title="Sales Forecast Adjustment", layout="wide")
st.markdown("<h1 style='text-align:center; color:teal;'>Sales Forecast Adjustment</h1>", unsafe_allow_html=True)
step = st.radio("Step", ["Upload CSV", "Generate Pivot", "Upload Updated Pivot", "Download Final Output"])

# Upload CSV
if step == "Upload CSV":
    uploaded_file = st.file_uploader("Upload Original CSV", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file, encoding='latin1', usecols=lambda col: col != 'Column7')
        df = rename_columns(df)
        st.session_state['df_original'] = df
        st.write("Original Data Preview:")
        st.dataframe(df.head(), use_container_width=True)

# Generate Pivot
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

# Upload Updated Pivot
elif step == "Upload Updated Pivot":
    uploaded_pivot = st.file_uploader("Upload Updated Pivot Table", type=["xlsx"])
    if uploaded_pivot:
        df = pd.read_excel(uploaded_pivot)
        st.session_state['pivot_updated'] = df
        st.success("✅ Pivot table uploaded.")
        st.write("Updated Pivot Preview:")
        st.dataframe(df.head(), use_container_width=True)

# Download Final Output
elif step == "Download Final Output":
    if all(k in st.session_state for k in ['pivot_updated', 'df_original', 'pivot']):
        df_updated = pivot_to_sage_format(st.session_state['pivot_updated'])
        df_generated = pivot_to_sage_format(st.session_state['pivot'])

        key_cols = ['SiteCode', 'LocationCode', 'Product', 'Description', 'Date']
        df_updated['key'] = df_updated[key_cols].astype(str).agg('-'.join, axis=1)
        df_generated['key'] = df_generated[key_cols].astype(str).agg('-'.join, axis=1)

        df_merged = df_updated.merge(df_generated[['key', 'Qty']], on='key', how='left', suffixes=('', '_orig'))
        df_changed = df_merged[df_merged['Qty'] != df_merged['Qty_orig']].drop(columns=['key', 'Qty_orig'])

        df_original = st.session_state['df_original'][key_cols + ['Qty']].copy()
        df_original['Description'] = df_original['Description'].fillna('')  # make sure no nulls
        df_final = pd.concat([df_original, df_changed]).drop_duplicates(subset=key_cols, keep='last').reset_index(drop=True)

        st.write("Final Output Preview:")
        st.dataframe(df_final.head(), use_container_width=True)

        df_final_export = revert_column_names(df_final)
        df_final_export['Column4'] = df_final_export['Column4'].replace('null', '')  # final cleanup

        csv = df_final_export.to_csv(index=False, header=False).encode('utf-8')
        st.download_button("Download Final CSV (No Header)", data=csv, file_name="final_output.csv")
    else:
        st.warning("Please complete all steps before downloading.")
