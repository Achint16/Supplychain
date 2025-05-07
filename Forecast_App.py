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

def revert_column_names(df):
    revert_map = {
        'SiteCode': 'Column1',
        'LocationCode': 'Column2',
        'Product': 'Column3',
        'Date': 'Column4',
        'Qty': 'Column5'
    }
    return df.rename(columns=revert_map)

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

def reverse_pivot(df_original, pivot_updated):
    pivot_updated.columns = pivot_updated.columns.str.strip()
    known_cols = [col for col in ['Site', 'Product'] if col in pivot_updated.columns]
    if not known_cols:
        st.error("❌ Pivot table is missing required index columns like 'Product' or 'Site'. Please check your file.")
        return pd.DataFrame()

    pivot_flat = pivot_updated.melt(id_vars=known_cols, var_name='Month', value_name='New_Qty')

    date_format = '%Y-%m-%d' if '-' in pivot_flat['Month'].iloc[0] and len(pivot_flat['Month'].iloc[0]) > 7 else '%Y-%m'
    pivot_flat['Month'] = pd.to_datetime(pivot_flat['Month'], format=date_format)

    df_check = df_original.copy()
    df_check['Date'] = df_check['Date'].astype(str).str.zfill(8)
    df_check['Date2'] = pd.to_datetime(df_check['Date'], format='%Y%m%d')
    df_check['Month'] = df_check['Date2'].dt.to_period('M').dt.to_timestamp()
    df_check['Site'] = df_check['SiteCode'] + '-' + df_check['LocationCode']

    if 'Site' in known_cols:
        original_combos = df_check[['Site', 'Product', 'Month']].drop_duplicates()
        merge_keys = ['Site', 'Product', 'Month']
    else:
        original_combos = df_check[['Product', 'Month']].drop_duplicates()
        merge_keys = ['Product', 'Month']

    new_combos = pivot_flat.merge(original_combos, on=merge_keys, how='left', indicator=True)
    added_combos = new_combos[new_combos['_merge'] == 'left_only'].copy()
    new_combos = added_combos.drop(columns=['_merge'])

    df = df_original.copy()
    df['Date'] = df['Date'].astype(str).str.zfill(8)
    df['Date2'] = pd.to_datetime(df['Date'], format='%Y%m%d')
    df['Month'] = df['Date2'].dt.to_period('M').dt.to_timestamp()
    df['Site'] = df['SiteCode'] + '-' + df['LocationCode']

    if 'Site' in known_cols:
        df['Orig_Month_Total'] = df.groupby(['Site', 'Product', 'Month'])['Qty'].transform('sum')
        df['Weight'] = df['Qty'] / df['Orig_Month_Total']
        df = df.merge(pivot_flat, on=['Site', 'Product', 'Month'], how='left')
    else:
        df['Orig_Month_Total'] = df.groupby(['Product', 'Month'])['Qty'].transform('sum')
        df['Weight'] = df['Qty'] / df['Orig_Month_Total']
        df = df.merge(pivot_flat, on=['Product', 'Month'], how='left')

    df['Qty'] = df['Weight'] * df['New_Qty']
    df_final = df[['SiteCode', 'LocationCode', 'Product', 'Date', 'Qty']].copy()

    if not new_combos.empty:
        new_combos['Date'] = new_combos['Month'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')
        if 'Site' in new_combos.columns:
            new_combos['SiteCode'] = new_combos['Site'].str.split('-').str[0]
            new_combos['LocationCode'] = new_combos['Site'].str.split('-').str[1]
        else:
            new_combos['SiteCode'] = ''
            new_combos['LocationCode'] = ''
        new_combos['Qty'] = new_combos['New_Qty']
        new_combos_display = new_combos[['SiteCode', 'LocationCode', 'Product', 'Date', 'Qty']]
        df_final = pd.concat([df_final, new_combos_display], ignore_index=True)

    df_final = revert_column_names(df_final)
    return df_final


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
        df = pd.read_csv(uploaded_file)
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
    if 'df_original' in st.session_state and 'pivot_updated' in st.session_state:
        final_output = reverse_pivot(st.session_state['df_original'], st.session_state['pivot_updated'])
        st.write("Final Output Preview:")
        st.dataframe(final_output.head(), use_container_width=True)

        csv = final_output.to_csv(index=False, header=False).encode('utf-8')
        st.download_button("Download Final CSV (No Header)", data=csv, file_name="final_output.csv")
    else:
        st.warning("Please complete the previous steps.")
