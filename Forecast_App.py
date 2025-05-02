import streamlit as st
import pandas as pd
from io import BytesIO

def generate_pivot(df, pivot_type, group_by_site):
    df['Date'] = df['Date'].astype(str).str.zfill(8)
    df['Date2'] = pd.to_datetime(df['Date'], format='%Y%m%d')
    df['Month'] = df['Date2'].dt.to_period('M').dt.to_timestamp()
    df['Week_Monday'] = df['Date2'].apply(lambda x: x - pd.Timedelta(days=x.weekday()))
    df['Site'] = df['Column1'] + '-' + df['Column2']

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
    index_cols = ['Site', 'Product'] if 'Site' in pivot_updated.columns else ['Product']
    pivot_flat = pivot_updated.melt(id_vars=index_cols, var_name='Month', value_name='New_Qty')

    date_format = '%Y-%m-%d' if '-' in pivot_flat['Month'].iloc[0] and len(pivot_flat['Month'].iloc[0]) > 7 else '%Y-%m'
    pivot_flat['Month'] = pd.to_datetime(pivot_flat['Month'], format=date_format)

    df_check = df_original.copy()
    df_check['Date'] = df_check['Date'].astype(str).str.zfill(8)
    df_check['Date2'] = pd.to_datetime(df_check['Date'], format='%Y%m%d')
    df_check['Month'] = df_check['Date2'].dt.to_period('M').dt.to_timestamp()
    df_check['Site'] = df_check['Column1'] + '-' + df_check['Column2']

    if 'Site' in index_cols:
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
    df['Site'] = df['Column1'] + '-' + df['Column2']

    if 'Site' in index_cols:
        df['Orig_Month_Total'] = df.groupby(['Site', 'Product', 'Month'])['Qty'].transform('sum')
        df['Weight'] = df['Qty'] / df['Orig_Month_Total']
        df = df.merge(pivot_flat, on=['Site', 'Product', 'Month'], how='left')
    else:
        df['Orig_Month_Total'] = df.groupby(['Product', 'Month'])['Qty'].transform('sum')
        df['Weight'] = df['Qty'] / df['Orig_Month_Total']
        df = df.merge(pivot_flat, on=['Product', 'Month'], how='left')

    df['Qty'] = df['Weight'] * df['New_Qty']
    df_final = df[['Column1', 'Column2', 'Product', 'Date', 'Qty']]

    # Show new rows only on demand
    if not new_combos.empty:
        st.warning("⚠️ New combinations were introduced that don't exist in the original file.")

        if st.checkbox("Show newly introduced records (with Date and Qty)"):
            if 'Date' not in new_combos.columns:
                new_combos['Date'] = new_combos['Month'].dt.to_period('M').dt.start_time.dt.strftime('%Y%m%d')

            new_combos['Column1'] = new_combos['Site'].str.split('-').str[0]
            new_combos['Column2'] = new_combos['Site'].str.split('-').str[1]
            new_combos['Qty'] = new_combos['New_Qty']
            new_combos_display = new_combos[['Column1', 'Column2', 'Product', 'Date', 'Qty']]
            st.dataframe(new_combos_display)

            new_combos_display['Source'] = 'New'
            df_final['Source'] = 'Original'
            df_final = pd.concat([df_final, new_combos_display], ignore_index=True)

    return df_final

def to_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

# -------------------------------
# Streamlit App UI Starts Here
# -------------------------------

st.markdown(
    "<h1 style='text-align:center; color:teal;'>Sales Forecast Adjustment</h1>",
    unsafe_allow_html=True
)

step = st.radio("Step", ["Upload CSV", "Generate Pivot", "Upload Updated Pivot", "Download Final Output"])

if step == "Upload CSV":
    uploaded_file = st.file_uploader("Upload Original CSV", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file)
        st.session_state['df_original'] = df
        st.write("Original Data Preview:")
        st.dataframe(df.head(), use_container_width=True)

elif step == "Generate Pivot":
    if 'df_original' in st.session_state:
        pivot_type = st.selectbox("Choose Pivot Type", ["Month", "Week", "Date"])
        group_by_site = st.checkbox("Group by Site (Column1 + Column2)?", value=True)
        pivot = generate_pivot(st.session_state['df_original'], pivot_type, group_by_site)
        st.session_state['pivot'] = pivot
        st.write("Pivot Table Preview:")
        st.dataframe(pivot.head(), use_container_width=True)

        excel_bytes = to_excel_download(pivot)
        st.download_button("Download Pivot Table for Sales Team", data=excel_bytes, file_name="pivot_table.xlsx")
    else:
        st.warning("Please upload a CSV first.")

elif step == "Upload Updated Pivot":
    uploaded_pivot = st.file_uploader("Upload Updated Pivot Table", type=["xlsx"])
    if uploaded_pivot:
        pivot_updated = pd.read_excel(uploaded_pivot)
        st.session_state['pivot_updated'] = pivot_updated
        st.write("Updated Pivot Preview:")
        st.dataframe(pivot_updated.head(), use_container_width=True)

elif step == "Download Final Output":
    if 'df_original' in st.session_state and 'pivot_updated' in st.session_state:
        final_output = reverse_pivot(st.session_state['df_original'], st.session_state['pivot_updated'])
        st.write("Final Output Preview (with Source column to highlight new rows):")
        st.dataframe(final_output.head(), use_container_width=True)

        csv = final_output.to_csv(index=False).encode('utf-8')
        st.download_button("Download Final CSV", data=csv, file_name="final_output.csv")
    else:
        st.warning("Please complete the previous steps.")
