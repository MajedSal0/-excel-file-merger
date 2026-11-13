import streamlit as st
from merger import merge_files, to_excel_bytes

st.set_page_config(page_title='Excel & CSV File Merger', page_icon='📁', layout='wide')
st.title('📁 Excel & CSV File Merger')
st.write('Combine multiple Excel and CSV files into one downloadable report. Columns are matched by header name.')
st.caption('Each file must have a header row. For Excel files, the first worksheet is used.')

files = st.file_uploader('Upload Excel or CSV files', type=['xlsx', 'csv'], accept_multiple_files=True)

if files:
    try:
        merged = merge_files(files)
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

    c1, c2, c3 = st.columns(3)
    c1.metric('Files merged', len(files))
    c2.metric('Total rows', f'{len(merged):,}')
    c3.metric('Columns', len(merged.columns))
    st.subheader('Merged data preview')
    st.dataframe(merged, use_container_width=True)
    st.info('If files have different column names, unmatched cells are left blank. Check the preview before downloading.')
    st.download_button('📥 Download merged Excel', to_excel_bytes(merged), file_name='merged_report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
