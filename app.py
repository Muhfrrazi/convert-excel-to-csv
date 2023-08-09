import streamlit as st
import pandas as pd
import os
from docx import Document

def process_files(csv_files):
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, sep=";")
        if 'NO;"NPWP15";"Nama_WP"' in df.columns:
            df['NO;"NPWP15";"Nama_WP"'] = df['NO;"NPWP15";"Nama_WP"'].str.replace('"', '')

        excel_file = csv_file.replace('.csv', '.xlsx')

        if os.path.exists(excel_file):
            st.warning(f"{excel_file} already exists. Overwriting it.")
            df.to_excel(excel_file, index=False)
            if 'NO;"NPWP15";"Nama_WP"' in df.columns:
                process_excel(excel_file)
            else:
                output_file = excel_file.replace('.xlsx', '_output.xlsx')
                output_label.text(f'Output File: {output_file}')
                history_listbox.write(excel_file)
        else:
            df.to_excel(excel_file, index=False)
            if 'NO;"NPWP15";"Nama_WP"' in df.columns:
                process_excel(excel_file)
            else:
                output_file = excel_file.replace('.xlsx', '_output.xlsx')
                output_label.text(f'Output File: {output_file}')
                history_listbox.write(excel_file)

def process_excel_to_csv(file_paths):
    dfs = []
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        df = df.applymap(lambda x: x.replace('"', '') if isinstance(x, str) else x)
        dfs.append(df)

    combined_df = pd.concat(dfs, ignore_index=True)
    output_file_csv = file_paths[0].replace('.xlsx', '_combined_output.csv')

    if os.path.exists(output_file_csv):
        st.warning(f"{output_file_csv} already exists. Overwriting it.")
        combined_df.to_csv(output_file_csv, index=False, header=True)
        output_label.text(f'Converted to CSV: {output_file_csv}')
        history_listbox.write(output_file_csv)
    else:
        combined_df.to_csv(output_file_csv, index=False, header=True)
        output_label.text(f'Converted to CSV: {output_file_csv}')
        history_listbox.write(output_file_csv)

def process_excel(excel_file):
    df = pd.read_excel(excel_file)

    if 'NO;"NPWP15";"Nama_WP"' in df.columns:
        df[['NO', 'NPWP15', 'Nama_WP']] = df['NO;"NPWP15";"Nama_WP"'].str.split(';', expand=True)
    else:
        st.warning("Column 'NO;NPWP15;Nama_WP' not found in the DataFrame.")

    df.drop(columns=['NO;"NPWP15";"Nama_WP"'], inplace=True)
    output_file = excel_file.replace('.xlsx', '_output.xlsx')
    df.to_excel(output_file, index=False)
    output_label.text(f'Output File: {output_file}')
    # update_history_listbox()  # Make sure this function is defined

st.title('File Converter')

csv_files = st.file_uploader("Choose CSV Files", type="csv", accept_multiple_files=True)
excel_files = st.file_uploader("Choose Excel Files", type="xlsx", accept_multiple_files=True)

if csv_files:
    process_files(csv_files)

if excel_files:
    process_excel_to_csv(excel_files)

output_label = st.empty()
history_listbox = st.empty()

st.write('Conversion Log:')
conversion_history = []

if st.button("Clear Log"):
    conversion_history.clear()

for item in conversion_history:
    st.write(item)

st.button("Print History to Word")
