import streamlit as st
import pandas as pd
import os

def process_files(csv_files):
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, sep=";")
        if 'NO;"NPWP15";"Nama_WP"' in df.columns:
            df['NO;"NPWP15";"Nama_WP"'] = df['NO;"NPWP15";"Nama_WP"'].str.replace('"', '')

        excel_file_name = os.path.splitext(csv_file.name)[0] + '.xlsx'
        excel_file_path = os.path.join('temp', excel_file_name)
        
        if os.path.exists(excel_file_path):
            st.warning(f"{excel_file_name} already exists. Overwriting it.")
        df.to_excel(excel_file_path, index=False)
        process_excel(excel_file_path)

def process_excel_to_csv(excel_files):
    dfs = []
    for excel_file in excel_files:
        df = pd.read_excel(excel_file)
        df = df.applymap(lambda x: x.replace('"', '') if isinstance(x, str) else x)
        dfs.append(df)

    combined_df = pd.concat(dfs, ignore_index=True)
    output_file_csv_name = os.path.splitext(excel_files[0].name)[0] + '_combined_output.csv'
    output_file_csv_path = os.path.join('temp', output_file_csv_name)

    if os.path.exists(output_file_csv_path):
        st.warning(f"{output_file_csv_name} already exists. Overwriting it.")
    combined_df.to_csv(output_file_csv_path, index=False, header=True)
    output_label.text(f'Converted to CSV: {output_file_csv_name}')
    history_listbox.write(output_file_csv_name)

def process_excel(excel_file):
    df = pd.read_excel(excel_file)

    if 'NO;"NPWP15";"Nama_WP"' in df.columns:
        split_columns = df['NO;"NPWP15";"Nama_WP"'].str.split(';', expand=True)
        if split_columns.shape[1] == 3:
            df[['NO', 'NPWP15', 'Nama_WP']] = split_columns
        else:
            st.warning("Unexpected format in 'NO;NPWP15;Nama_WP' column.")
    else:
        st.warning("Column 'NO;NPWP15;Nama_WP' not found in the DataFrame.")

    df.drop(columns=['NO;"NPWP15";"Nama_WP"'], inplace=True)
    output_file_name = os.path.splitext(excel_file.name)[0] + '_output.xlsx'
    output_file_path = os.path.join('temp', output_file_name)
    df.to_excel(output_file_path, index=False)
    output_label.text(f'Output File: {output_file_name}')

st.title('File Converter')

os.makedirs('temp', exist_ok=True)

csv_files = st.file_uploader("Choose CSV Files", type="csv", accept_multiple_files=True)
excel_files = st.file_uploader("Choose Excel Files", type="xlsx", accept_multiple_files=True)

output_label = st.empty()
history_listbox = st.empty()

if csv_files:
    process_files(csv_files)

if excel_files:
    process_excel_to_csv(excel_files)

st.write('Conversion Log:')
conversion_history = []

if st.button("Clear Log"):
    conversion_history.clear()

for item in conversion_history:
    st.write(item)

st.button("Print History to Word")
