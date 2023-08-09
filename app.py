import streamlit as st
import pandas as pd
import os

def process_files(csv_files):
    output_excel_files = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, sep=";")
        if 'NO;"NPWP15";"Nama_WP"' in df.columns:
            df['NO;"NPWP15";"Nama_WP"'] = df['NO;"NPWP15";"Nama_WP"'].str.replace('"', '')

        excel_file_name = os.path.splitext(csv_file.name)[0] + '.xlsx'
        excel_file_path = os.path.join('temp', excel_file_name)
        
        if os.path.exists(excel_file_path):
            st.warning(f"{excel_file_name} already exists. Overwriting it.")
        df.to_excel(excel_file_path, index=False)
        output_excel_files.append(excel_file_path)
    
    return output_excel_files

def process_excel(excel_file, excel_file_name):
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
    output_file_name = os.path.splitext(excel_file_name)[0] + '_output.xlsx'
    output_file_path = os.path.join('temp', output_file_name)
    df.to_excel(output_file_path, index=False)
    return output_file_name

st.title('File Converter')

os.makedirs('temp', exist_ok=True)

csv_files = st.file_uploader("Choose CSV Files", type="csv", accept_multiple_files=True)
excel_files = st.file_uploader("Choose Excel Files", type="xlsx", accept_multiple_files=True)

output_label = st.empty()
history_listbox = st.empty()

if csv_files:
    output_excel_files = process_files(csv_files)
    st.write('Conversion Log:')
    for file_path in output_excel_files:
        st.write(file_path)

output_excel_names = []  # Initialize the list before the conditional block

# Setelah loop yang memproses file Excel
if excel_files:
    output_excel_names = []
    for excel_file in excel_files:
        output_file_name = process_excel(excel_file, excel_file.name)
        output_excel_names.append(output_file_name)
    
    st.write('Conversion Log:')
    for file_name in output_excel_names:
        st.write(file_name)

    # Tambahkan tombol unduh untuk setiap file yang dihasilkan
    for output_file_name in output_excel_names:
        output_path = os.path.join('temp', output_file_name)
        if os.path.exists(output_path):
            st.download_button(f"Download {output_file_name}", output_path)

if st.button("Clear Log"):
    conversion_history.clear()

for item in conversion_history:
    st.write(item)

st.button("Print History to Word")
