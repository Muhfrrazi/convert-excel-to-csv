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
    output_excel_files = process_files(csv_files)  # Get the output Excel files

    st.write('Conversion Log:')
    for file_path in output_excel_files:
        st.write(file_path)

    # Tambahkan tombol unduh dengan tipe file Excel (xlsx) yang sesuai
    for output_file_path in output_excel_files:
        if os.path.exists(output_file_path):
            output_file_name = os.path.basename(output_file_path)
            st.download_button(f"Download {output_file_name}", output_file_path, file_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


output_excel_names = []  # Initialize the list before the conditional block
conversion_history = []

# Setelah loop yang memproses file Excel
if excel_files:
    output_excel_files = []  # Initialize list to store output Excel file paths
    
    for excel_file in excel_files:
        output_file_name = process_excel(excel_file, excel_file.name)
        output_excel_files.append(output_file_name)
    
    st.write('Conversion Log:')
    for file_name in output_excel_files:
        st.write(file_name)

    # Tambahkan tombol unduh dengan tipe file Excel (xlsx) yang sesuai
    # Inside the loop for processing Excel files
for output_file_name in output_excel_files:
    output_path = os.path.join('temp', output_file_name)
    if os.path.exists(output_path):
        output_file_name = os.path.basename(output_file_path)
        st.download_button(
            label=f"Download {output_file_name}",
            data=output_path,
            file_name=output_file_name)


# After the loops, you can use the conversion history
st.write('Conversion Log:')
for item in conversion_history:
    st.write(item)

if st.button("Clear Log"):
    conversion_history.clear()

for item in conversion_history:
    st.write(item)

st.button("Print History to Word")
