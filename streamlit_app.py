import streamlit as st
import os
import re  # Import library RegEx
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
from tempfile import NamedTemporaryFile

# Path folder xlsx untuk menyimpan file terjemahan
output_folder_path = r"temp"

# Membuat database kode mata pelajaran
database_mata_pelajaran = {
    # Database kode mata pelajaran
}

# Fungsi untuk menerjemahkan jadwal
def translate_jadwal(file_path):
    # Membaca jadwal dari file Excel menggunakan openpyxl
    wb = load_workbook(file_path)
    sheet = wb.active

    # Menerjemahkan kode mata pelajaran dengan RegEx
    kode_guru_pattern = r'\d+'  # Pola RegEx untuk mencocokkan angka dalam teks

    for row in sheet.iter_rows(min_row=2):
        for index, cell in enumerate(row):
            if sheet.cell(row=1, column=cell.column).value == "Jam":
                # Kolom "Jam" tidak diterjemahkan
                continue
            cell_value = cell.value

            # Menggunakan RegEx untuk mencari kode guru dalam teks
            kode_guru_list = re.findall(kode_guru_pattern, cell_value)

            # Menerjemahkan kode guru yang ditemukan
            mata_pelajaran = []
            for kode_guru in kode_guru_list:
                if kode_guru in database_mata_pelajaran:
                    mata_pelajaran.append(database_mata_pelajaran[kode_guru])

            # Menggabungkan hasil terjemahan ke dalam teks sel
            cell.value = ', '.join(mata_pelajaran)

            # Mengatur format sel agar rapi
            max_length = len(cell.value)
            adjusted_width = (max_length + 2) * 1.2
            column_letter = get_column_letter(cell.column)
            sheet.column_dimensions[column_letter].width = adjusted_width
            cell.alignment = Alignment(wrapText=True)

    # Menyimpan jadwal terjemahan ke file Excel sementara
    temp_file = NamedTemporaryFile(delete=False)
    wb.save(temp_file.name)
    temp_file.close()

    return temp_file.name

# Judul halaman Streamlit
st.title("Jadwal mansa translator")
st.write("Konversi file pdf jadwal terlebih dahulu menjadi excel melalui layanan online yang banyak tersedia, "
         "setelah itu baru masukkan ke dalam website ini")

# Pengguna memilih file Excel untuk diterjemahkan
uploaded_file = st.file_uploader("Unggah File Excel", type=["xlsx"])

# Jika file Excel diunggah
if uploaded_file is not None:
    # Menyimpan file Excel sementara
    with open(os.path.join(output_folder_path, uploaded_file.name), "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Menjalankan terjemahan jadwal
    translated_file_path = translate_jadwal(os.path.join(output_folder_path, uploaded_file.name))

    # Pengguna dapat mengunduh file terjemahan
    st.download_button("Unduh Hasil Terjemahan", data=open(translated_file_path, "rb").read(), file_name=uploaded_file.name)

    # Menghapus file sementara setelah diunduh
    os.remove(translated_file_path)
