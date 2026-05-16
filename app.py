import streamlit as st
import pdfplumber
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# Konfigurasi Halaman Web
st.set_page_config(page_title="Konverter Kertas Kerja BOSP", page_icon="📊")
st.title("Konverter PDF Kertas Kerja ke Excel Final 📊")
st.write("Upload file PDF Kertas Kerja Anda. Aplikasi ini akan otomatis membaca, membersihkan kolom yang berantakan, dan menyusunnya menjadi tabel Excel yang rapi siap cetak/edit.")

# Fitur Upload File
uploaded_file = st.file_uploader("Pilih file PDF BOSP", type="pdf")

if uploaded_file is not None:
    st.info("Sedang memproses dan merapikan dokumen. Mohon tunggu sebentar...")
    
    try:
        all_data = []
        # 1. Membaca PDF
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)
        
        if all_data:
            # 2. PROSES PEMBERSIHAN DATA (Membersihkan kolom yang berantakan)
            df = pd.DataFrame(all_data)
            
            # Membuang baris header bawaan PDF yang rusak (biasanya 2 baris pertama)
            df_data = df.iloc[2:].copy()
            
            # Memastikan jumlah kolom minimal 10 untuk menghindari error
            if df_data.shape[1] > 10:
                df_data = df_data.iloc[:, :10]
            elif df_data.shape[1] < 10:
                for i in range(df_data.shape[1], 10):
                    df_data[i] = ""
                    
            df_data = df_data.fillna("")
            
            # Menggabungkan Kode Program yang sering terpisah-pisah di kolom 2, 3, dan 4
            df_data['Kode_Program'] = df_data[2].astype(str).str.strip() + " " + df_data[3].astype(str).str.strip() + " " + df_data[4].astype(str).str.strip()
            df_data['Kode_Program'] = df_data['Kode_Program'].str.replace(r'\s+', ' ', regex=True).str.strip()
            
            # Mengambil hanya kolom yang benar (No, Rekening, Kode Program, Uraian, Vol, Satuan, Harga, Jumlah)
            cleaned_data = df_data[[0, 1, 'Kode_Program', 5, 6, 7, 8, 9]].values.tolist()

            # 3. MEMBUAT EXCEL FINAL YANG RAPI (Seperti Kertas_Kerja_BOSP_Tabel_Saja.xlsx)
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Kertas Kerja"

            # Menulis Header Tabel
            headers1 = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Rincian Perhitungan", "", "", "Jumlah"]
            headers2 = ["", "", "", "", "Volume", "Satuan", "Tarif Harga", ""]

            ws.append(headers1)
            ws.append(headers2)

            # Menggabungkan Sel Header (Merge Cells)
            ws.merge_cells('A1:A2')
            ws.merge_cells('B1:B2')
            ws.merge_cells('C1:C2')
            ws.merge_cells('D1:D2')
            ws.merge_cells('E1:G1') # Merge Rincian Perhitungan
            ws.merge_cells('H1:H2')

            # Pengaturan Gaya Tabel (Garis & Huruf Tebal)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            bold_font = Font(bold=True)

            # Menerapkan gaya ke Header
            for row in range(1, 3):
                for col in range(1, 9):
                    cell = ws.cell(row=row, column=col)
                    cell.font = bold_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = thin_border

            # Memasukkan Data Bersih ke dalam Excel
            for row_idx, row_data in enumerate(cleaned_data, start=3):
                for col_idx, val in enumerate(row_data, start=1):
                    # Bersihkan teks dari baris baru (enter)
                    val_str = str(val).replace('\n', ' ').strip() if val is not None else ""
                    
                    cell = ws.cell(row=row_idx, column=col_idx, value=val_str)
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    
                    # Ubah tipe data Harga & Jumlah menjadi Angka (Format Uang)
                    if col_idx in [7, 8]:
                        if val_str and val_str != "":
                            try:
                                num_val = float(val_str.replace('.', ''))
                                cell.value = num_val
                                cell.number_format = '#,##0'
                            except:
                                pass
                    # Ubah tipe data Volume menjadi Angka Bulat
                    elif col_idx in [5]:
                        if val_str and val_str != "":
                            try:
                                num_val = int(val_str)
                                cell.value = num_val
                            except:
                                pass

                    # Menebalkan baris Kategori (Jika Kode Rekening kosong tapi ada No/Uraian)
                    if row_data[1] == "" and row_data[0] != "": 
                        cell.font = bold_font
                        
                    # Rata Tengah untuk kolom No, Kode, Vol, Satuan
                    if col_idx in [1, 2, 3, 5, 6]:
                        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

            # Mengatur Lebar Kolom agar pas
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 18
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 50
            ws.column_dimensions['E'].width = 8
            ws.column_dimensions['F'].width = 12
            ws.column_dimensions['G'].width = 12
            ws.column_dimensions['H'].width = 15

            # Mencari baris "Jumlah" paling bawah untuk ditebalkan otomatis
            for row_idx in range(3, ws.max_row + 1):
                cell_d = ws.cell(row=row_idx, column=4)
                if cell_d.value and "Jumlah" in str(cell_d.value):
                    cell_d.font = bold_font
                    cell_d.alignment = Alignment(horizontal="right", vertical="center")
                    ws.cell(row=row_idx, column=8).font = bold_font

            # Menyimpan file ke memori (Output Akhir)
            wb.save(output)
            processed_data = output.getvalue()
            
            st.success("✅ Berhasil! Data yang berantakan telah dirapikan menjadi format tabel Excel final.")
            
            # Tombol Download
            st.download_button(
                label="📥 Unduh Excel Final (Rapi)",
                data=processed_data,
                file_name="Kertas_Kerja_BOSP_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Tidak ada tabel yang ditemukan dalam PDF ini.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat merapikan file: {e}")
