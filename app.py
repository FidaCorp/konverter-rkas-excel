import streamlit as st
import pdfplumber
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import re

# Konfigurasi Tampilan Aplikasi Web
st.set_page_config(page_title="Konverter BOSP Rapi", page_icon="📊", layout="wide")
st.title("Konverter PDF Kertas Kerja BOSP (Tabel Saja) 📊")
st.write("Silakan upload file PDF Kertas Kerja. Aplikasi akan otomatis menghapus judul halaman yang berulang, merapikan kolom, dan menampilkan review sebelum Anda mengunduhnya.")

# Fitur Unggah Dokumen PDF
uploaded_file = st.file_uploader("Upload File PDF Kertas Kerja BOSP", type="pdf")

if uploaded_file is not None:
    st.info("Sedang memproses, menyaring, dan merapikan data tabel...")
    
    try:
        raw_rows = []
        # 1. Ekstraksi Data Kasar dari PDF
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    raw_rows.extend(table)
        
        if raw_rows:
            cleaned_data = []
            
            # Kata kunci sampah atau judul halaman berulang yang harus dihilangkan total
            garbage_keywords = [
                "rincian kertas kerja", "tahun anggaran", "npsn", "nama sekolah", 
                "alamat", "kabupaten", "provinsi", "bulan", "sumber dana", 
                "penerimaan", "total penerimaan", "belanja", "no. urut", "kode rekening",
                "rincian perhitungan", "tarif harga", "no. urut"
            ]
            
            # 2. Proses Pembersihan dan Penyelarasan Kolom
            for row in raw_rows:
                if not row:
                    continue
                
                # Menggabungkan seluruh baris menjadi satu teks kecil untuk pengecekan kata kunci sampah
                row_str = " ".join([str(x) for x in row if x is not None]).lower()
                
                # JIKA baris mengandung kata kunci judul/kop, langsung dilewati (dihapus)
                if any(kw in row_str for kw in garbage_keywords):
                    continue
                
                # Pastikan jumlah kolom minimal ada 10 agar pemetaan indeks tidak bergeser
                if len(row) < 10:
                    row = row + [""] * (10 - len(row))
                
                # Hilangkan nilai None/kosong menjadi teks bersih
                row = ["" if x is None else str(x).strip() for x in row]
                
                no_urut = row[0]
                kode_rek = row[1]
                
                # Menggabungkan Kode Program dari pembacaan kolom indeks 2, 3, dan 4 yang berantakan
                kode_prog = f"{row[2]} {row[3]} {row[4]}".strip()
                kode_prog = re.sub(r'\s+', ' ', kode_prog) # Rapikan spasi ganda
                
                uraian = row[5]
                volume = row[6]
                satuan = row[7]
                tarif = row[8]
                jumlah = row[9]
                
                # Skip jika baris benar-benar kosong total tanpa ada isi apa pun
                if not any([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah]):
                    continue
                
                # Skip jika baris sisa yang tidak berguna ikut terbaca
                if uraian.lower() in ["uraian", ""] and jumlah == "" and kode_rek == "":
                    continue
                
                # Masukkan data yang sudah lurus dan bersih ke list final
                cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
            
            if cleaned_data:
                # 3. MEMBUAT DATAFRAME UNTUK FITUR REVIEW (PRATINJAU) Di LAYAR
                kolom_tabel = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"]
                df_preview = pd.DataFrame(cleaned_data, columns=kolom_tabel)
                
                st.subheader("👀 Review Hasil Tabel (Silakan Cek Terlebih Dahulu)")
                st.dataframe(df_preview, use_container_width=True)
                
                # 4. PROSES GENERATE FILE EXCEL FINAL (Sama seperti Kertas_Kerja_BOSP_Tabel_Saja.xlsx)
                output = io.BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Rincian Kertas Kerja"
                
                # Susun struktur kepala tabel rangkap (Merge header)
                headers1 = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Rincian Perhitungan", "", "", "Jumlah"]
                headers2 = ["", "", "", "", "Volume", "Satuan", "Tarif Harga", ""]
                
                ws.append(headers1)
                ws.append(headers2)
                
                ws.merge_cells('A1:A2')
                ws.merge_cells('B1:B2')
                ws.merge_cells('C1:C2')
                ws.merge_cells('D1:D2')
                ws.merge_cells('E1:G1') # Gabungkan judul Rincian Perhitungan
                ws.merge_cells('H1:H2')
                
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                bold_font = Font(bold=True)
                
                # Desain style untuk Kepala Tabel (Header)
                for r in range(1, 3):
                    for c in range(1, 9):
                        cell = ws.cell(row=r, column=c)
                        cell.font = bold_font
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.border = thin_border
                
                # Memasukkan isi data ke lembar Excel
                for row_idx, row_data in enumerate(cleaned_data, start=3):
                    for col_idx, val_str in enumerate(row_data, start=1):
                        # Hilangkan enter di dalam sel agar teks lurus satu baris
                        val_bersih = str(val_str).replace('\n', ' ').strip()
                        
                        cell = ws.cell(row=row_idx, column=col_idx, value=val_bersih)
                        cell.border = thin_border
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                        
                        # Format Angka Tarif Harga & Jumlah Uang otomatis (.000)
                        if col_idx in [7, 8]:
                            if val_bersih and val_bersih != "":
                                try:
                                    num_val = float(val_bersih.replace('.', '').replace(',', ''))
                                    cell.value = num_val
                                    cell.number_format = '#,##0'
                                ] except:
                                    pass
                        # Format Angka Volume
                        elif col_idx in [5]:
                            if val_bersih and val_bersih != "":
                                try:
                                    num_val = int(val_bersih)
                                    cell.value = num_val
                                except:
                                    pass
                                    
                        # Menebalkan otomatis baris Kategori Standar/Kegiatan utama
                        if row_data[1] == "" and row_data[0] != "":
                            cell.font = bold_font
                            
                        # Format rata tengah untuk nomor dan kode program
                        if col_idx in [1, 2, 3, 5, 6]:
                            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                            
                # Mengatur Ukuran Lebar Kolom Excel secara presisi
                ws.column_dimensions['A'].width = 6
                ws.column_dimensions['B'].width = 18
                ws.column_dimensions['C'].width = 12
                ws.column_dimensions['D'].width = 50
                ws.column_dimensions['E'].width = 8
                ws.column_dimensions['F'].width = 12
                ws.column_dimensions['G'].width = 12
                ws.column_dimensions['H'].width = 15
                
                # Deteksi otomatis baris Total "Jumlah" paling bawah untuk ditebalkan hurufnya
                for r_idx in range(3, ws.max_row + 1):
                    cell_d = ws.cell(row=r_idx, column=4)
                    if cell_d.value and "Jumlah" in str(cell_d.value):
                        cell_d.font = bold_font
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
                        ws.cell(row=r_idx, column=8).font = bold_font
                
                # Simpan workbook Excel ke dalam memori aplikasi
                wb.save(output)
                excel_data = output.getvalue()
                
                st.success("✅ Pembersihan sukses! Silakan lihat preview di atas sebelum men-download.")
                
                # 5. TOMBOL DOWNLOAD DI TARUH DI PALING BAWAH SETELAH REVIEW
                st.download_button(
                    label="📥 Download Excel Final (Tabel Rapi)",
                    data=excel_data,
                    file_name="Kertas_Kerja_BOSP_Tabel_Saja.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Data tabel gagal diekstrak secara bersih.")
        else:
            st.warning("Format tabel di dalam file PDF tidak terdeteksi.")
            
    except Exception as e:
        st.error(f"Terjadi error saat merapikan data: {e}")
