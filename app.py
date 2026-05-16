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
st.write("Silakan upload file PDF Kertas Kerja. Aplikasi akan menyelaraskan kolom, menghapus judul berulang, dan menampilkan review sebelum Anda mengunduhnya.")

uploaded_file = st.file_uploader("Upload File PDF Kertas Kerja BOSP", type="pdf")

if uploaded_file is not None:
    st.info("Sedang menyelaraskan kolom dan merapikan data tabel...")
    
    try:
        all_data = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)
        
        if all_data:
            # 1. KONVERSI KE DATAFRAME TERLEBIH DAHULU
            # Ini adalah kunci agar kolom tidak bergeser (Alignment Grid Otomatis)
            df = pd.DataFrame(all_data)
            df = df.fillna("")
            
            # Pastikan minimal ada 10 kolom untuk pemetaan yang tepat
            while df.shape[1] < 10:
                df[df.shape[1]] = ""
                
            cleaned_data = []
            
            # Kata kunci Kop Surat / Header berulang yang harus dihapus
            garbage_keywords = [
                "rincian kertas kerja", "tahun anggaran", "npsn", "nama sekolah", 
                "alamat", "kabupaten", "provinsi", "bulan", "sumber dana", 
                "penerimaan", "total penerimaan", "belanja", "no. urut", "kode rekening",
                "rincian perhitungan", "tarif harga"
            ]
            
            # 2. PROSES PEMETAAN BARIS YANG SUDAH SEJAJAR
            for index, row in df.iterrows():
                # Ambil 10 kolom pertama dari grid yang sudah lurus
                r = [str(row[c]).strip() if c in df.columns else "" for c in range(10)]
                
                row_str = " ".join(r).lower()
                
                # Lewati baris kosong
                if not row_str.strip():
                    continue
                
                # Lewati baris Kop Surat / Judul Halaman
                if any(kw in row_str for kw in garbage_keywords):
                    continue
                
                # PEMETAAN KOLOM (Pasti lurus karena sudah dikunci oleh DataFrame)
                no_urut = r[0]
                kode_rek = r[1]
                
                # Kode Program digabung karena kadang terpisah di 3 kolom (2, 3, 4)
                kode_prog = f"{r[2]} {r[3]} {r[4]}".strip()
                kode_prog = re.sub(r'\s+', ' ', kode_prog)
                
                uraian = r[5]
                volume = r[6]
                satuan = r[7]
                tarif = r[8]
                jumlah = r[9]
                
                # Deteksi jika ini adalah baris "Jumlah" (Total paling bawah) 
                if "jumlah" in row_str and not uraian and not jumlah:
                    uraian = "Jumlah"
                    # Cari angkanya dari kanan ke kiri
                    for idx in range(9, -1, -1):
                        if r[idx] and sum(c.isdigit() for c in r[idx]) > 3:
                            jumlah = r[idx]
                            break
                            
                # Pastikan tidak menginput baris yang nyaris kosong
                if not any([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah]):
                    continue
                    
                cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
            
            if cleaned_data:
                # 3. FITUR REVIEW (Pratinjau)
                kolom_tabel = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"]
                df_preview = pd.DataFrame(cleaned_data, columns=kolom_tabel)
                
                st.subheader("👀 Review Hasil Tabel (Pastikan Kolom Sudah Lurus)")
                st.dataframe(df_preview, use_container_width=True)
                
                # 4. BUAT EXCEL FINAL SESUAI FORMAT YANG DIMINTA
                output = io.BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Rincian Kertas Kerja"
                
                headers1 = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Rincian Perhitungan", "", "", "Jumlah"]
                headers2 = ["", "", "", "", "Volume", "Satuan", "Tarif Harga", ""]
                
                ws.append(headers1)
                ws.append(headers2)
                
                ws.merge_cells('A1:A2')
                ws.merge_cells('B1:B2')
                ws.merge_cells('C1:C2')
                ws.merge_cells('D1:D2')
                ws.merge_cells('E1:G1')
                ws.merge_cells('H1:H2')
                
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                bold_font = Font(bold=True)
                
                for r_idx in range(1, 3):
                    for c_idx in range(1, 9):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        cell.font = bold_font
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.border = thin_border
                
                for row_idx, row_data in enumerate(cleaned_data, start=3):
                    for col_idx, val_str in enumerate(row_data, start=1):
                        val_bersih = str(val_str).replace('\n', ' ').strip()
                        cell = ws.cell(row=row_idx, column=col_idx, value=val_bersih)
                        cell.border = thin_border
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                        
                        # Format Angka Otomatis
                        if col_idx in [7, 8]:
                            if val_bersih:
                                try:
                                    num_val = float(val_bersih.replace('.', '').replace(',', ''))
                                    cell.value = num_val
                                    cell.number_format = '#,##0'
                                except:
                                    pass
                        elif col_idx in [5]:
                            if val_bersih:
                                try:
                                    num_val = int(val_bersih)
                                    cell.value = num_val
                                except:
                                    pass
                                    
                        # Cetak Tebal (Bold) untuk Baris Kategori
                        if row_data[1] == "" and row_data[0] != "" and row_data[0].lower() != "jumlah":
                            cell.font = bold_font
                            
                        # Rata Tengah
                        if col_idx in [1, 2, 3, 5, 6]:
                            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                            
                ws.column_dimensions['A'].width = 6
                ws.column_dimensions['B'].width = 18
                ws.column_dimensions['C'].width = 12
                ws.column_dimensions['D'].width = 50
                ws.column_dimensions['E'].width = 8
                ws.column_dimensions['F'].width = 12
                ws.column_dimensions['G'].width = 12
                ws.column_dimensions['H'].width = 15
                
                # Cetak Tebal Baris "Jumlah" Paling Bawah
                for r_idx in range(3, ws.max_row + 1):
                    cell_d = ws.cell(row=r_idx, column=4)
                    if cell_d.value and "jumlah" in str(cell_d.value).lower():
                        cell_d.font = bold_font
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
                        ws.cell(row=r_idx, column=8).font = bold_font
                
                wb.save(output)
                excel_data = output.getvalue()
                
                st.success("✅ Tabel sudah lurus dan rapi! Silakan cek review di atas dan Download Excel di bawah.")
                
                # Tombol Download
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
