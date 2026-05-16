import streamlit as st
import pdfplumber
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import re

st.set_page_config(page_title="Konverter BOSP Sempurna", page_icon="📊", layout="wide")
st.title("Konverter PDF Kertas Kerja BOSP (Anti-Bergeser) 📊")
st.write("Aplikasi ini menggunakan sistem **Smart Alignment (Pembacaan Kanan-Kiri)** untuk memastikan semua kolom lurus secara otomatis, tidak terpengaruh oleh format PDF yang berantakan.")

uploaded_file = st.file_uploader("Upload File PDF Kertas Kerja BOSP", type="pdf")

if uploaded_file is not None:
    st.info("Sedang mengekstrak dan menyusun tabel. Memastikan semua kolom lurus...")
    
    try:
        raw_rows = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    raw_rows.extend(table)
        
        if raw_rows:
            cleaned_data = []
            # Daftar kata pada baris judul / kop yang akan dihapus
            garbage_keywords = [
                "rincian kertas kerja", "tahun anggaran", "npsn", "nama sekolah", 
                "alamat", "kabupaten", "provinsi", "bulan", "sumber dana", 
                "penerimaan", "total penerimaan", "belanja", "kode rekening",
                "rincian perhitungan", "tarif harga", "no. urut"
            ]
            
            for row in raw_rows:
                if not row:
                    continue
                
                # 1. Kumpulkan hanya teks yang ada isinya, buang kolom kosong
                cells = [str(x).strip() for x in row if x is not None and str(x).strip() != ""]
                if not cells:
                    continue
                
                row_str_lower = " ".join(cells).lower()
                
                # Abaikan baris judul kertas kerja yang berulang
                if any(kw in row_str_lower for kw in garbage_keywords):
                    continue
                
                no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah = "", "", "", "", "", "", "", ""
                
                # Cek jika ini adalah Baris Keseluruhan "Jumlah" di paling bawah
                if "jumlah" in row_str_lower and len(cells) <= 3:
                    uraian = "Jumlah"
                    for c in reversed(cells):
                        if sum(char.isdigit() for char in c) > 3:
                            jumlah = c
                            break
                    cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
                    continue
                
                # --- SISTEM PEMBACAAN KIRI (Mengambil No dan Kode) ---
                # Cek No Urut (Berakhiran Titik)
                if re.match(r'^\d+\.$', cells[0]):
                    no_urut = cells.pop(0)
                
                # Cek Kode Rekening Belanja (Berawalan 5.1...)
                if cells and re.match(r'^5\.\d', cells[0]):
                    kode_rek = cells.pop(0)
                
                # Cek Kode Program (misal 03. 03.)
                prog_parts = []
                while cells and re.match(r'^(?:\d{2}\.\s*)+$', cells[0]):
                    prog_parts.append(cells.pop(0))
                if prog_parts:
                    kode_prog = " ".join(prog_parts)
                
                # Cek jika Kode Program terhubung langsung dengan Uraian ('07. 12. 03. Honorarium')
                if cells and re.match(r'^((?:\d{2}\.\s*)+)\s*(.+)$', cells[0]):
                    m = re.match(r'^((?:\d{2}\.\s*)+)\s*(.+)$', cells[0])
                    kode_prog += (" " if kode_prog else "") + m.group(1).strip()
                    if m.group(2).strip():
                        cells[0] = m.group(2).strip()
                    else:
                        cells.pop(0)
                
                if not cells:
                    continue
                    
                # --- SISTEM PEMBACAAN KANAN (Mengambil Angka dan Uraian agar tidak geser) ---
                if kode_rek:
                    # Jika ada Kode Rekening, sistem membaca mundur dari ujung kanan:
                    # [-1]=Jumlah, [-2]=Tarif, [-3]=Satuan, [-4]=Volume. Sisanya adalah Uraian.
                    if len(cells) >= 5:
                        uraian = " ".join(cells[:-4])
                        volume = cells[-4]
                        satuan = cells[-3]
                        tarif = cells[-2]
                        jumlah = cells[-1]
                    elif len(cells) == 4:
                        # Kasus langka: Volume dan Satuan menempel jadi satu kata ("20 hari")
                        m = re.match(r'^([\d\.,]+)\s+([a-zA-Z\s/]+)$', cells[-3])
                        if m:
                            uraian = " ".join(cells[:-3])
                            volume = m.group(1).strip()
                            satuan = m.group(2).strip()
                            tarif = cells[-2]
                            jumlah = cells[-1]
                        else:
                            uraian = " ".join(cells[:-2])
                            tarif = cells[-2]
                            jumlah = cells[-1]
                    else:
                        uraian = " ".join(cells)
                else:
                    # Jika baris Standar / Sub-Kategori (Hanya memiliki Uraian dan Jumlah)
                    if len(cells) >= 2:
                        uraian = " ".join(cells[:-1])
                        jumlah = cells[-1]
                    elif len(cells) == 1:
                        uraian = cells[0]
                
                # Masukkan data ke format akhir
                if any([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah]):
                    cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
            
            if cleaned_data:
                # Menampilkan Pratinjau di layar web
                df_preview = pd.DataFrame(cleaned_data, columns=["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"])
                st.subheader("👀 Review Hasil Tabel (Semua Kolom Kini Lurus Sempurna)")
                st.dataframe(df_preview, use_container_width=True)
                
                # Membuat Excel
                output = io.BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Rincian Kertas Kerja"
                
                # Membuat Header Rangkap
                ws.append(["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Rincian Perhitungan", "", "", "Jumlah"])
                ws.append(["", "", "", "", "Volume", "Satuan", "Tarif Harga", ""])
                
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
                        
                        # Ubah Format ke Angka Rupiah
                        if col_idx in [7, 8] and val_bersih:
                            try:
                                cell.value = float(val_bersih.replace('.', '').replace(',', ''))
                                cell.number_format = '#,##0'
                            except:
                                pass
                        elif col_idx in [5] and val_bersih:
                            try:
                                cell.value = int(val_bersih.replace('.', ''))
                            except:
                                pass
                                
                        # Cetak Tebal Baris Kategori
                        if not row_data[1] and row_data[0] != "Jumlah" and row_data[3].lower() != "jumlah":
                            cell.font = bold_font
                            
                        # Rata Tengah Kolom Khusus
                        if col_idx in [1, 2, 3, 5, 6]:
                            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                
                # Mengatur Lebar Kolom
                ws.column_dimensions['A'].width = 6
                ws.column_dimensions['B'].width = 18
                ws.column_dimensions['C'].width = 12
                ws.column_dimensions['D'].width = 50
                ws.column_dimensions['E'].width = 8
                ws.column_dimensions['F'].width = 12
                ws.column_dimensions['G'].width = 12
                ws.column_dimensions['H'].width = 15
                
                # Tebalkan Total "Jumlah" Otomatis
                for r_idx in range(3, ws.max_row + 1):
                    cell_d = ws.cell(row=r_idx, column=4)
                    if cell_d.value and "jumlah" in str(cell_d.value).lower():
                        cell_d.font = bold_font
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
                        ws.cell(row=r_idx, column=8).font = bold_font
                
                wb.save(output)
                
                st.success("✅ Tabel selesai diurutkan! Silakan cek di bagian Review.")
                st.download_button("📥 Download Excel Final", output.getvalue(), "Kertas_Kerja_BOSP_Tabel_Rapi.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("Data tabel kosong. Pastikan PDF yang di-upload adalah Kertas Kerja yang benar.")
        else:
            st.warning("Gagal membaca isi tabel pada PDF.")
    except Exception as e:
        st.error(f"Error pada sistem: {e}")
