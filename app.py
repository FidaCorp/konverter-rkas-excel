import streamlit as st
import pdfplumber
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import re

# Konfigurasi Tampilan Aplikasi Web (Wide Mode agar Review terlihat jelas)
st.set_page_config(page_title="Konverter BOSP Sempurna", page_icon="📊", layout="wide")
st.title("Konverter PDF Kertas Kerja BOSP (Anti-Bergeser) 📊")
st.write("Aplikasi ini menggunakan **Smart Column Alignment** untuk mendeteksi data yang menyatu (seperti pada Tenaga Administrasi Harian Lepas) dan mengembalikannya ke kolom yang benar secara otomatis.")

# Fitur Unggah Dokumen PDF
uploaded_file = st.file_uploader("Upload File PDF Kertas Kerja BOSP", type="pdf")

if uploaded_file is not None:
    st.info("Sedang menganalisis struktur dokumen dan menyelaraskan kolom...")
    
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
            
            # Kata kunci Kop Surat / Header Halaman yang wajib dibuang total agar tidak mengulang
            garbage_keywords = [
                "rincian kertas kerja", "tahun anggaran", "npsn", "nama sekolah", 
                "alamat", "kabupaten", "provinsi", "bulan", "sumber dana", 
                "penerimaan", "total penerimaan", "belanja", "no. urut", "kode rekening",
                "rincian perhitungan", "tarif harga", "volume", "satuan"
            ]
            
            # 2. PROSES DETEKSI & PENYELARASAN KOLOM PINTAR (ANTI-BERGESER)
            for row in raw_rows:
                if not row:
                    continue
                
                # Ambil semua teks yang tidak kosong dalam satu baris
                cells = [str(x).strip() for x in row if x is not None and str(x).strip() != ""]
                if not cells:
                    continue
                
                # Gabungkan untuk cek baris sampah/judul halaman
                row_str_lower = " ".join(cells).lower()
                if any(kw in row_str_lower for kw in garbage_keywords):
                    continue
                
                # Inisialisasi kolom standar
                no_urut = ""
                kode_rek = ""
                kode_prog = ""
                uraian = ""
                volume = ""
                satuan = ""
                tarif = ""
                jumlah = ""
                
                # Cek apakah baris ini memiliki Kode Rekening Belanja (dimulai dengan angka 5)
                has_kode_rek = any(re.match(r'^5\.\d', c) for c in cells)
                
                if has_kode_rek:
                    # BARIS DETAIL BELANJA (Misal: Honorarium, Semen, Tenaga Administrasi)
                    if re.match(r'^5\.\d', cells[0]):
                        no_urut = ""
                        kode_rek = cells[0]
                        rem_cells = cells[1:]
                    else:
                        no_urut = cells[0]
                        kode_rek = cells[1]
                        rem_cells = cells[2:]
                    
                    # Kasus A: Kolom terpisah sempurna dari PDF (6 elemen tersisa)
                    if len(rem_cells) == 6:
                        kode_prog = rem_cells[0]
                        uraian = rem_cells[1]
                        volume = rem_cells[2]
                        satuan = rem_cells[3]
                        tarif = rem_cells[4]
                        jumlah = rem_cells[5]
                        
                    # Kasus B: Kolom menyatu/bergeser akibat PDF tanpa garis (4 elemen tersisa)
                    # Contoh kasus: ['07. 12. 03. Tenaga Administrasi Harian Lepas', '20 hari', '35.000', '700.000']
                    elif len(rem_cells) == 4:
                        # Pisahkan Kode Program & Uraian menggunakan Regex
                        prog_uraian = rem_cells[0]
                        match_p = re.match(r'^((?:\d{2}\.\s*)+)\s*(.*)$', prog_uraian)
                        if match_p:
                            kode_prog = match_p.group(1).strip()
                            uraian = match_p.group(2).strip()
                        else:
                            kode_prog = ""
                            uraian = prog_uraian
                        
                        # Pisahkan Volume & Satuan
                        vol_sat = rem_cells[1]
                        match_v = re.match(r'^([\d\.,]+)\s*(.*)$', vol_sat)
                        if match_v:
                            volume = match_v.group(1).strip()
                            satuan = match_v.group(2).strip()
                        else:
                            volume = ""
                            satuan = vol_sat
                            
                        # Tarif dan Jumlah otomatis diisi dari sisa kolom kanan
                        tarif = rem_cells[2]
                        jumlah = rem_cells[3]
                    
                    # Kasus C: Cadangan Fleksibel jika panjang kolom acak lainnya
                    else:
                        for c in rem_cells:
                            if re.match(r'^(?:\d{2}\.\s*)+$', c):
                                kode_prog = c
                            elif re.match(r'^((?:\d{2}\.\s*)+)\s+(.+)$', c):
                                m = re.match(r'^((?:\d{2}\.\s*)+)\s+(.+)$', c)
                                kode_prog = m.group(1).strip()
                                uraian = m.group(2).strip()
                            elif re.match(r'^(\d+)\s+([a-zA-Z_/]+)$', c):
                                m = re.match(r'^(\d+)\s+([a-zA-Z_/]+)$', c)
                                volume = m.group(1).strip()
                                      satuan = m.group(2).strip()
                            elif c.lower() in ["hari", "orang/hari", "biji", "buah", "rim", "zak", "kg", "m3", "kotak", "bulan", "kwh", "orang / hari"]:
                                satuan = c
                            elif sum(char.isdigit() for char in c) > 0 and not kode_prog:
                                if c.isdigit() and int(c) < 1000 and not volume:
                                    volume = c
                                elif not tarif:
                                    tarif = c
                                else:
                                    jumlah = c
                            else:
                                if not uraian:
                                    uraian = c
                                else:
                                    uraian += " " + c
                else:
                    # BARIS SUB-HEADER UTAMA / KATEGORI / TOTAL JUMLAH
                    if "jumlah" in row_str_lower:
                        no_urut = ""
                        kode_rek = ""
                        kode_prog = ""
                        uraian = "Jumlah"
                        for c in reversed(cells):
                            if sum(char.isdigit() for char in c) > 3:
                                jumlah = c
                                break
                    else:
                        if len(cells) >= 3:
                            if re.match(r'^\d+\.$', cells[0]):
                                no_urut = cells[0]
                                kode_prog = cells[1]
                                uraian = cells[2]
                                if len(cells) >= 4:
                                    jumlah = cells[3]
                            else:
                                no_urut = ""
                                kode_prog = cells[0]
                                uraian = cells[1]
                                if len(cells) >= 3:
                                    jumlah = cells[2]
                        elif len(cells) == 2:
                            if re.match(r'^(?:\d{2}\.\s*)+$', cells[0]) or re.match(r'^\d+$', cells[0]):
                                kode_prog = cells[0]
                                uraian = cells[1]
                            else:
                                uraian = cells[0]
                                jumlah = cells[1]
                                
                # Pastikan baris kosong tidak masuk ke daftar
                if not any([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah]):
                    continue
                
                cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
            
            if cleaned_data:
                # 3. FITUR REVIEW (PRATINJAU) DI LAYAR UNTUK DICEK USER
                kolom_tabel = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"]
                df_preview = pd.DataFrame(cleaned_data, columns=kolom_tabel)
                
                st.subheader("👀 Review Hasil Tabel (Silakan periksa, kolom sekarang sudah lurus sempurna)")
                st.dataframe(df_preview, use_container_width=True)
                
                # 4. PROSES BUAT EXCEL SEPERTI Kertas_Kerja_BOSP_Tabel_Saja.xlsx
                output = io.BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Rincian Kertas Kerja"
                
                # Struktur Kepala Tabel Rangkap
                headers1 = ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Rincian Perhitungan", "", "", "Jumlah"]
                headers2 = ["", "", "", "", "Volume", "Satuan", "Tarif Harga", ""]
                
                ws.append(headers1)
                ws.append(headers2)
                
                # Gabungkan Sel Header
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
                
                # Tulis data bersih ke lembar Excel
                for row_idx, row_data in enumerate(cleaned_data, start=3):
                    for col_idx, val_str in enumerate(row_data, start=1):
                        val_bersih = str(val_str).replace('\n', ' ').strip()
                        cell = ws.cell(row=row_idx, column=col_idx, value=val_bersih)
                        cell.border = thin_border
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                        
                        # Format Angka Tarif & Jumlah (.000)
                        if col_idx in [7, 8]:
                            if val_bersih:
                                try:
                                    num_val = float(val_bersih.replace('.', '').replace(',', ''))
                                    cell.value = num_val
                                    cell.number_format = '#,##0'
                                except:
                                    pass
                        # Format Angka Volume
                        elif col_idx in [5]:
                            if val_bersih:
                                try:
                                    num_val = int(val_bersih)
                                    cell.value = num_val
                                except:
                                    pass
                                    
                        # Cetak Tebal Baris Kategori Utama
                        if row_data[1] == "" and row_data[0] != "" and row_data[0].lower() != "jumlah":
                            cell.font = bold_font
                            
                        # Rata Tengah kolom kode & nomor
                        if col_idx in [1, 2, 3, 5, 6]:
                            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                            
                # Pengaturan Lebar Kolom Standar Rapi
                ws.column_dimensions['A'].width = 6
                ws.column_dimensions['B'].width = 18
                ws.column_dimensions['C'].width = 12
                ws.column_dimensions['D'].width = 50
                ws.column_dimensions['E'].width = 8
                ws.column_dimensions['F'].width = 12
                ws.column_dimensions['G'].width = 12
                ws.column_dimensions['H'].width = 15
                
                # Cetak Tebal Otomatis Baris "Jumlah" Paling Bawah
                for r_idx in range(3, ws.max_row + 1):
                    cell_d = ws.cell(row=r_idx, column=4)
                    if cell_d.value and "jumlah" in str(cell_d.value).lower():
                        cell_d.font = bold_font
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
                        ws.cell(row=r_idx, column=8).font = bold_font
                
                wb.save(output)
                excel_data = output.getvalue()
                
                st.success("✅ Penyelarasan selesai! Semua data sudah lurus di kolomnya masing-masing.")
                
                # Tombol Download ditaruh di paling bawah setelah kolom Review
                st.download_button(
                    label="📥 Download Excel Final (Tabel Rapi)",
                    data=excel_data,
                    file_name="Kertas_Kerja_BOSP_Tabel_Saja.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Data gagal dibersihkan.")
        else:
            st.warning("Tidak ada tabel terdeteksi di PDF.")
            
    except Exception as e:
        st.error(f"Terjadi error: {e}")
