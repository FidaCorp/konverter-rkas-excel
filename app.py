import streamlit as st
import pdfplumber
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import re

# ==========================================
# 1. KONFIGURASI TAMPILAN HALAMAN
# ==========================================
st.set_page_config(page_title="Konverter RKAS BOSP", page_icon="🏫", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 2.8rem; font-weight: 800; color: #1E88E5; margin-bottom: 0rem; }
    .sub-title { font-size: 1.2rem; color: #555555; margin-bottom: 2rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 2rem; }
    .stTabs [data-baseweb="tab"] { height: 3rem; white-space: pre-wrap; font-size: 1.1rem; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SIDEBAR (PANEL KIRI)
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135692.png", width=80) 
    st.title("Panduan Sistem")
    st.markdown("""
    Pilih jenis dokumen yang ingin Anda rapikan:
    
    📄 **MODE PDF (Rekomendasi):**
    Ubah PDF Kertas Kerja BOSP langsung menjadi format Excel rapi.
    
    📊 **MODE CSV/EXCEL KOTOR:**
    Rapikan file Excel/CSV (hasil *convert* aplikasi lain) yang tata letaknya bergeser dan berantakan.
    """)
    st.info("💡 **Tips:** Sistem akan otomatis menghapus teks Kop Surat yang berulang dan menyejajarkan kembali nilai tarif dan volumenya.")

# ==========================================
# 3. FUNGSI MERAPIKAN EXCEL (CORE FUNCTION)
# ==========================================
def create_styled_excel(cleaned_data):
    """Fungsi standar untuk membuat desain Excel yang bagus"""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rincian Kertas Kerja"
    
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
    header_fill = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid") 
    
    for r_idx in range(1, 3):
        for c_idx in range(1, 9):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = bold_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
    
    for row_idx, row_data in enumerate(cleaned_data, start=3):
        for col_idx, val_str in enumerate(row_data, start=1):
            val_bersih = str(val_str).replace('\n', ' ').strip()
            cell = ws.cell(row=row_idx, column=col_idx, value=val_bersih)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            
            # Format Uang (.000)
            if col_idx in [7, 8] and val_bersih:
                try:
                    cell.value = float(val_bersih.replace('.', '').replace(',', ''))
                    cell.number_format = '#,##0'
                except: pass
            # Format Volume
            elif col_idx in [5] and val_bersih:
                try: cell.value = float(val_bersih.replace(',', '.'))
                except: pass
                    
            if not row_data[1] and row_data[0] != "Jumlah" and row_data[3].lower() != "jumlah":
                cell.font = bold_font
                
            if col_idx in [1, 2, 3, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 50
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 16
    
    for r_idx in range(3, ws.max_row + 1):
        cell_d = ws.cell(row=r_idx, column=4)
        if cell_d.value and "jumlah" in str(cell_d.value).lower():
            cell_d.font = bold_font
            cell_d.alignment = Alignment(horizontal="right", vertical="center")
            ws.cell(row=r_idx, column=8).font = bold_font
            
    wb.save(output)
    return output.getvalue()


# ==========================================
# 4. KONTEN UTAMA (TABS)
# ==========================================
st.markdown('<p class="main-title">Aplikasi Konverter BOSP 📊</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Rapikan dan Luruskan Kertas Kerja RKAS yang berantakan dalam 1x klik.</p>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📄 MODE PDF (Standar)", "📊 MODE CSV / EXCEL KOTOR"])

# ------------------------------------------
# TAB 1: MODE PDF LAMA (TIDAK ADA KODE YANG DIRUSAK/DIUBAH LOGIKANYA)
# ------------------------------------------
with tab1:
    st.write("**Gunakan mode ini untuk mengunggah file Kertas Kerja berformat PDF.**")
    uploaded_pdf = st.file_uploader("📂 Upload File PDF Kertas Kerja", type="pdf", key="pdf_uploader")
    
    if uploaded_pdf is not None:
        with st.spinner("⏳ Membedah dokumen dan mengaktifkan Smart Alignment... Mohon tunggu..."):
            try:
                raw_rows = []
                with pdfplumber.open(uploaded_pdf) as pdf:
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            raw_rows.extend(table)
                
                if raw_rows:
                    cleaned_data = []
                    garbage_keywords = [
                        "rincian kertas kerja", "tahun anggaran", "npsn", "nama sekolah", 
                        "alamat", "kabupaten", "provinsi", "bulan", "sumber dana", 
                        "penerimaan", "total penerimaan", "belanja", "kode rekening",
                        "rincian perhitungan", "tarif harga", "no. urut"
                    ]
                    
                    for row in raw_rows:
                        if not row: continue
                        cells = [str(x).strip() for x in row if x is not None and str(x).strip() != ""]
                        if not cells: continue
                        
                        row_str_lower = " ".join(cells).lower()
                        if any(kw in row_str_lower for kw in garbage_keywords): continue
                        
                        no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah = "", "", "", "", "", "", "", ""
                        
                        if "jumlah" in row_str_lower and len(cells) <= 3:
                            uraian = "Jumlah"
                            for c in reversed(cells):
                                if sum(char.isdigit() for char in c) > 3:
                                    jumlah = c
                                    break
                            cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
                            continue
                        
                        if re.match(r'^\d+\.$', cells[0]): no_urut = cells.pop(0)
                        if cells and re.match(r'^5\.\d', cells[0]): kode_rek = cells.pop(0)
                        
                        prog_parts = []
                        while cells and re.match(r'^(?:\d{2}\.\s*)+$', cells[0]):
                            prog_parts.append(cells.pop(0))
                        if prog_parts: kode_prog = " ".join(prog_parts)
                        
                        if cells and re.match(r'^((?:\d{2}\.\s*)+)\s*(.+)$', cells[0]):
                            m = re.match(r'^((?:\d{2}\.\s*)+)\s*(.+)$', cells[0])
                            kode_prog += (" " if kode_prog else "") + m.group(1).strip()
                            if m.group(2).strip(): cells[0] = m.group(2).strip()
                            else: cells.pop(0)
                        
                        if not cells: continue
                            
                        if kode_rek:
                            if len(cells) >= 5:
                                uraian = " ".join(cells[:-4])
                                volume = cells[-4]
                                satuan = cells[-3]
                                tarif = cells[-2]
                                jumlah = cells[-1]
                            elif len(cells) == 4:
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
                            if len(cells) >= 2:
                                uraian = " ".join(cells[:-1])
                                jumlah = cells[-1]
                            elif len(cells) == 1:
                                uraian = cells[0]
                        
                        if any([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah]):
                            cleaned_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
                    
                    if cleaned_data:
                        st.success(f"✅ Berhasil mengekstrak {len(cleaned_data)} baris data!")
                        st.markdown("### 👀 Pratinjau Data (Review)")
                        df_preview = pd.DataFrame(cleaned_data, columns=["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"])
                        st.dataframe(df_preview, use_container_width=True, height=350)
                        
                        excel_data = create_styled_excel(cleaned_data)
                        
                        st.divider()
                        st.markdown("<h3 style='text-align: center; color: #1E88E5;'>🎉 Dokumen Anda Telah Siap!</h3>", unsafe_allow_html=True)
                        st.write("")
                        col1, col2, col3 = st.columns([1, 2, 1])
                        with col2:
                            st.download_button("📥 DOWNLOAD EXCEL FINAL SEKARANG", data=excel_data, file_name="RKAS_BOSP_PDF_Rapi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
                    else: st.warning("Data tabel kosong.")
                else: st.warning("Gagal membaca isi tabel pada PDF.")
            except Exception as e: st.error(f"Error pada sistem: {e}")

# ------------------------------------------
# TAB 2: MODE CSV / EXCEL KOTOR (BARU)
# ------------------------------------------
with tab2:
    st.write("**Gunakan mode ini jika Anda sudah punya file CSV / Excel yang berantakan (seperti hasil convert aplikasi lain).**")
    uploaded_csv = st.file_uploader("📂 Upload File CSV atau Excel yang Kotor", type=["csv", "xlsx", "xls"], key="csv_uploader")
    
    if uploaded_csv is not None:
        with st.spinner("⏳ Menata ulang kolom CSV/Excel secara pintar..."):
            try:
                # Membaca File Apapun yang masuk
                if uploaded_csv.name.endswith('.csv'):
                    df_raw = pd.read_csv(uploaded_csv, header=None)
                else:
                    df_raw = pd.read_excel(uploaded_csv, header=None)
                
                # Konversi semua ke string agar tidak error
                df_raw = df_raw.fillna("").astype(str)
                
                cleaned_csv_data = []
                garbage_keywords = [
                    "rincian kertas kerja", "tahun anggaran", "npsn", "nama sekolah", 
                    "alamat", "kabupaten", "provinsi", "bulan", "sumber dana", 
                    "penerimaan", "total penerimaan", "belanja", "kode rekening",
                    "rincian perhitungan", "tarif harga", "no. urut", "kertas kerja perbulan", "kepala sekolah", "kec. besuk"
                ]
                
                # Menggunakan logika Anti-Bergeser yang sama seperti PDF, tapi untuk baris CSV
                for index, row in df_raw.iterrows():
                    cells = [x.strip() for x in row.tolist() if x.strip() != ""]
                    if not cells: continue
                    
                    row_str_lower = " ".join(cells).lower()
                    
                    if any(kw in row_str_lower for kw in garbage_keywords): continue
                    
                    no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah = "", "", "", "", "", "", "", ""
                    
                    if "jumlah" in row_str_lower and len(cells) <= 3:
                        uraian = "Jumlah"
                        for c in reversed(cells):
                            if sum(char.isdigit() for char in c) > 3:
                                jumlah = c
                                break
                        cleaned_csv_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])
                        continue
                        
                    if re.match(r'^\d+\.$', cells[0]): no_urut = cells.pop(0)
                    if cells and re.match(r'^5\.\d', cells[0]): kode_rek = cells.pop(0)
                    
                    # Cek Kode Program pada CSV yang kadang format tanggal (2026-03-03) atau angka (03.)
                    prog_parts = []
                    while cells and (re.match(r'^(?:\d{2}\.\s*)+$', cells[0]) or re.match(r'^\d{4}-\d{2}-\d{2}$', cells[0])):
                        prog_parts.append(cells.pop(0))
                    if prog_parts: kode_prog = " ".join(prog_parts)
                    
                    if not cells: continue
                    
                    if kode_rek:
                        if len(cells) >= 5:
                            uraian = " ".join(cells[:-4])
                            volume = cells[-4]
                            satuan = cells[-3]
                            tarif = cells[-2]
                            jumlah = cells[-1]
                        elif len(cells) == 4:
                            # Jika Vol + Satuan menempel seperti "10 pack"
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
                        if len(cells) >= 2:
                            uraian = " ".join(cells[:-1])
                            jumlah = cells[-1]
                        elif len(cells) == 1:
                            uraian = cells[0]
                    
                    if any([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah]):
                        cleaned_csv_data.append([no_urut, kode_rek, kode_prog, uraian, volume, satuan, tarif, jumlah])

                if cleaned_csv_data:
                    st.success(f"✅ Data CSV/Excel berantakan berhasil diluruskan! ({len(cleaned_csv_data)} baris)")
                    st.markdown("### 👀 Pratinjau Data CSV/Excel")
                    df_preview_csv = pd.DataFrame(cleaned_csv_data, columns=["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"])
                    st.dataframe(df_preview_csv, use_container_width=True, height=350)
                    
                    excel_data_csv = create_styled_excel(cleaned_csv_data)
                    
                    st.divider()
                    st.markdown("<h3 style='text-align: center; color: #1E88E5;'>🎉 Dokumen Anda Telah Siap!</h3>", unsafe_allow_html=True)
                    st.write("")
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        st.download_button("📥 DOWNLOAD EXCEL FINAL SEKARANG", data=excel_data_csv, file_name="RKAS_BOSP_Dari_CSV_Rapi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
                else: st.warning("Gagal menyaring data. Pastikan format dokumennya mirip Kertas Kerja.")
            except Exception as e: st.error(f"Terjadi kesalahan saat memproses CSV/Excel: {e}")
