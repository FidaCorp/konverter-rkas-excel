import streamlit as st
import pdfplumber
import pandas as pd
import io

# Judul Aplikasi
st.set_page_config(page_title="Konverter Kertas Kerja BOSP ke Excel")
st.title("Konverter PDF Kertas Kerja ke Excel 📊")
st.write("Upload file PDF RAPBS atau Rincian Kertas Kerja Anda, dan unduh hasilnya dalam format Excel yang bisa diedit.")

# Fitur Upload File
uploaded_file = st.file_uploader("Pilih file PDF", type="pdf")

if uploaded_file is not None:
    st.info("Sedang memproses dokumen...")
    
    try:
        all_data = []
        # Membaca PDF dan mengekstrak tabel
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)
        
        if all_data:
            # Mengubah data menjadi format tabel (DataFrame)
            # Asumsi baris pertama adalah header (No, Kode Rekening, Uraian, dll)
            df = pd.DataFrame(all_data[1:], columns=all_data[0])
            
            # Membuat file Excel di memori
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Rincian Kertas Kerja')
            processed_data = output.getvalue()
            
            st.success("Berhasil! File Excel siap diunduh.")
            
            # Tombol Download
            st.download_button(
                label="📥 Unduh File Excel",
                data=processed_data,
                file_name="Output_Kertas_Kerja_BOSP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Tidak ada tabel yang ditemukan dalam PDF ini.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
