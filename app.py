
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO

st.markdown("**By : Reza Fahlevi Lubis BKP @zavibis**")
st.title("Rekap Faktur Pajak ke Excel (Multi File)")

bulan_map = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

def extract(pattern, text, flags=re.DOTALL, default="-", postproc=lambda x: x.strip()):
    match = re.search(pattern, text, flags)
    return postproc(match.group(1)) if match else default

def extract_tanggal(text):
    match = re.search(r",\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    return f"{match.group(1).zfill(2)}/{match.group(2)}/{match.group(3)}" if match else "-"

def extract_nitku_pembeli(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "NPWP" in line and i > 0:
            prev_line = lines[i-1]
            match = re.search(r"#(\d{22})", prev_line)
            if match:
                return match.group(1)
    return "-"

def extract_tabel_rinci(text):
    blocks = re.findall(r"(\d+)\s+(\d{6})\s+(.*?)\s+Rp\s*([0-9.,]+)", text, re.DOTALL)
    data = []
    for no, kode, desc, harga in blocks:
        desc_block = re.search(rf"{kode}\s+(.*?PPnBM.*?=\s*Rp\s*0,00)", text, re.DOTALL)
        full_desc = desc_block.group(1).replace("\n", " ") if desc_block else desc
        harga_fix = re.sub(r"[^0-9,]", "", harga)
        data.append({
            "No": no,
            "Kode Barang/Jasa": kode,
            "Nama Barang Kena Pajak / Jasa Kena Pajak": full_desc.strip(),
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga_fix
        })
    return data

def extract_data_from_text(text):
    return {
        "Kode dan Nomor Seri Faktur Pajak": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", text),
        "Nama Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", text),
        "alamat Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:.*?Alamat\s*:\s*(.*?)\s*NPWP", text),
        "npwp Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9.]+)", text),
        "Nama Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", text),
        "Alamat Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:": extract(r"Pembeli Barang Kena Pajak.*?Alamat\s*:\s*(.*?)\s*#", text),
        "NPWP Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:": extract(r"NPWP\s*:\s*([0-9.]+)\s*NIK", text),
        "NITKU Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:": extract_nitku_pembeli(text),
        "Dasar Pengenaan Pajak": extract(r"Dasar Pengenaan Pajak\s*([0-9.]+,[0-9]+)", text),
        "Jumlah PPN": extract(r"Jumlah PPN.*?([0-9.]+,[0-9]+)", text),
        "Jumlah PPnBM": extract(r"Jumlah PPnBM.*?([0-9.]+,[0-9]+)", text),
        "Kota": extract(r"\n([A-Z .,]+),\s*\d{1,2}\s+\w+\s+\d{4}", text),
        "Tanggal faktur pajak": extract_tanggal(text),
        "referensi": extract(r"Referensi:\s*(.*?)\n", text),
        "Penandatangan": extract(r"Ditandatangani secara elektronik\n(.*?)\n", text),
    }

uploaded_files = st.file_uploader("Upload satu atau beberapa PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Eksekusi Convert"):
        all_data = []
        tabel_rinci = []

        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            with fitz.open(stream=uploaded_file.read(), filetype="pdf") as doc:
                full_text = ""
                for page in doc:
                    full_text += page.get_text()

            data = extract_data_from_text(full_text)
            data["Nama asli file"] = filename
            data["Kode Faktur"] = data["Kode dan Nomor Seri Faktur Pajak"][:2]

            try:
                tgl_parts = data["Tanggal faktur pajak"].split("/")
                data["Masa"] = bulan_map.get(tgl_parts[1], "-")
                data["Tahun"] = tgl_parts[2]
            except:
                data["Masa"] = "-"
                data["Tahun"] = "-"

            all_data.append(data)
            tabel_rinci.extend(extract_tabel_rinci(full_text))

        df = pd.DataFrame(all_data)
        df_rinci = pd.DataFrame(tabel_rinci)

        st.success("Semua file berhasil diekstrak!")
        st.dataframe(df)
        st.dataframe(df_rinci)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Rekap Header")
            df_rinci.to_excel(writer, index=False, sheet_name="Detil Barang")
        buffer.seek(0)
        st.download_button("Download Rekap Excel", buffer, file_name="rekap_faktur_multi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
