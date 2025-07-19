
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
    result = []
    pattern = re.compile(
        r"(?P<no>\d+)\s+(?P<kode>\d{6})\s+(?P<deskripsi>.+?PPnBM.*?=\s*Rp\s*0,00).*?(?P<harga>[0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})",
        re.DOTALL
    )
    matches = pattern.finditer(text)
    for m in matches:
        result.append({
            "No": m.group("no"),
            "Kode Barang/Jasa": m.group("kode"),
            "Nama Barang Kena Pajak / Jasa Kena Pajak": " ".join(m.group("deskripsi").split()),
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": m.group("harga").replace(".", "").replace(",", ",")
        })
    return result
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
        "Jumlah PPnBM": extract(r"Jumlah PPnBM.*?([0-9.]+,[0-9]+)", text),
        "Kota": extract(r"\n([A-Z .,]+),\s*\d{1,2}\s+\w+\s+\d{4}", text),
        "Tanggal faktur pajak": extract_tanggal(text),
        "referensi": extract(r"Referensi:\s*(.*?)\n", text),
        "Penandatangan": extract(r"Ditandatangani secara elektronik\n(.*?)\n", text),
    }

uploaded_files = st.file_uploader("Upload satu atau beberapa PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Eksekusi Convert"):
        final_rows = []

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

            rinci = extract_tabel_rinci(full_text)
            for row in rinci:
                merged = row | data
                
                
                # Hitung DPP dan PPN berdasarkan kode faktur
                try:
                    harga_str = merged["Harga Jual / Penggantian / Uang Muka / Termin (Rp)"].replace(".", "").replace(",", "")
                    harga = int(harga_str)
                    kode_faktur = merged.get("Kode Faktur", "")
                    if kode_faktur == "01":
                        dpp = harga
                        ppn = round(dpp * 0.12)
                    elif kode_faktur == "05":
                        dpp = harga
                        ppn = round(dpp * 11 / 12 * 0.12)
                    else:
                        dpp = round(harga * 11 / 12)
                        ppn = round(dpp * 0.12)
                    merged["DPP"] = dpp
                    merged["PPN"] = ppn
                except:
                    merged["DPP"] = ""
                    merged["PPN"] = ""
                final_rows.append(merged)

        df = pd.DataFrame(final_rows)
        st.success("Semua file berhasil diekstrak!")
        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        st.download_button("Download Rekap Excel", buffer, file_name="rekap_faktur_multi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
