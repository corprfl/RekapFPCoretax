
import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

st.title("Rekap Faktur Pajak ke Excel (Multi File) - Revisi Fix Harga")

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
        r"(\d+)\s+(\d{6})\s+((?:.*?)(?:PPnBM.*?)=\s*Rp\s*0,00.*?)\s+([0-9.]+,[0-9]{2})",
        re.DOTALL
    )
    for m in pattern.finditer(text):
        nama_brg = " ".join(m.group(3).split())
        harga = m.group(4).replace(".", "").replace(",", "")  # raw numeric string
        result.append({
            "No": m.group(1),
            "Kode Barang/Jasa": m.group(2),
            "Nama Barang Kena Pajak / Jasa Kena Pajak": nama_brg,
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
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

uploaded_files = st.file_uploader("Upload PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Eksekusi Convert"):
        final_rows = []
        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            with fitz.open(stream=uploaded_file.read(), filetype="pdf") as doc:
                full_text = "".join([page.get_text() for page in doc])

            data = extract_data_from_text(full_text)
            data["Nama asli file"] = filename
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
                try:
                    harga = int(row["Harga Jual / Penggantian / Uang Muka / Termin (Rp)"])
                    kode_faktur = merged.get("Kode Faktur", merged["Kode dan Nomor Seri Faktur Pajak"][:2])
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

                for kol in ["DPP", "PPN"]:
                    val = merged[kol]
                    if isinstance(val, (int, float)):
                        merged[kol] = f"{val:.2f}".replace(".", ",")

                final_rows.append(merged)

        df = pd.DataFrame(final_rows)
        st.success("Semua file berhasil diekstrak!")
        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        st.download_button("Download Rekap Excel", buffer, file_name="rekap_faktur.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
