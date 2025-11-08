import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

# ===============================================================
# 🧾 APLIKASI REKAP FAKTUR PAJAK KE EXCEL (MULTI FILE) - COMBINED
# ===============================================================

st.title("Rekap Faktur Pajak ke Excel (Multi File) - Versi Gabungan")

st.markdown("""
### 📘 Deskripsi Aplikasi
Aplikasi ini dapat membaca **dua jenis Faktur Pajak**:
- Faktur dengan **kode barang (6 digit)**  
- Faktur tanpa kode barang (deskripsi langsung)  

Hasil ekstraksi berupa Excel dengan kolom lengkap termasuk total DPP, PPN, PPnBM.

---
**By: Reza Fahlevi Lubis BKP @zavibis**
""")

# ===============================================================
# KONFIGURASI & FUNGSI UMUM
# ===============================================================

bulan_map = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

def extract(pattern, text, flags=re.DOTALL, default="-", postproc=lambda x: x.strip()):
    match = re.search(pattern, text, flags)
    return postproc(match.group(1)) if match else default

def extract_tanggal(text):
    match = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if match:
        hari = match.group(2).zfill(2)
        bulan = bulan_map.get(match.group(3), "-")
        tahun = match.group(4)
        return f"{hari}/{bulan}/{tahun}"
    return "-"

def extract_nitku_pembeli(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "NPWP" in line and i > 0:
            prev_line = lines[i-1]
            match = re.search(r"#(\d{22})", prev_line)
            if match:
                return match.group(1)
    return "-"

# ===============================================================
# TABEL RINCI UNTUK FAKTUR DENGAN KODE BARANG
# ===============================================================
def extract_tabel_rinci(text):
    result = []
    pattern = re.compile(
        r"(\d+)\s+(\d{6})\s+([\s\S]*?)\n\s*([\d.,]+)\s*(?=\n\d+\s+\d{6}|\nHarga Jual|$)",
        re.MULTILINE
    )
    for m in pattern.finditer(text):
        no = m.group(1)
        kode = m.group(2)
        deskripsi = " ".join(m.group(3).split())

        raw = m.group(4).replace(".", "").replace(",", ".")
        try:
            harga = float(raw)
        except:
            harga = 0.0

        result.append({
            "No": no,
            "Kode Barang/Jasa": kode,
            "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
        })
    return result

# ===============================================================
# TABEL RINCI UNTUK FAKTUR TANPA KODE BARANG
# ===============================================================
def extract_tabel_rinci_final(txt):
    result = []
    table_pattern = r'(\d+)\s+(.*?)(?=\d+\s+\w+|Harga\s+Jual.*?Penggantian|Dikurangi\s+Potongan|$)'
    matches = re.findall(table_pattern, txt, re.DOTALL)
    for no, content in matches:
        try:
            if int(no) > 10:
                continue
        except:
            continue
        content = content.strip()
        harga_matches = re.findall(r'Rp\s*([\d.,]+)', content)
        harga = 0.0
        if harga_matches:
            try:
                harga = float(harga_matches[0].replace('.', '').replace(',', '.'))
            except:
                harga = 0.0
        deskripsi = re.sub(r'Rp\s*[\d.,]+.*?(?=\n|$)', '', content, flags=re.MULTILINE)
        deskripsi = re.sub(r'Potongan\s+Harga.*?(?=\n|$)', '', deskripsi, flags=re.MULTILINE | re.IGNORECASE)
        deskripsi = re.sub(r'PPnBM.*?(?=\n|$)', '', deskripsi, flags=re.MULTILINE | re.IGNORECASE)
        deskripsi = re.sub(r'x\s*[\d.,]+\s*Bulan.*?(?=\n|$)', '', deskripsi, flags=re.MULTILINE | re.IGNORECASE)
        deskripsi = re.sub(r'\s+', ' ', deskripsi).strip()
        if len(deskripsi) > 5 and harga > 0:
            result.append({
                "No": no,
                "Kode Barang/Jasa": "-",
                "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
            })
    return result

# ===============================================================
# PILIH OTOMATIS SESUAI STRUKTUR PDF
# ===============================================================
def extract_tabel_rinci_combined(text):
    if re.search(r"\n\s*\d+\s+\d{6}\s+", text):
        return extract_tabel_rinci(text)
    else:
        return extract_tabel_rinci_final(text)

# ===============================================================
# TOTAL BAGIAN BAWAH
# ===============================================================
def extract_total_section(text):
    def get_val(pat):
        m = re.search(pat, text, re.DOTALL)
        if not m:
            return 0.0
        raw = m.group(1).replace(".", "").replace(",", ".")
        try:
            return float(raw)
        except:
            return 0.0
    return {
        "Total Harga Jual / Penggantian / Uang Muka / Termin": get_val(r"Harga\s+Jual\s*/\s*Penggantian\s*/\s*Uang\s*Muka\s*/\s*Termin\s*([\d.,]+)"),
        "Dikurangi Potongan Harga (Total)": get_val(r"Dikurangi\s+Potongan\s+Harga\s*([\d.,]+)"),
        "Dikurangi Uang Muka yang telah diterima (Total)": get_val(r"Dikurangi\s+Uang\s+Muka\s+yang\s+telah\s+diterima\s*([\d.,]+)"),
        "Dasar Pengenaan Pajak (Total)": get_val(r"Dasar\s+Pengenaan\s+Pajak\s*([\d.,]+)"),
        "PPN (Total)": get_val(r"Jumlah\s*PPN.*?([\d.,]+)"),
        "Jumlah PPnBM (Total)": get_val(r"Jumlah\s*PPnBM.*?([\d.,]+)")
    }

# ===============================================================
# METADATA
# ===============================================================
def extract_data_from_text(text):
    return {
        "Kode dan Nomor Seri Faktur Pajak": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", text),
        "Nama Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", text),
        "Alamat Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:.*?Alamat\s*:\s*(.*?)\s*NPWP", text),
        "NPWP Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9.]+)", text),
        "Nama Pembeli Barang/Jasa": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", text),
        "Alamat Pembeli Barang/Jasa": extract(r"Pembeli Barang Kena Pajak.*?Alamat\s*:\s*(.*?)\s*#", text),
        "NPWP Pembeli Barang/Jasa": extract(r"NPWP\s*:\s*([0-9.]+)\s*NIK", text),
        "NITKU Pembeli": extract_nitku_pembeli(text),
        "Kota": extract(r"\n([A-Z .,]+),\s*\d{1,2}\s+\w+\s+\d{4}", text),
        "Tanggal Faktur Pajak": extract_tanggal(text),
        "Referensi": extract(r"Referensi:\s*(.*?)\n", text),
        "Penandatangan": extract(r"Ditandatangani secara elektronik\n(.*?)\n", text),
    }

def extract_kode_status(kode_seri):
    if not kode_seri or len(kode_seri) < 3:
        return "-", "-"
    kode_faktur = kode_seri[:2]
    status = "Normal" if kode_seri[2] == "0" else "Pengganti"
    return kode_faktur, status

# ===============================================================
# PROSES UTAMA
# ===============================================================
uploaded_files = st.file_uploader("Upload PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Eksekusi Convert"):
        final_rows = []

        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            pdf_bytes = uploaded_file.read()
            with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
                full_text = "".join([page.get_text() for page in doc])

            data = extract_data_from_text(full_text)
            kode_seri = data.get("Kode dan Nomor Seri Faktur Pajak", "")
            kode_faktur, status = extract_kode_status(kode_seri)
            data["Kode Faktur"] = kode_faktur
            data["Status Faktur"] = status
            data["Nama Asli File"] = filename
            total_section = extract_total_section(full_text)
            data.update(total_section)

            try:
                tgl_parts = data["Tanggal Faktur Pajak"].split("/")
                data["Masa"] = tgl_parts[1]
                data["Tahun"] = tgl_parts[2]
            except:
                data["Masa"] = "-"
                data["Tahun"] = "-"

            rinci = extract_tabel_rinci_combined(full_text)
            if not rinci:
                rinci = [{
                    "No": "-",
                    "Kode Barang/Jasa": "-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak": "Tidak terbaca",
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0.0
                }]

            for row in rinci:
                merged = {**row, **data}
                final_rows.append(merged)

        df = pd.DataFrame(final_rows)
        for col in df.columns:
            if "Rp" in col or "Total" in col:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        st.success("✅ Semua file berhasil diekstrak dan dikonversi ke Excel!")
        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl", float_format="%.0f")
        buffer.seek(0)
        st.download_button(
            "📥 Download Rekap Excel",
            buffer,
            file_name="rekap_faktur_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
