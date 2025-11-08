import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

# ===============================================================
# 🧾 APLIKASI REKAP FAKTUR PAJAK KE EXCEL (MULTI FILE)
# ===============================================================

st.title("Rekap Faktur Pajak ke Excel (Multi File)")

st.markdown("""
### 📘 Deskripsi Aplikasi
Aplikasi ini digunakan untuk **mengekstrak data dari Faktur Pajak PDF** menjadi **file Excel**
berisi rincian lengkap seperti Kode Faktur, Nama PKP, Pembeli, NITKU, Harga Jual, PPN, dan
**bagian total Faktur (DPP, Potongan, Uang Muka, PPN, PPnBM)**.

### ⚙️ Cara Penggunaan
1. Upload satu atau beberapa file PDF Faktur Pajak menggunakan tombol di bawah.  
2. Klik **"Eksekusi Convert"** untuk memproses.  
3. Setelah selesai, hasil tampil di layar dan bisa diunduh ke Excel.

### ⚠️ Disclaimer
- Aplikasi **tidak menyimpan file atau data pribadi Anda**.  
- Semua proses dilakukan **langsung di perangkat lokal Anda (client-side)**.  
- Hasil hanya untuk kebutuhan rekap internal.

---
**By: Reza Fahlevi Lubis BKP @zavibis**
""")

# ===============================================================
# KONFIGURASI DASAR
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
    """Ambil tanggal faktur pajak format dd/mm/yyyy"""
    match = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if match:
        hari = match.group(2).zfill(2)
        bulan = bulan_map.get(match.group(3), "-")
        tahun = match.group(4)
        return f"{hari}/{bulan}/{tahun}"
    return "-"


def extract_nitku_pembeli(text):
    """Ambil NITKU dari akhir alamat pembeli (setelah tanda #)"""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "NPWP" in line and i > 0:
            prev_line = lines[i-1]
            match = re.search(r"#(\d{22})", prev_line)
            if match:
                return match.group(1)
    return "-"


def extract_tabel_rinci(text):
    """Ambil setiap blok item Faktur dengan harga di bawahnya"""
    result = []
    pattern = re.compile(
        r"(\d+)\s+(\d{6})\s+([\s\S]*?)\n\s*([\d.]+,[\d]{2})\s*(?=\n\d+\s+\d{6}|\nHarga Jual|$)",
        re.MULTILINE
    )
    for m in pattern.finditer(text):
        no = m.group(1)
        kode = m.group(2)
        deskripsi = " ".join(m.group(3).split())
        harga_str = m.group(4).replace(".", "").replace(",", ".")
        try:
            harga = float(harga_str)
        except:
            harga = 0.0
        result.append({
            "No": no,
            "Kode Barang/Jasa": kode,
            "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
        })
    return result


def extract_total_section(text):
    """Ambil bagian total di bawah tabel"""
    def get_val(pat):
        m = re.search(pat, text)
        if not m:
            return "-"
        return m.group(1).replace(".", "").replace(",", ".")
    return {
        "Total Harga Jual / Penggantian / Uang Muka / Termin": get_val(r"Harga Jual\s*/\s*Penggantian\s*/\s*Uang Muka\s*/\s*Termin\s*([0-9.]+,[0-9]+)"),
        "Dikurangi Potongan Harga (Total)": get_val(r"Dikurangi Potongan Harga\s*([0-9.]+,[0-9]+)"),
        "Dikurangi Uang Muka yang telah diterima (Total)": get_val(r"Dikurangi Uang Muka yang telah diterima\s*([0-9.]+,[0-9]+)"),
        "Dasar Pengenaan Pajak (Total)": get_val(r"Dasar Pengenaan Pajak\s*([0-9.]+,[0-9]+)"),
        "PPN (Total)": get_val(r"Jumlah PPN.*?([0-9.]+,[0-9]+)"),
        "Jumlah PPnBM (Total)": get_val(r"Jumlah PPnBM.*?([0-9.]+,[0-9]+)")
    }


def extract_data_from_text(text):
    """Ambil metadata Faktur"""
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

# ===============================================================
# UPLOAD & PROSES PDF
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
            data["Nama Asli File"] = filename

            # ambil total-section
            total_section = extract_total_section(full_text)
            data.update(total_section)

            # ambil bulan & tahun dari tanggal
            try:
                tgl_parts = data["Tanggal Faktur Pajak"].split("/")
                data["Masa"] = tgl_parts[1]
                data["Tahun"] = tgl_parts[2]
            except:
                data["Masa"] = "-"
                data["Tahun"] = "-"

            rinci = extract_tabel_rinci(full_text)
            if not rinci:
                rinci = [{
                    "No": "-",
                    "Kode Barang/Jasa": "-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak": "-",
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0
                }]

            for row in rinci:
                merged = {**row, **data}
                final_rows.append(merged)

        # ===============================================================
        # HASIL OUTPUT
        # ===============================================================
        df = pd.DataFrame(final_rows)
        st.success("✅ Semua file berhasil diekstrak dan dikonversi ke Excel!")
        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)

        st.download_button(
            "📥 Download Rekap Excel",
            buffer,
            file_name="rekap_faktur.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
