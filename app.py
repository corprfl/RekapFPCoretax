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
Aplikasi ini digunakan untuk **mengekstrak data dari Faktur Pajak PDF** secara otomatis 
menjadi **file Excel** yang berisi rincian lengkap seperti:
Kode dan Nomor Seri Faktur Pajak, Nama PKP, Pembeli, NITKU, Harga Jual, DPP, PPN, 
dan **Tanggal Faktur Pajak**.

### ⚙️ Cara Penggunaan
1. Upload satu atau beberapa file PDF Faktur Pajak menggunakan tombol di bawah.  
2. Klik tombol **"Eksekusi Convert"** untuk memproses.  
3. Setelah proses selesai, hasil akan tampil di layar dan bisa diunduh ke Excel.

### ⚠️ Disclaimer
- Aplikasi ini **tidak menyimpan file atau data pribadi Anda**.  
- Semua proses dilakukan **langsung di perangkat lokal Anda (client-side)**.  
- Hasil file hanya digunakan untuk kebutuhan rekap internal.

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
    """Ambil tanggal faktur pajak format dd/mm/yyyy dari teks lokasi + tanggal"""
    match = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if match:
        hari = match.group(2).zfill(2)
        bulan_huruf = match.group(3)
        bulan = bulan_map.get(bulan_huruf, "-")
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


def extract_tabel_rinci(text):
    """Ambil seluruh blok item sampai PPnBM = Rp 0,00"""
    result = []
    pattern = re.compile(
        r"(\d+)\s+(\d{6})\s+([\s\S]*?PPnBM\s*\(.*?\)\s*=\s*Rp\s*0,00)\s*([\d.]+,[\d]{2})",
        re.MULTILINE
    )
    for m in pattern.finditer(text):
        nama_brg = " ".join(m.group(3).split())
        harga_str = m.group(4).replace(".", "").replace(",", ".")
        try:
            harga = float(harga_str)
        except:
            harga = 0
        result.append({
            "No": m.group(1),
            "Kode Barang/Jasa": m.group(2),
            "Nama Barang Kena Pajak / Jasa Kena Pajak": nama_brg,
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
        })
    return result


def extract_data_from_text(text):
    """Ambil seluruh metadata faktur"""
    return {
        "Kode dan Nomor Seri Faktur Pajak": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", text),
        "Nama Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", text),
        "Alamat Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:.*?Alamat\s*:\s*(.*?)\s*NPWP", text),
        "NPWP Pengusaha Kena Pajak": extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9.]+)", text),
        "Nama Pembeli Barang/Jasa": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", text),
        "Alamat Pembeli Barang/Jasa": extract(r"Pembeli Barang Kena Pajak.*?Alamat\s*:\s*(.*?)\s*#", text),
        "NPWP Pembeli Barang/Jasa": extract(r"NPWP\s*:\s*([0-9.]+)\s*NIK", text),
        "NITKU Pembeli": extract_nitku_pembeli(text),
        "Jumlah PPnBM": extract(r"Jumlah PPnBM.*?([0-9.]+,[0-9]+)", text),
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

            # Ambil bulan & tahun dari tanggal faktur
            try:
                tgl_parts = data["Tanggal Faktur Pajak"].split("/")
                data["Masa"] = tgl_parts[1]
                data["Tahun"] = tgl_parts[2]
            except:
                data["Masa"] = "-"
                data["Tahun"] = "-"

            # Ekstrak tabel rinci
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
                try:
                    harga = float(row["Harga Jual / Penggantian / Uang Muka / Termin (Rp)"])
                    kode_faktur = merged.get("Kode dan Nomor Seri Faktur Pajak", "")
                    if kode_faktur.startswith("01"):
                        dpp = harga
                        ppn = round(dpp * 0.12)
                    elif kode_faktur.startswith("05"):
                        dpp = harga
                        ppn = round(dpp * 11 / 12 * 0.12)
                    else:
                        dpp = round(harga * 11 / 12)
                        ppn = round(dpp * 0.12)
                    merged["DPP"] = f"{dpp:,.0f}".replace(",", ".")
                    merged["PPN"] = f"{ppn:,.0f}".replace(",", ".")
                except:
                    merged["DPP"] = "-"
                    merged["PPN"] = "-"
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
