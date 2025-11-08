import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

st.title("Rekap Faktur Pajak ke Excel (Multi File)")

st.markdown("""
### 📘 Deskripsi Aplikasi
Konversi **Faktur Pajak PDF** menjadi **Excel** lengkap dengan rincian barang, DPP, PPN, PPnBM, dan identitas PKP.

### ⚙️ Cara Pakai
1. Upload satu atau beberapa PDF Faktur Pajak.  
2. Klik **Eksekusi Convert**.  
3. Lihat hasilnya, lalu unduh Excel.

### ⚠️ Disclaimer
Semua proses berjalan di perangkat Anda — **tidak ada data disimpan**.  
---
**By Reza Fahlevi Lubis BKP @zavibis**
""")

# ---------------------------------------------------------------
bulan_map = {
    "Januari":"01","Februari":"02","Maret":"03","April":"04",
    "Mei":"05","Juni":"06","Juli":"07","Agustus":"08",
    "September":"09","Oktober":"10","November":"11","Desember":"12"
}

def extract(pattern, text, flags=re.DOTALL, default="-", post=lambda x:x.strip()):
    m = re.search(pattern, text, flags)
    return post(m.group(1)) if m else default

def extract_tanggal(text):
    m = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if m:
        d = m.group(2).zfill(2)
        b = bulan_map.get(m.group(3), "-")
        y = m.group(4)
        return f"{d}/{b}/{y}"
    return "-"

def extract_nitku_pembeli(text):
    lines = text.splitlines()
    for i, l in enumerate(lines):
        if "NPWP" in l and i > 0:
            prev = lines[i-1]
            m = re.search(r"#(\d{22})", prev)
            if m:
                return m.group(1)
    return "-"

# ---------------------------------------------------------------
# PARSER UTAMA — gabung deskripsi penuh (v11.7)
def extract_tabel_rinci(text):
    result = []
    body = re.split(r"\nHarga\s+Jual\s*/\s*Penggantian", text, 1)[0]
    lines = [l.strip() for l in body.splitlines() if l.strip()]
    i = 0

    while i < len(lines):
        line = lines[i]
        if re.match(r"^No\s*\d+\.?", line) or re.match(r"^\d+\.?$", line):
            no = re.sub(r"\D", "", line)
            desc_parts = [line]
            i += 1
            # ambil baris sampai No berikutnya / total section
            while i < len(lines) and not re.match(r"^(No\s*\d+\.?|\d+\.?|Harga\s+Jual)", lines[i]):
                desc_parts.append(lines[i])
                i += 1

            blok = " ".join(desc_parts)
            blok = re.sub(r"\s+", " ", blok).strip()

            # ambil angka terakhir dari blok
            harga_match = re.findall(r"([\d.,]+)(?=[^\d.,]*$)", blok)
            harga = 0.0
            if harga_match:
                raw = harga_match[-1].replace(".", "").replace(",", ".")
                try:
                    harga = float(raw)
                except:
                    harga = 0.0

            # bersihkan deskripsi tanpa hilang konteks
            deskripsi = re.sub(r"Rp\s*[\d.,]+\s*x.*", "", blok)
            deskripsi = re.sub(r"Potongan Harga\s*=\s*Rp\s*[\d.,]+", "", deskripsi)
            deskripsi = re.sub(r"PPnBM\s*\(.*?\)\s*=\s*Rp\s*[\d.,]+", "", deskripsi)
            deskripsi = re.sub(r"\s+", " ", deskripsi).strip()

            result.append({
                "No": no,
                "Kode Barang/Jasa": "-",
                "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
            })
        else:
            i += 1

    return result

# ---------------------------------------------------------------
def extract_total_section(text):
    def val(p):
        m = re.search(p, text, re.DOTALL)
        if not m:
            return 0.0
        raw = m.group(1).strip().replace(".", "").replace(",", ".")
        try:
            return float(raw)
        except:
            return 0.0
    return {
        "Total Harga Jual / Penggantian / Uang Muka / Termin":
            val(r"Harga\s*Jual\s*/\s*Penggantian\s*/\s*Uang\s*Muka\s*/\s*Termin\s*([\d.,]+)"),
        "Dikurangi Potongan Harga (Total)":
            val(r"Dikurangi\s+Potongan\s+Harga\s*([\d.,]*)"),
        "Dikurangi Uang Muka yang telah diterima (Total)":
            val(r"Dikurangi\s+Uang\s+Muka\s+yang\s+telah\s+diterima\s*([\d.,]*)"),
        "Dasar Pengenaan Pajak (Total)":
            val(r"Dasar\s+Pengenaan\s+Pajak\s*([\d.,]+)"),
        "PPN (Total)":
            val(r"Jumlah\s*PPN.*?([\d.,]+)"),
        "Jumlah PPnBM (Total)":
            val(r"Jumlah\s*PPnBM.*?([\d.,]+)")
    }

# ---------------------------------------------------------------
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
        "Penandatangan": extract(r"Ditandatangani secara elektronik\n(.*?)\n", text)
    }

def extract_kode_status(kode):
    if not kode or len(kode) < 3:
        return "-", "-"
    return kode[:2], ("Normal" if kode[2] == "0" else "Pengganti")

# ---------------------------------------------------------------
uploaded_files = st.file_uploader("Upload PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    if st.button("Eksekusi Convert"):
        rows = []
        for f in uploaded_files:
            data_bytes = f.read()
            with fitz.open(stream=data_bytes, filetype="pdf") as doc:
                text = "".join(p.get_text() for p in doc)

            meta = extract_data_from_text(text)
            kode = meta.get("Kode dan Nomor Seri Faktur Pajak", "")
            kf, stt = extract_kode_status(kode)
            meta.update({"Kode Faktur": kf, "Status Faktur": stt, "Nama Asli File": f.name})
            meta.update(extract_total_section(text))

            tgl = meta["Tanggal Faktur Pajak"].split("/")
            meta["Masa"] = tgl[1] if len(tgl) > 1 else "-"
            meta["Tahun"] = tgl[2] if len(tgl) > 2 else "-"

            items = extract_tabel_rinci(text)
            if not items:
                items = [{
                    "No": "-",
                    "Kode Barang/Jasa": "-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak": "-",
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0.0
                }]
            for it in items:
                rows.append({**it, **meta})

        df = pd.DataFrame(rows)
        numcols = [
            "Total Harga Jual / Penggantian / Uang Muka / Termin",
            "Dikurangi Potongan Harga (Total)",
            "Dikurangi Uang Muka yang telah diterima (Total)",
            "Dasar Pengenaan Pajak (Total)",
            "PPN (Total)", "Jumlah PPnBM (Total)",
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)"
        ]
        for c in numcols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        order = [
            "Kode Faktur", "Status Faktur", "Kode dan Nomor Seri Faktur Pajak",
            "Nama Pengusaha Kena Pajak", "NPWP Pengusaha Kena Pajak",
            "Nama Pembeli Barang/Jasa", "NPWP Pembeli Barang/Jasa", "NITKU Pembeli",
            "Tanggal Faktur Pajak", "Kota",
            "No", "Kode Barang/Jasa", "Nama Barang Kena Pajak / Jasa Kena Pajak",
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)",
            "Total Harga Jual / Penggantian / Uang Muka / Termin",
            "Dikurangi Potongan Harga (Total)",
            "Dikurangi Uang Muka yang telah diterima (Total)",
            "Dasar Pengenaan Pajak (Total)",
            "PPN (Total)", "Jumlah PPnBM (Total)",
            "Referensi", "Penandatangan", "Nama Asli File", "Masa", "Tahun"
        ]
        df = df[[c for c in order if c in df.columns]]

        st.success("✅ Parsing dan konversi berhasil!")
        st.dataframe(df)

        buf = BytesIO()
        df.to_excel(buf, index=False, engine="openpyxl", float_format="%.0f")
        buf.seek(0)
        st.download_button("📥 Download Rekap Excel", buf, "rekap_faktur.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
