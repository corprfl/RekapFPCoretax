import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

st.title("Rekap Faktur Pajak ke Excel (Multi File) - Versi 3")

st.markdown("""
### 📘 Deskripsi
Aplikasi ini membaca **Faktur Pajak dengan dan tanpa kode barang** secara otomatis.
- Faktur dengan kode barang → tetap rapi per item.  
- Faktur tanpa kode barang → deskripsi diambil **utuh** (termasuk baris “x 1,00 Bulan”, “Potongan Harga”, “PPnBM”, dsb).  

---
**By: Reza Fahlevi Lubis BKP @zavibis**
""")

# =====================================================
# KONFIGURASI DASAR
# =====================================================
bulan_map = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

def extract(pat, txt, flags=re.DOTALL, default="-"):
    m = re.search(pat, txt, flags)
    return m.group(1).strip() if m else default

def extract_tanggal(txt):
    m = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", txt)
    if m:
        d = m.group(2).zfill(2)
        b = bulan_map.get(m.group(3), "-")
        y = m.group(4)
        return f"{d}/{b}/{y}"
    return "-"

def extract_nitku(txt):
    for i, l in enumerate(txt.splitlines()):
        if "NPWP" in l and i > 0:
            prev = txt.splitlines()[i-1]
            m = re.search(r"#(\d{22})", prev)
            if m: return m.group(1)
    return "-"

# =====================================================
# 1️⃣ Faktur dengan kode barang
# =====================================================
def extract_tabel_dengan_kode(txt):
    result = []
    pat = re.compile(
        r"(\d+)\s+(\d{6})\s+([\s\S]*?)\n\s*([\d.,]+)\s*(?=\n\d+\s+\d{6}|\nHarga Jual|$)",
        re.MULTILINE
    )
    for m in pat.finditer(txt):
        no, kode = m.group(1), m.group(2)
        deskripsi = " ".join(m.group(3).split())
        try:
            harga = float(m.group(4).replace(".", "").replace(",", "."))
        except:
            harga = 0.0
        result.append({
            "No": no,
            "Kode Barang/Jasa": kode,
            "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
        })
    return result

# =====================================================
# 2️⃣ Faktur tanpa kode barang (ambil semua deskripsi)
# =====================================================
def extract_tabel_tanpa_kode(txt):
    result = []
    # Pisahkan tiap item berdasarkan nomor urut (1, 2, 3...)
    blocks = re.split(r'\n(?=\d+\s*\n)', txt)
    for blk in blocks:
        blk = blk.strip()
        m = re.match(r"(\d+)\s+(.*)", blk, re.DOTALL)
        if not m:
            continue
        no = m.group(1)
        content = m.group(2)

        # Ambil angka terakhir (biasanya harga)
        harga_match = re.findall(r'\b([\d.,]+)\b\s*$', content)
        harga = 0.0
        if harga_match:
            try:
                harga = float(harga_match[-1].replace('.', '').replace(',', '.'))
            except:
                harga = 0.0

        # Ambil seluruh deskripsi sebelum angka terakhir
        deskripsi = re.sub(r'\b[\d.,]+\b\s*$', '', content, flags=re.DOTALL)
        deskripsi = re.sub(r'\s+', ' ', deskripsi).strip()

        if len(deskripsi) > 5 and harga > 0:
            result.append({
                "No": no,
                "Kode Barang/Jasa": "-",
                "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
            })

    # fallback khusus MANDE
    if not result:
        specific = []
        for pattern, no in [
            (r"Rental\s+Heavy\s+Duty\s+Equipment[\s\S]+?PPnBM.*?=\s*Rp\s*0,00", "1"),
            (r"Double\s+Cabin[\s\S]+?PPnBM.*?=\s*Rp\s*0,00", "2"),
        ]:
            m = re.search(pattern, txt, re.IGNORECASE)
            if m:
                desc = " ".join(m.group(0).split())
                harga_match = re.search(r'Rp\s*([\d.,]+)', desc)
                if harga_match:
                    try:
                        harga = float(harga_match.group(1).replace('.', '').replace(',', '.'))
                    except:
                        harga = 0.0
                    specific.append({
                        "No": no,
                        "Kode Barang/Jasa": "-",
                        "Nama Barang Kena Pajak / Jasa Kena Pajak": desc,
                        "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
                    })
        result = specific
    return result

# =====================================================
# 3️⃣ Auto deteksi
# =====================================================
def extract_tabel_auto(txt):
    if re.search(r"\n\s*\d+\s+\d{6}\s+", txt):
        return extract_tabel_dengan_kode(txt)
    else:
        return extract_tabel_tanpa_kode(txt)

# =====================================================
# Bagian total & metadata
# =====================================================
def extract_total(txt):
    def val(p):
        m = re.search(p, txt, re.DOTALL)
        if not m: return 0.0
        try: return float(m.group(1).replace('.', '').replace(',', '.'))
        except: return 0.0
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

def extract_meta(txt):
    return {
        "Kode dan Nomor Seri Faktur Pajak": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", txt),
        "Nama PKP": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", txt),
        "NPWP PKP": extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9.]+)", txt),
        "Nama Pembeli": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", txt),
        "NPWP Pembeli": extract(r"NPWP\s*:\s*([0-9.]+)\s*NIK", txt),
        "NITKU Pembeli": extract_nitku(txt),
        "Kota": extract(r"\n([A-Z .,]+),\s*\d{1,2}\s+\w+\s+\d{4}", txt),
        "Tanggal Faktur Pajak": extract_tanggal(txt),
        "Penandatangan": extract(r"Ditandatangani secara elektronik\n(.*?)\n", txt)
    }

def kode_status(k):
    if not k or len(k) < 3: return "-", "-"
    return k[:2], ("Normal" if k[2] == "0" else "Pengganti")

# =====================================================
# EKSEKUSI
# =====================================================
upl = st.file_uploader("Upload Faktur Pajak (PDF)", type=["pdf"], accept_multiple_files=True)
if upl:
    if st.button("Eksekusi Convert"):
        rows = []
        for f in upl:
            txt = "".join([p.get_text() for p in fitz.open(stream=f.read(), filetype="pdf")])
            meta = extract_meta(txt)
            kode = meta["Kode dan Nomor Seri Faktur Pajak"]
            kf, stt = kode_status(kode)
            meta.update({"Kode Faktur": kf, "Status Faktur": stt, "Nama Asli File": f.name})
            meta.update(extract_total(txt))
            tgl = meta["Tanggal Faktur Pajak"].split("/")
            meta["Masa"] = tgl[1] if len(tgl) > 1 else "-"
            meta["Tahun"] = tgl[2] if len(tgl) > 2 else "-"

            items = extract_tabel_auto(txt)
            if not items:
                items = [{"No": "-", "Kode Barang/Jasa": "-", "Nama Barang Kena Pajak / Jasa Kena Pajak": "Tidak terbaca", "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0.0}]
            for it in items:
                rows.append({**it, **meta})

        df = pd.DataFrame(rows)
        for c in df.columns:
            if "Rp" in c or "Total" in c:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        st.success(f"✅ Berhasil! Total {len(df)} baris diekstrak.")
        st.dataframe(df)

        buf = BytesIO()
        df.to_excel(buf, index=False, engine="openpyxl", float_format="%.0f"); buf.seek(0)
        st.download_button("📥 Download Excel", buf, "rekap_faktur_v3.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
