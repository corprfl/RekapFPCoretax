import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from streamlit_sortables import sort_items  # ✅ untuk drag & drop

# =====================================================
# 🧾 EXTRACTOR ISI FAKTUR PAJAK KE EXCEL
# =====================================================

st.title("Extractor isi Faktur Pajak ke Excel (Versi Pilih Kolom + Drag & Drop)")

st.markdown("""
### 📘 Deskripsi Aplikasi
Aplikasi ini digunakan untuk **mengekstrak isi Faktur Pajak (PDF)** menjadi **file Excel** secara otomatis.  
Data yang diambil mencakup:
- Informasi utama faktur (Nomor, Tanggal, Nama PKP, NPWP, Pembeli, dsb)
- Detail barang/jasa (kode atau deskripsi lengkap)
- Nilai transaksi (DPP, PPN, PPnBM, potongan, dan total lainnya)

### ⚙️ Fungsi Utama
Extractor ini membantu Anda mengonversi **banyak Faktur Pajak PDF** menjadi Excel secara cepat,  
tanpa perlu mengetik manual satu per satu.

---

### 🧩 Panduan Penggunaan

**🗂️ Langkah 1 — Upload File Faktur**
- Klik tombol **Browse files** untuk memilih satu atau beberapa file PDF Faktur Pajak.

**📖 Langkah 2 — Baca File**
- Tekan tombol **📖 Baca File** untuk mengekstrak isi faktur.  
- Hasil pembacaan akan muncul sebagai tabel di layar.

**⚙️ Langkah 3 — Pilih Kolom**
- Gunakan daftar **Pilih Kolom** untuk menentukan kolom mana yang ingin diekspor ke Excel.  
- Klik **✅ Pilih Semua Kolom** untuk memilih semuanya, atau **❌ Hapus Semua Kolom** untuk reset.  
- Anda juga dapat **menyusun ulang urutan kolom** dengan cara **drag & drop** di tampilan kotak biru.

**📊 Langkah 4 — Preview & Download**
- Pratinjau 5 baris pertama dari kolom yang dipilih akan ditampilkan.  
- Tekan tombol **📥 Konversi & Download Excel** untuk mengunduh hasil akhir.

---

### ⚠️ Disclaimer
Semua proses dijalankan **langsung di perangkat Anda (client-side)**.  
Tidak ada file yang diunggah atau disimpan di server.  
Keamanan dan kerahasiaan dokumen Anda sepenuhnya terjamin.

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
# EKSTRAKSI ISI TABEL FAKTUR
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

def extract_tabel_tanpa_kode(txt):
    result = []
    blocks = re.split(r'\n(?=\d+\s*\n)', txt)
    for blk in blocks:
        blk = blk.strip()
        m = re.match(r"(\d+)\s+(.*)", blk, re.DOTALL)
        if not m:
            continue
        no = m.group(1)
        content = m.group(2).strip()
        content = re.split(r'\n(?=\d+\s*$)|\nHarga Jual', content)[0]

        harga_match = re.findall(r'\b([\d.,]+)\b\s*$', content)
        harga = 0.0
        if harga_match:
            try:
                harga = float(harga_match[-1].replace('.', '').replace(',', '.'))
            except:
                harga = 0.0

        deskripsi = re.sub(r'\b[\d.,]+\b\s*$', '', content, flags=re.DOTALL)
        deskripsi = re.sub(r'\s+', ' ', deskripsi).strip()

        if len(deskripsi) > 5 and harga > 0:
            result.append({
                "No": no,
                "Kode Barang/Jasa": "-",
                "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
            })
    return result

def extract_tabel_auto(txt):
    if re.search(r"\n\s*\d+\s+\d{6}\s+", txt):
        return extract_tabel_dengan_kode(txt)
    else:
        return extract_tabel_tanpa_kode(txt)

# =====================================================
# METADATA & TOTAL
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
# EKSEKUSI: UPLOAD → BACA → PILIH KOLOM → DRAG ORDER → DOWNLOAD
# =====================================================
upl = st.file_uploader("Upload Faktur Pajak (PDF)", type=["pdf"], accept_multiple_files=True)

if upl:
    if st.button("📖 Baca File"):
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

        st.session_state["data_faktur"] = df
        st.success(f"✅ File berhasil dibaca! {len(df)} baris data ditemukan.")
        st.dataframe(df)

# --- Step 2: Pilih Kolom & Urutan ---
if "data_faktur" in st.session_state:
    df = st.session_state["data_faktur"]

    st.markdown("### 🧩 Pilih & Atur Urutan Kolom untuk Diekspor")

    if "kolom_terpilih" not in st.session_state:
        st.session_state["kolom_terpilih"] = list(df.columns)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Pilih Semua Kolom", use_container_width=True):
            st.session_state["kolom_terpilih"] = list(df.columns)
            st.experimental_rerun()
    with col2:
        if st.button("❌ Hapus Semua Kolom", use_container_width=True):
            st.session_state["kolom_terpilih"] = []
            st.experimental_rerun()

    kolom_default = [
        c for c in st.session_state.get("kolom_terpilih", [])
        if c in list(df.columns)
    ]

    kolom_terpilih = st.multiselect(
        "Pilih kolom yang ingin disertakan dalam hasil konversi:",
        options=list(df.columns),
        default=kolom_default,
        key="kolom_multiselect"
    )
    st.session_state["kolom_terpilih"] = kolom_terpilih

    if kolom_terpilih:
        st.markdown("### ↕️ Atur Urutan Kolom (Drag & Drop)")
        ordered_cols = sort_items(
            kolom_terpilih,
            direction="horizontal",
            multi_containers=False,
            key="sortable_cols"
        )

        if ordered_cols:
            df_filtered = df[ordered_cols]
            st.markdown("### 🔍 Preview Hasil Kolom Terpilih (5 Baris Pertama)")
            st.dataframe(df_filtered.head(5))

            buf = BytesIO()
            df_filtered.to_excel(buf, index=False, engine="openpyxl", float_format="%.0f")
            buf.seek(0)

            st.download_button(
                "📥 Konversi & Download Excel",
                buf,
                "rekap_faktur_terpilih.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Belum ada kolom yang dipilih.")
