import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from streamlit_sortables import sort_items

st.set_page_config(
    page_title="Extractor Faktur Pajak Coretax",
    layout="wide"
)

# =====================================================
# ====== CSS CUSTOM ======
# =====================================================
st.markdown("""
<style>
div.stButton > button:first-child {
    border-radius:8px;
    font-weight:600;
    padding:0.5em 1.2em;
    font-size:15px;
}
div[data-testid="stButton"] button:hover {
    transform:scale(1.03);
}
button[kind="primary"], .stDownloadButton button {
    background:#2ecc71!important;
    color:white!important;
    font-weight:700!important;
}
hr {
    margin: 1.5em 0;
}
.footer {
    text-align:center;
    font-size:13px;
    color:#777;
    margin-top:40px;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# ===== INIT SESSION STATE =====
# =====================================================
for k, v in {
    "step": None,
    "data_faktur": None,
    "ordered_cols": None
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# =====================================================
# ================= HEADER ============================
# =====================================================
st.title("🐾 Extractor Isi Faktur Pajak Coretax ke Excel")

st.markdown("""
**By Reza Fahlevi Lubis BKP – @zavibis**

Aplikasi ini membantu mengekstrak **detail baris dan metadata Faktur Pajak Coretax (PDF)**  
menjadi file **Excel terstruktur**, siap untuk analisis, rekonsiliasi, dan arsip perpajakan.
""")

st.markdown("---")

# =====================================================
# ================= DESKRIPSI =========================
# =====================================================
st.subheader("📌 Deskripsi Aplikasi")
st.markdown("""
Extractor ini dirancang untuk membaca **PDF Faktur Pajak Coretax DJP** dan menghasilkan:

- Data **header faktur** (PKP, pembeli, NPWP, NITKU, tanggal, referensi)
- Data **detail barang/jasa per baris**
- Nilai **DPP, PPN, PPnBM, dan total**
- Output **Excel (.xlsx)** dengan urutan kolom yang bisa diatur manual

Semua proses dilakukan **lokal di browser**, tanpa menyimpan data ke server.
""")

# =====================================================
# ================= PETUNJUK ==========================
# =====================================================
st.subheader("🧭 Petunjuk Penggunaan")
st.markdown("""
1. **Upload** satu atau beberapa file PDF Faktur Pajak Coretax  
2. Klik **📖 Baca File** untuk mengekstrak data  
3. **Periksa hasil ekstraksi** pada tabel preview  
4. Klik **✔️ Data Sudah Sesuai** jika sudah benar  
5. **Atur urutan kolom** dengan drag & drop  
6. Klik **📥 Download Excel** untuk mengunduh hasil rekap  

💡 Tips:
- Gunakan PDF asli hasil unduhan DJP Coretax
- Pastikan teks di PDF **bukan hasil scan**
""")

# =====================================================
# ================= DISCLAIMER ========================
# =====================================================
st.subheader("⚠️ Disclaimer")
st.markdown("""
- Aplikasi ini **bukan produk resmi DJP**
- Hasil ekstraksi **perlu diverifikasi ulang** sebelum digunakan untuk pelaporan pajak
- Akurasi bergantung pada **konsistensi format PDF Coretax**
- Penulis tidak bertanggung jawab atas kesalahan akibat penggunaan tanpa verifikasi
""")

st.markdown("---")

# =====================================================
# ================= UTILITAS ==========================
# =====================================================
bulan_map = {
    "Januari":"01","Februari":"02","Maret":"03","April":"04",
    "Mei":"05","Juni":"06","Juli":"07","Agustus":"08",
    "September":"09","Oktober":"10","November":"11","Desember":"12"
}

def extract(pat, txt, flags=re.DOTALL, default="-"):
    m = re.search(pat, txt, flags)
    return m.group(1).strip() if m else default

def extract_tanggal(txt):
    m = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", txt)
    if m:
        return f"{m.group(2).zfill(2)}/{bulan_map.get(m.group(3),'-')}/{m.group(4)}"
    return "-"

def extract_nitku(txt):
    lines = txt.splitlines()
    for i, l in enumerate(lines):
        if "NPWP" in l and i > 0:
            m = re.search(r"#(\d{22})", lines[i-1])
            if m:
                return m.group(1)
    return "-"

def extract_referensi(txt):
    m = re.search(r"\(Referensi:\s*(.*?)\)", txt, re.IGNORECASE)
    return m.group(1).strip() if m else "-"

def extract_total(txt):
    def val(p):
        m = re.search(p, txt, re.DOTALL)
        if not m:
            return 0.0
        try:
            return float(m.group(1).replace(".","").replace(",","."))
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
        "Penandatangan": extract(r"Ditandatangani secara elektronik\n(.*?)\n", txt),
        "Keterangan Tambahan": extract(r"Keterangan\s*:\s*(.*)", txt),
        "Nomor Referensi": extract_referensi(txt)
    }

def extract_tabel_auto(txt):
    pat = re.compile(
        r"(\d+)\s+(\d{6})\s+([\s\S]*?)\n\s*([\d.,]+)\s*(?=\n\d+\s+\d{6}|\nHarga Jual|$)",
        re.M,
    )
    res = []
    for m in pat.finditer(txt):
        res.append({
            "No": m.group(1),
            "Kode Barang/Jasa": m.group(2),
            "Nama Barang Kena Pajak / Jasa Kena Pajak": " ".join(m.group(3).split()),
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":
                float(m.group(4).replace(".","").replace(",","."))
        })
    return res

# =====================================================
# ================= STEP 1 ============================
# =====================================================
upl = st.file_uploader(
    "📄 Upload Faktur Pajak Coretax (PDF)",
    type=["pdf"],
    accept_multiple_files=True
)

if upl and st.button("📖 Baca File", type="primary"):
    rows = []

    for f in upl:
        txt = "".join([p.get_text() for p in fitz.open(stream=f.read(), filetype="pdf")])
        meta = extract_meta(txt)
        meta.update(extract_total(txt))
        meta["Nama Asli File"] = f.name

        tgl = meta["Tanggal Faktur Pajak"].split("/")
        meta["Masa"] = tgl[1] if len(tgl) > 1 else "-"
        meta["Tahun"] = tgl[2] if len(tgl) > 2 else "-"

        items = extract_tabel_auto(txt) or [{
            "No":"-",
            "Kode Barang/Jasa":"-",
            "Nama Barang Kena Pajak / Jasa Kena Pajak":"Tidak terbaca",
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":0.0
        }]

        for it in items:
            rows.append({**it, **meta})

    st.session_state.data_faktur = pd.DataFrame(rows)
    st.session_state.step = "cek"
    st.success("✅ Data berhasil diekstrak")
    st.dataframe(st.session_state.data_faktur)

# =====================================================
# ================= STEP 2 ============================
# =====================================================
if st.session_state.step == "cek":
    if st.button("✔️ Data Sudah Sesuai"):
        st.session_state.step = "urut"

# =====================================================
# ================= STEP 3 ============================
# =====================================================
if st.session_state.step in ["urut", "preview"]:
    df = st.session_state.data_faktur
    st.markdown("### ↕️ Atur Urutan Kolom")
    ordered = sort_items(list(df.columns), direction="horizontal")
    st.session_state.ordered_cols = ordered

    if st.button("✔️ Tetapkan Urutan Kolom"):
        st.session_state.step = "preview"

# =====================================================
# ================= STEP 4 ============================
# =====================================================
if st.session_state.step == "preview":
    df = st.session_state.data_faktur[st.session_state.ordered_cols]
    st.markdown("### 🔍 Preview Data")
    st.dataframe(df.head())

    buf = BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl", float_format="%.0f")
    buf.seek(0)

    st.download_button(
        "📥 Download Excel",
        buf,
        "rekap_detail_faktur_coretax.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# ================= FOOTER ============================
# =====================================================
st.markdown("""
<div class="footer">
    © 2025 — By Reza Fahlevi Lubis BKP (@zavibis)<br>
    Extractor Faktur Pajak Coretax • Streamlit App
</div>
""", unsafe_allow_html=True)
