import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from streamlit_sortables import sort_items

# =====================================================
# 🧾 EXTRACTOR ISI FAKTUR PAJAK KE EXCEL
# =====================================================
st.set_page_config(page_title="Extractor Faktur Pajak", layout="wide")

# ====== CSS CUSTOM UNTUK WARNA TOMBOL ======
st.markdown("""
<style>
div.stButton > button:first-child {
    border-radius: 10px;
    font-weight: 600;
    padding: 0.6em 1.2em;
    border: none;
}
div[data-testid="stButton"] button:hover {
    transform: scale(1.03);
}

/* Tombol Hijau */
button[kind="primary"], #tetapkan-kolom button, #urutan-kolom button, .stDownloadButton button {
    background-color: #2ecc71 !important;
    color: white !important;
    font-weight: 600 !important;
}

/* Tombol Biru */
#pilih-semua button {
    background-color: #3498db !important;
    color: white !important;
}

/* Tombol Merah */
#hapus-semua button {
    background-color: #e74c3c !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# JUDUL DAN PENJELASAN
# =====================================================
st.title("Extractor isi Faktur Pajak ke Excel (Wizard Pilih & Urut Kolom)")

st.markdown("""
### 🧭 Alur Penggunaan
1️⃣ **Upload Faktur Pajak (PDF)**  
2️⃣ Tekan **📖 Baca File**  
3️⃣ Pilih kolom yang akan diekspor → klik **✅ Tetapkan Kolom Terpilih**  
4️⃣ Urutkan kolom (drag & drop) → klik **↕️ Tetapkan Urutan Kolom**  
5️⃣ Tampilkan preview & tekan **📥 Download Excel**

Semua proses berjalan di perangkat Anda.  
**Tidak ada file yang dikirim ke server.**
---
**By: Reza Fahlevi Lubis BKP @zavibis**
""")

# =====================================================
# UTILITAS EKSTRAKSI
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

def extract_tabel_auto(txt):
    if re.search(r"\n\s*\d+\s+\d{6}\s+", txt):
        pat = re.compile(
            r"(\d+)\s+(\d{6})\s+([\s\S]*?)\n\s*([\d.,]+)\s*(?=\n\d+\s+\d{6}|\nHarga Jual|$)",
            re.MULTILINE
        )
        result = []
        for m in pat.finditer(txt):
            result.append({
                "No": m.group(1),
                "Kode Barang/Jasa": m.group(2),
                "Nama Barang Kena Pajak / Jasa Kena Pajak": " ".join(m.group(3).split()),
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": float(m.group(4).replace(".", "").replace(",", "."))
            })
        return result
    else:
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
                    pass
            deskripsi = re.sub(r'\b[\d.,]+\b\s*$', '', content).strip()
            if len(deskripsi) > 5 and harga > 0:
                result.append({
                    "No": no,
                    "Kode Barang/Jasa": "-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
                })
        return result

def extract_total(txt):
    def val(p):
        m = re.search(p, txt)
        if not m: return 0.0
        try: return float(m.group(1).replace('.', '').replace(',', '.'))
        except: return 0.0
    return {
        "Dasar Pengenaan Pajak (Total)": val(r"Dasar\s+Pengenaan\s+Pajak\s*([\d.,]+)"),
        "PPN (Total)": val(r"Jumlah\s*PPN.*?([\d.,]+)")
    }

def extract_meta(txt):
    return {
        "Kode dan Nomor Seri Faktur Pajak": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", txt),
        "Nama PKP": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", txt),
        "NPWP PKP": extract(r"NPWP\s*:\s*([0-9.]+)", txt),
        "Nama Pembeli": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", txt),
        "NPWP Pembeli": extract(r"NPWP\s*:\s*([0-9.]+)", txt),
        "NITKU Pembeli": extract_nitku(txt),
        "Tanggal Faktur Pajak": extract_tanggal(txt),
    }

def kode_status(k):
    if not k or len(k) < 3: return "-", "-"
    return k[:2], ("Normal" if k[2] == "0" else "Pengganti")

# =====================================================
# STEP 1 — UPLOAD & BACA FILE
# =====================================================
upl = st.file_uploader("Upload Faktur Pajak (PDF)", type=["pdf"], accept_multiple_files=True)

if upl and st.button("📖 Baca File", type="primary"):
    rows = []
    for f in upl:
        txt = "".join([p.get_text() for p in fitz.open(stream=f.read(), filetype="pdf")])
        meta = extract_meta(txt)
        kode = meta["Kode dan Nomor Seri Faktur Pajak"]
        kf, stt = kode_status(kode)
        meta.update({"Kode Faktur": kf, "Status Faktur": stt, "Nama Asli File": f.name})
        meta.update(extract_total(txt))
        items = extract_tabel_auto(txt)
        if not items:
            items = [{"No": "-", "Kode Barang/Jasa": "-", "Nama Barang Kena Pajak / Jasa Kena Pajak": "Tidak terbaca", "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0.0}]
        for it in items:
            rows.append({**it, **meta})
    df = pd.DataFrame(rows)
    st.session_state["data_faktur"] = df
    st.session_state["step"] = "pilih"
    st.success(f"✅ {len(df)} baris berhasil dibaca.")
    st.dataframe(df)

# =====================================================
# STEP 2 — PILIH KOLOM
# =====================================================
if st.session_state.get("step") in ["pilih", "urut", "preview"] and "data_faktur" in st.session_state:
    df = st.session_state["data_faktur"]

    st.markdown("### 🧩 Pilih Kolom yang Akan Dikonversi")
    kolom_tersedia = list(df.columns)
    kolom_simpan = st.session_state.get("kolom_terpilih", [])
    kolom_default = [c for c in kolom_simpan if c in kolom_tersedia]

    col1, col2 = st.columns(2)
    with col1:
        with st.container():
            st.markdown('<div id="pilih-semua">', unsafe_allow_html=True)
            if st.button("✅ Pilih Semua Kolom", use_container_width=True):
                st.session_state["kolom_terpilih"] = list(df.columns)
            st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        with st.container():
            st.markdown('<div id="hapus-semua">', unsafe_allow_html=True)
            if st.button("❌ Hapus Semua Kolom", use_container_width=True):
                st.session_state["kolom_terpilih"] = []
            st.markdown('</div>', unsafe_allow_html=True)

    kolom_terpilih = st.multiselect(
        "Pilih kolom:",
        options=kolom_tersedia,
        default=kolom_default,
        key="kolom_multiselect"
    )
    st.session_state["kolom_terpilih"] = kolom_terpilih

    st.markdown('<div id="tetapkan-kolom">', unsafe_allow_html=True)
    if st.button("✅ Tetapkan Kolom Terpilih"):
        if kolom_terpilih:
            st.session_state["step"] = "urut"
    st.markdown('</div>', unsafe_allow_html=True)

# =====================================================
# STEP 3 — URUTKAN KOLOM
# =====================================================
if st.session_state.get("step") in ["urut", "preview"] and st.session_state.get("kolom_terpilih"):
    st.markdown("### ↕️ Urutkan Kolom (Drag & Drop)")
    ordered_cols = sort_items(
        st.session_state["kolom_terpilih"],
        direction="horizontal",
        multi_containers=False,
        key="sortable_cols"
    )
    st.session_state["ordered_cols"] = ordered_cols

    st.markdown('<div id="urutan-kolom">', unsafe_allow_html=True)
    if st.button("↕️ Tetapkan Urutan Kolom"):
        if ordered_cols:
            st.session_state["step"] = "preview"
    st.markdown('</div>', unsafe_allow_html=True)

# =====================================================
# STEP 4 — PREVIEW & DOWNLOAD
# =====================================================
if st.session_state.get("step") == "preview" and st.session_state.get("ordered_cols"):
    df = st.session_state["data_faktur"]
    ordered_cols = st.session_state["ordered_cols"]
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
