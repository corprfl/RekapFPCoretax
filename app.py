import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from streamlit_sortables import sort_items

st.set_page_config(page_title="Extractor Faktur Pajak", layout="wide")

# ========== STYLE ==========
st.markdown("""
<style>
.stepper {
    display: flex; justify-content: space-between; margin: 15px 0 30px 0;
}
.step {
    flex: 1; text-align: center; padding: 6px 10px; border-radius: 6px; 
    font-weight: 600; margin: 0 4px; color: white;
}
.active { background-color: #2ecc71; }
.done { background-color: #3498db; }
.pending { background-color: #555; }

div.stButton > button:first-child {
    border-radius: 8px;
    font-weight: 600;
    padding: 0.6em 1.2em;
    border: none;
}
div[data-testid="stButton"] button:hover { transform: scale(1.03); }

button[kind="primary"], .stDownloadButton button {
    background-color: #2ecc71 !important;
    color: white !important;
}
#pilih-semua button {
    background-color: #3498db !important;
    color: white !important;
}
#hapus-semua button {
    background-color: #e74c3c !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# ========== FUNGSI UTAMA ==========
bulan_map = {"Januari":"01","Februari":"02","Maret":"03","April":"04",
    "Mei":"05","Juni":"06","Juli":"07","Agustus":"08",
    "September":"09","Oktober":"10","November":"11","Desember":"12"}

def extract(pat, txt, flags=re.DOTALL, default="-"):
    m = re.search(pat, txt, flags)
    return m.group(1).strip() if m else default

def extract_tanggal(txt):
    m = re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", txt)
    if m:
        return f"{m.group(2).zfill(2)}/{bulan_map.get(m.group(3), '-')}/{m.group(4)}"
    return "-"

def extract_meta(txt):
    return {
        "Kode dan Nomor Seri Faktur Pajak": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", txt),
        "Nama PKP": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", txt),
        "NPWP PKP": extract(r"NPWP\s*:\s*([0-9.]+)", txt),
        "Nama Pembeli": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", txt),
        "NPWP Pembeli": extract(r"NPWP\s*:\s*([0-9.]+)", txt),
        "Tanggal Faktur Pajak": extract_tanggal(txt)
    }

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
            if not m: continue
            no, content = m.group(1), m.group(2).strip()
            harga_match = re.findall(r'\b([\d.,]+)\b\s*$', content)
            harga = float(harga_match[-1].replace('.', '').replace(',', '.')) if harga_match else 0
            deskripsi = re.sub(r'\b[\d.,]+\b\s*$', '', content).strip()
            if deskripsi and harga>0:
                result.append({
                    "No": no, "Kode Barang/Jasa": "-", 
                    "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
                })
        return result

# ========== JUDUL ==========
st.title("Extractor isi Faktur Pajak ke Excel (Wizard Kolom)")

# ===== STEP INDICATOR =====
step = st.session_state.get("step", "upload")
steps = ["upload", "baca", "pilih", "urut", "preview"]
labels = ["Upload", "Baca", "Pilih Kolom", "Urutkan", "Preview"]

st.markdown('<div class="stepper">' + "".join(
    f'<div class="step {"active" if s==step else "done" if steps.index(s)<steps.index(step) else "pending"}">{i+1}. {labels[i]}</div>'
    for i,s in enumerate(steps)
) + '</div>', unsafe_allow_html=True)

# ===== UPLOAD =====
upl = st.file_uploader("Upload Faktur Pajak (PDF)", type=["pdf"], accept_multiple_files=True)

if upl and st.button("📖 Baca File", type="primary"):
    rows = []
    for f in upl:
        txt = "".join([p.get_text() for p in fitz.open(stream=f.read(), filetype="pdf")])
        meta = extract_meta(txt)
        items = extract_tabel_auto(txt)
        if not items:
            items = [{"No": "-", "Kode Barang/Jasa": "-", "Nama Barang Kena Pajak / Jasa Kena Pajak": "Tidak terbaca", 
                      "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0.0}]
        for it in items: rows.append({**it, **meta})
    df = pd.DataFrame(rows)
    st.session_state["data_faktur"] = df
    st.session_state["step"] = "pilih"
    st.success(f"✅ {len(df)} baris berhasil dibaca.")
    st.dataframe(df)

# ===== PILIH KOLOM =====
if st.session_state.get("step") in ["pilih", "urut", "preview"] and "data_faktur" in st.session_state:
    df = st.session_state["data_faktur"]
    st.markdown("### 🧩 Pilih Kolom yang Akan Dikonversi")
    kolom_tersedia = list(df.columns)
    kolom_terpilih = st.session_state.get("kolom_terpilih", kolom_tersedia)

    c1,c2 = st.columns(2)
    with c1:
        st.markdown('<div id="pilih-semua">', unsafe_allow_html=True)
        if st.button("✅ Pilih Semua Kolom", use_container_width=True):
            st.session_state["kolom_terpilih"] = kolom_tersedia
            kolom_terpilih = kolom_tersedia
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div id="hapus-semua">', unsafe_allow_html=True)
        if st.button("❌ Hapus Semua Kolom", use_container_width=True):
            st.session_state["kolom_terpilih"] = []
            kolom_terpilih = []
        st.markdown('</div>', unsafe_allow_html=True)

    kolom_terpilih = st.multiselect(
        "Pilih kolom:",
        options=kolom_tersedia,
        default=[c for c in kolom_terpilih if c in kolom_tersedia],
        key="kolom_multiselect"
    )
    st.session_state["kolom_terpilih"] = kolom_terpilih

    if st.button("✅ Tetapkan Kolom Terpilih") and kolom_terpilih:
        st.session_state["step"] = "urut"

# ===== URUTKAN =====
if st.session_state.get("step") in ["urut", "preview"] and st.session_state.get("kolom_terpilih"):
    st.markdown("### ↕️ Urutkan Kolom (Drag & Drop)")
    ordered = sort_items(
        st.session_state["kolom_terpilih"],
        direction="horizontal",
        multi_containers=False,
        key="sortable_cols"
    )
    st.session_state["ordered_cols"] = ordered
    if st.button("↕️ Tetapkan Urutan Kolom") and ordered:
        st.session_state["step"] = "preview"

# ===== PREVIEW =====
if st.session_state.get("step") == "preview" and st.session_state.get("ordered_cols"):
    df = st.session_state["data_faktur"][st.session_state["ordered_cols"]]
    st.markdown("### 🔍 Preview Hasil Kolom Terpilih (5 Baris Pertama)")
    st.dataframe(df.head(5))
    buf = BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl", float_format="%.0f")
    buf.seek(0)
    st.download_button("📥 Konversi & Download Excel", buf, "rekap_faktur.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
