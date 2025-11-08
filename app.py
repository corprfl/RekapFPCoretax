import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from streamlit_sortables import sort_items

st.set_page_config(page_title="Extractor Faktur Pajak", layout="wide")

# ===== CSS WARNA =====
st.markdown("""
<style>
div.stButton > button:first-child {
    border-radius:8px;font-weight:600;padding:0.5em 1.2em;
}
div[data-testid="stButton"] button:hover{transform:scale(1.03);}
button[kind="primary"],.stDownloadButton button,
#tetapkan-kolom button,#urutan-kolom button,#data-sesuai button{
    background:#2ecc71!important;color:white!important;font-weight:600!important;
}
#pilih-semua button{background:#3498db!important;color:white!important;}
#hapus-semua button{background:#e74c3c!important;color:white!important;}
</style>
""", unsafe_allow_html=True)

# ======= INIT STATE =======
for key, default in {
    "step": None, "data_faktur": None,
    "kolom_terpilih": None, "ordered_cols": None
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# =====================================================
# JUDUL DAN DESKRIPSI
# =====================================================
st.title("Extractor isi Faktur Pajak ke Excel (Anti-Rerun Version)")
st.markdown("""
**Flow:**  
1️⃣ Upload Faktur Pajak → 📖 Baca File  
2️⃣ ✅ Data Sesuai → Pilih Kolom  
3️⃣ ✅ Tetapkan Kolom Terpilih  
4️⃣ ↕️ Urutkan Kolom → ✅ Tetapkan Urutan  
5️⃣ 🔍 Preview → 📥 Download Excel  
---
Semua proses berjalan lokal (client-side).  
**By: Reza Fahlevi Lubis BKP @zavibis**
""")

# =====================================================
# UTILITAS
# =====================================================
bulan_map = {"Januari":"01","Februari":"02","Maret":"03","April":"04",
    "Mei":"05","Juni":"06","Juli":"07","Agustus":"08",
    "September":"09","Oktober":"10","November":"11","Desember":"12"}

def extract(pat,txt,flags=re.DOTALL,default="-"):
    m=re.search(pat,txt,flags)
    return m.group(1).strip() if m else default

def extract_tanggal(txt):
    m=re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})",txt)
    if m: return f"{m.group(2).zfill(2)}/{bulan_map.get(m.group(3),'-')}/{m.group(4)}"
    return "-"

def extract_nitku(txt):
    for i,l in enumerate(txt.splitlines()):
        if "NPWP" in l and i>0:
            prev=txt.splitlines()[i-1]
            m=re.search(r"#(\d{22})",prev)
            if m: return m.group(1)
    return "-"

def extract_total(txt):
    def val(p):
        m=re.search(p,txt,re.DOTALL)
        if not m: return 0.0
        try:return float(m.group(1).replace(".","").replace(",","."))
        except:return 0.0
    return {
        "Dasar Pengenaan Pajak (Total)":val(r"Dasar\s+Pengenaan\s+Pajak\s*([\d.,]+)"),
        "PPN (Total)":val(r"Jumlah\s*PPN.*?([\d.,]+)")
    }

def extract_meta(txt):
    return {
        "Kode dan Nomor Seri Faktur Pajak":extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)",txt),
        "Nama PKP":extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat",txt),
        "NPWP PKP":extract(r"NPWP\s*:\s*([0-9.]+)",txt),
        "Nama Pembeli":extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat",txt),
        "NPWP Pembeli":extract(r"NPWP\s*:\s*([0-9.]+)\s*NIK",txt),
        "NITKU Pembeli":extract_nitku(txt),
        "Tanggal Faktur Pajak":extract_tanggal(txt)
    }

def extract_tabel_auto(txt):
    if re.search(r"\n\s*\d+\s+\d{6}\s+",txt):
        pat=re.compile(r"(\d+)\s+(\d{6})\s+([\s\S]*?)\n\s*([\d.,]+)\s*(?=\n\d+\s+\d{6}|\nHarga Jual|$)",re.M)
        result=[]
        for m in pat.finditer(txt):
            result.append({
                "No":m.group(1),"Kode Barang/Jasa":m.group(2),
                "Nama Barang Kena Pajak / Jasa Kena Pajak":" ".join(m.group(3).split()),
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":
                    float(m.group(4).replace(".","").replace(",","."))
            })
        return result
    else:
        result=[]
        blocks=re.split(r'\n(?=\d+\s*\n)',txt)
        for blk in blocks:
            blk=blk.strip()
            m=re.match(r"(\d+)\s+(.*)",blk,re.DOTALL)
            if not m: continue
            no,content=m.group(1),m.group(2).strip()
            harga_match=re.findall(r'\b([\d.,]+)\b\s*$',content)
            harga=float(harga_match[-1].replace(".","").replace(",","."))
            deskripsi=re.sub(r'\b[\d.,]+\b\s*$','',content).strip()
            if len(deskripsi)>5 and harga>0:
                result.append({
                    "No":no,"Kode Barang/Jasa":"-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak":deskripsi,
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":harga})
        return result

# =====================================================
# STEP 1 — UPLOAD DAN BACA
# =====================================================
upl = st.file_uploader("Upload Faktur Pajak (PDF)",type=["pdf"],accept_multiple_files=True)
if upl and st.button("📖 Baca File",type="primary",key="baca"):
    rows=[]
    for f in upl:
        txt="".join([p.get_text() for p in fitz.open(stream=f.read(),filetype="pdf")])
        meta=extract_meta(txt); meta.update(extract_total(txt)); meta["Nama Asli File"]=f.name
        tgl=meta["Tanggal Faktur Pajak"].split("/")
        meta["Masa"]=tgl[1] if len(tgl)>1 else "-"; meta["Tahun"]=tgl[2] if len(tgl)>2 else "-"
        items=extract_tabel_auto(txt)
        if not items:
            items=[{"No":"-","Kode Barang/Jasa":"-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak":"Tidak terbaca",
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":0.0}]
        for it in items: rows.append({**it,**meta})
    df=pd.DataFrame(rows)
    st.session_state.data_faktur=df; st.session_state.step="cek"
    st.success(f"✅ {len(df)} baris berhasil dibaca."); st.dataframe(df)

# =====================================================
# STEP 2 — KONFIRMASI
# =====================================================
if st.session_state.step=="cek" and st.session_state.data_faktur is not None:
    st.markdown('<div id="data-sesuai">',unsafe_allow_html=True)
    if st.button("✅ Data Sesuai",key="data_ok"): st.session_state.step="pilih"
    st.markdown('</div>',unsafe_allow_html=True)

# =====================================================
# STEP 3 — PILIH KOLOM (PAKAI FORM ANTI-RERUN)
# =====================================================
if st.session_state.step in ["pilih","urut","preview"]:
    df=st.session_state.data_faktur
    kolom_tersedia=list(df.columns)
    if st.session_state.kolom_terpilih is None:
        st.session_state.kolom_terpilih=kolom_tersedia

    with st.form("form_kolom"):
        st.markdown("### 🧩 Pilih Kolom yang Akan Dikonversi")
        col1,col2=st.columns(2)
        with col1:
            st.form_submit_button("✅ Pilih Semua Kolom",on_click=lambda:st.session_state.update(kolom_terpilih=kolom_tersedia))
        with col2:
            st.form_submit_button("❌ Hapus Semua Kolom",on_click=lambda:st.session_state.update(kolom_terpilih=[]))
        kolom_terpilih=st.multiselect("Pilih kolom:",
            options=kolom_tersedia,
            default=[c for c in st.session_state.kolom_terpilih if c in kolom_tersedia],
            key="multi")
        submit=st.form_submit_button("✅ Tetapkan Kolom Terpilih")
        if submit and kolom_terpilih:
            st.session_state.kolom_terpilih=kolom_terpilih
            st.session_state.step="urut"

# =====================================================
# STEP 4 — URUTKAN KOLOM
# =====================================================
if st.session_state.step in ["urut","preview"] and st.session_state.kolom_terpilih:
    st.markdown("### ↕️ Urutkan Kolom (Drag & Drop)")
    ordered=sort_items(st.session_state.kolom_terpilih,direction="horizontal",multi_containers=False,key="sortcols")
    st.session_state.ordered_cols=ordered
    if st.button("↕️ Tetapkan Urutan Kolom",key="btn_order"):
        if ordered: st.session_state.step="preview"

# =====================================================
# STEP 5 — PREVIEW & DOWNLOAD
# =====================================================
if st.session_state.step=="preview" and st.session_state.ordered_cols:
    df=st.session_state.data_faktur
    cols=st.session_state.ordered_cols
    df_filtered=df[cols]
    st.markdown("### 🔍 Preview (5 Baris Pertama)")
    st.dataframe(df_filtered.head(5))
    buf=BytesIO()
    df_filtered.to_excel(buf,index=False,engine="openpyxl",float_format="%.0f"); buf.seek(0)
    st.download_button("📥 Konversi & Download Excel",buf,"rekap_faktur.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
