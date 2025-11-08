import streamlit as st
import pandas as pd
import fitz, re
from io import BytesIO

st.title("Rekap Faktur Pajak ke Excel (Multi File)")

st.markdown("""
### 📘 Deskripsi
Konversi PDF Faktur Pajak ke Excel — membaca deskripsi barang/jasa secara utuh, presisi antara batas atas (setelah NITKU) dan batas bawah (sebelum total).  
Semua proses dilakukan **lokal**, tidak ada file disimpan.

---
**By Reza Fahlevi Lubis BKP @zavibis**
""")

# =========================================================
bulan_map = {
    "Januari":"01","Februari":"02","Maret":"03","April":"04","Mei":"05",
    "Juni":"06","Juli":"07","Agustus":"08","September":"09","Oktober":"10",
    "November":"11","Desember":"12"
}

def extract(pat,txt,flags=re.DOTALL,default="-"):
    m=re.search(pat,txt,flags)
    return m.group(1).strip() if m else default

def extract_tanggal(txt):
    m=re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})",txt)
    if m:
        d=m.group(2).zfill(2); b=bulan_map.get(m.group(3),"-"); y=m.group(4)
        return f"{d}/{b}/{y}"
    return "-"

def extract_nitku(txt):
    for i,l in enumerate(txt.splitlines()):
        if "NPWP" in l and i>0:
            prev=txt.splitlines()[i-1]
            m=re.search(r"#(\d{22})",prev)
            if m: return m.group(1)
    return "-"

# =========================================================
def extract_tabel_rinci(txt):
    """Ambil daftar barang utuh (No x + seluruh deskripsi + harga terakhir)"""
    result = []

    # Tentukan batas atas & bawah area barang
    atas = re.search(r"(#\d{22}.*?\n[-=]{3,}\s*\n)", txt, re.DOTALL)
    bawah = re.search(r"\n\s*Harga\s+Jual\s*/\s*Penggantian\s*/?\s*Uang\s*Muka\s*/?\s*Termin", txt)
    if atas and bawah:
        area = txt[atas.end():bawah.start()]
    elif atas:
        area = txt[atas.end():]
    elif bawah:
        area = txt[:bawah.start()]
    else:
        area = txt

    # Gabungkan semua baris menjadi satu paragraf agar deskripsi tidak pecah
    lines = [re.sub(r"\s+", " ", l.strip()) for l in area.splitlines() if l.strip()]
    area = " ".join(lines)

    # Ambil setiap blok mulai dari "No x" sampai "No berikutnya"
    pattern = re.compile(r"(No\s*\d+\.?\s.*?)(?=No\s*\d+\.?|$)", re.DOTALL)
    for match in pattern.finditer(area):
        blok = match.group(1).strip()
        if not blok:
            continue

        # Nomor urut
        no_match = re.search(r"No\s*(\d+)", blok)
        no = no_match.group(1) if no_match else "-"

        # Harga terakhir
        harga_match = re.findall(r"([\d.,]+)(?!.*[\d.,])", blok)
        harga = 0.0
        if harga_match:
            try:
                harga = float(harga_match[-1].replace(".","").replace(",","."))
            except:
                harga = 0.0

        # Bersihkan "No x" di depan & jadikan deskripsi satu kalimat rapi
        deskripsi = re.sub(r"^No\s*\d+\.?\s*", "", blok)
        deskripsi = re.sub(r"\s+", " ", deskripsi).strip()

        if deskripsi and not re.search(r"Harga\s+Jual|Dasar\s+Pengenaan", deskripsi, re.I):
            result.append({
                "No": no,
                "Kode Barang/Jasa": "-",
                "Nama Barang Kena Pajak / Jasa Kena Pajak": deskripsi,
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": harga
            })
    return result

# =========================================================
def extract_total(txt):
    def val(p):
        m=re.search(p,txt,re.DOTALL)
        if not m: return 0.0
        try: return float(m.group(1).replace(".","").replace(",","."))
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

# =========================================================
def extract_meta(txt):
    return {
        "Kode dan Nomor Seri Faktur Pajak":extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)",txt),
        "Nama PKP":extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat",txt),
        "NPWP PKP":extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9.]+)",txt),
        "Nama Pembeli":extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat",txt),
        "NPWP Pembeli":extract(r"NPWP\s*:\s*([0-9.]+)\s*NIK",txt),
        "NITKU Pembeli":extract_nitku(txt),
        "Kota":extract(r"\n([A-Z .,]+),\s*\d{1,2}\s+\w+\s+\d{4}",txt),
        "Tanggal Faktur Pajak":extract_tanggal(txt),
        "Penandatangan":extract(r"Ditandatangani secara elektronik\n(.*?)\n",txt)
    }

def kode_status(kode):
    if not kode or len(kode)<3: return "-","-"
    return kode[:2],("Normal" if kode[2]=="0" else "Pengganti")

# =========================================================
upl=st.file_uploader("Upload PDF Faktur Pajak",type=["pdf"],accept_multiple_files=True)
if upl:
    if st.button("Eksekusi Convert"):
        rows=[]
        for f in upl:
            with fitz.open(stream=f.read(),filetype="pdf") as doc:
                txt="".join(p.get_text() for p in doc)

            meta=extract_meta(txt)
            kode=meta["Kode dan Nomor Seri Faktur Pajak"]
            kf,stt=kode_status(kode)
            meta.update({"Kode Faktur":kf,"Status Faktur":stt,"Nama Asli File":f.name})
            meta.update(extract_total(txt))
            tgl=meta["Tanggal Faktur Pajak"].split("/")
            meta["Masa"]=tgl[1] if len(tgl)>1 else "-"
            meta["Tahun"]=tgl[2] if len(tgl)>2 else "-"

            items=extract_tabel_rinci(txt)
            if not items:
                items=[{"No":"-","Kode Barang/Jasa":"-",
                        "Nama Barang Kena Pajak / Jasa Kena Pajak":"-",
                        "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":0.0}]
            for it in items: rows.append({**it,**meta})

        df=pd.DataFrame(rows)
        for c in df.columns:
            if "Rp" in c or "Total" in c: 
                df[c]=pd.to_numeric(df[c],errors="coerce").fillna(0.0)

        st.success("✅ Parsing faktur berhasil — deskripsi utuh & batas atas/bawah akurat.")
        st.dataframe(df)

        buf=BytesIO()
        df.to_excel(buf,index=False,engine="openpyxl",float_format="%.0f"); buf.seek(0)
        st.download_button("📥 Download Rekap Excel",buf,"rekap_faktur.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
