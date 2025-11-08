import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

st.title("Rekap Faktur Pajak ke Excel (Multi File)")

st.markdown("""
### 📘 Deskripsi Aplikasi
Ekstraktor Faktur Pajak dari PDF ke Excel — membaca deskripsi barang secara utuh dan total dengan akurasi tinggi.

### ⚙️ Cara Pakai
1. Upload satu atau beberapa file PDF Faktur Pajak.  
2. Klik **Eksekusi Convert**.  
3. Hasil rekap muncul dan bisa diunduh ke Excel.

### ⚠️ Disclaimer
- Semua proses dilakukan **lokal di perangkat Anda**.  
- Tidak ada file yang disimpan ke server.  
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
    m=re.search(pattern,text,flags)
    return post(m.group(1)) if m else default

def extract_tanggal(text):
    m=re.search(r"\b([A-Z .,]+),\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})",text)
    if m:
        d=m.group(2).zfill(2)
        b=bulan_map.get(m.group(3),"-")
        y=m.group(4)
        return f"{d}/{b}/{y}"
    return "-"

def extract_nitku_pembeli(text):
    lines=text.splitlines()
    for i,l in enumerate(lines):
        if "NPWP" in l and i>0:
            prev=lines[i-1]
            m=re.search(r"#(\d{22})",prev)
            if m: return m.group(1)
    return "-"

# ---------------------------------------------------------------
# PARSER BARANG DENGAN BATAS BAWAH "Harga Jual / Penggantian"
def extract_tabel_rinci(text):
    """Ambil daftar barang lengkap dan berhenti sebelum bagian total"""
    result=[]
    # potong hanya sebelum bagian total
    section=re.split(r"Harga\s+Jual\s*/\s*Penggantian\s*/?\s*Uang\s*Muka\s*/?\s*Termin", text, maxsplit=1)[0]

    # regex utama: blok dimulai dengan No. sampai No berikutnya
    pattern=re.compile(r"(?m)(No\s*\d+\.?|^\d+\.?)\s*(.*?)(?=(?:No\s*\d+\.?|^\d+\.?|$))", re.DOTALL)
    for m in pattern.finditer(section):
        blok=m.group(0).strip()

        # nomor
        no_match=re.search(r"(\d+)",blok)
        no=no_match.group(1) if no_match else "-"

        # ambil harga terakhir
        harga_match=re.findall(r"([\d.,]+)(?!.*[\d.,])",blok)
        harga=0.0
        if harga_match:
            raw=harga_match[-1].replace(".","").replace(",",".")
            try: harga=float(raw)
            except: harga=0.0

        # deskripsi penuh tapi tanpa No
        deskripsi=re.sub(r"^\s*No\s*\d+\.?\s*","",blok)
        deskripsi=re.sub(r"\s+"," ",deskripsi)
        deskripsi=deskripsi.strip()

        # skip kalau sudah bagian total
        if re.search(r"Harga\s+Jual\s*/|Dasar\s+Pengenaan",deskripsi,re.I):
            continue

        if deskripsi:
            result.append({
                "No":no,
                "Kode Barang/Jasa":"-",
                "Nama Barang Kena Pajak / Jasa Kena Pajak":deskripsi,
                "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":harga
            })
    return result

# ---------------------------------------------------------------
def extract_total_section(text):
    def val(p):
        m=re.search(p,text,re.DOTALL)
        if not m: return 0.0
        raw=m.group(1).strip().replace(".","").replace(",",".")
        try: return float(raw)
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
    if not kode or len(kode)<3:
        return "-","-"
    return kode[:2], ("Normal" if kode[2]=="0" else "Pengganti")

# ---------------------------------------------------------------
uploaded_files=st.file_uploader("Upload PDF Faktur Pajak",type=["pdf"],accept_multiple_files=True)
if uploaded_files:
    if st.button("Eksekusi Convert"):
        rows=[]
        for f in uploaded_files:
            data=f.read()
            with fitz.open(stream=data,filetype="pdf") as doc:
                text="".join(p.get_text() for p in doc)

            meta=extract_data_from_text(text)
            kode=meta.get("Kode dan Nomor Seri Faktur Pajak","")
            kf,stt=extract_kode_status(kode)
            meta.update({"Kode Faktur":kf,"Status Faktur":stt,"Nama Asli File":f.name})
            meta.update(extract_total_section(text))

            tgl=meta["Tanggal Faktur Pajak"].split("/")
            meta["Masa"]=tgl[1] if len(tgl)>1 else "-"
            meta["Tahun"]=tgl[2] if len(tgl)>2 else "-"

            items=extract_tabel_rinci(text)
            if not items:
                items=[{
                    "No":"-","Kode Barang/Jasa":"-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak":"-",
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)":0.0
                }]
            for it in items:
                rows.append({**it,**meta})

        df=pd.DataFrame(rows)
        numcols=[
            "Total Harga Jual / Penggantian / Uang Muka / Termin",
            "Dikurangi Potongan Harga (Total)",
            "Dikurangi Uang Muka yang telah diterima (Total)",
            "Dasar Pengenaan Pajak (Total)",
            "PPN (Total)","Jumlah PPnBM (Total)",
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)"
        ]
        for c in numcols:
            if c in df.columns:
                df[c]=pd.to_numeric(df[c],errors="coerce").fillna(0.0)

        order=[
            "Kode Faktur","Status Faktur","Kode dan Nomor Seri Faktur Pajak",
            "Nama Pengusaha Kena Pajak","NPWP Pengusaha Kena Pajak",
            "Nama Pembeli Barang/Jasa","NPWP Pembeli Barang/Jasa","NITKU Pembeli",
            "Tanggal Faktur Pajak","Kota",
            "No","Kode Barang/Jasa","Nama Barang Kena Pajak / Jasa Kena Pajak",
            "Harga Jual / Penggantian / Uang Muka / Termin (Rp)",
            "Total Harga Jual / Penggantian / Uang Muka / Termin",
            "Dikurangi Potongan Harga (Total)",
            "Dikurangi Uang Muka yang telah diterima (Total)",
            "Dasar Pengenaan Pajak (Total)","PPN (Total)","Jumlah PPnBM (Total)",
            "Referensi","Penandatangan","Nama Asli File","Masa","Tahun"
        ]
        df=df[[c for c in order if c in df.columns]]

        st.success("✅ Parsing faktur berhasil tanpa bagian total ikut terbaca!")
        st.dataframe(df)

        buf=BytesIO()
        df.to_excel(buf,index=False,engine="openpyxl",float_format="%.0f")
        buf.seek(0)
        st.download_button("📥 Download Rekap Excel",buf,"rekap_faktur.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
