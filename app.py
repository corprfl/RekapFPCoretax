# =====================================================
# EKSEKUSI DENGAN FITUR PILIH KOLOM + PREVIEW
# =====================================================
upl = st.file_uploader("Upload Faktur Pajak (PDF)", type=["pdf"], accept_multiple_files=True)

# --- Step 1: Baca File ---
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
                items = [{
                    "No": "-",
                    "Kode Barang/Jasa": "-",
                    "Nama Barang Kena Pajak / Jasa Kena Pajak": "Tidak terbaca",
                    "Harga Jual / Penggantian / Uang Muka / Termin (Rp)": 0.0
                }]
            for it in items:
                rows.append({**it, **meta})

        df = pd.DataFrame(rows)
        for c in df.columns:
            if "Rp" in c or "Total" in c:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        st.session_state["data_faktur"] = df
        st.success(f"✅ File berhasil dibaca! {len(df)} baris data ditemukan.")
        st.dataframe(df)

# --- Step 2: Pilih Kolom ---
if "data_faktur" in st.session_state:
    df = st.session_state["data_faktur"]

    st.markdown("### 🧩 Pilih Kolom untuk Diekspor")

    # default semua kolom
    if "kolom_terpilih" not in st.session_state:
        st.session_state["kolom_terpilih"] = list(df.columns)

    # tombol pilih semua / hapus semua
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Pilih Semua Kolom"):
            st.session_state["kolom_terpilih"] = list(df.columns)
    with col2:
        if st.button("❌ Hapus Semua Kolom"):
            st.session_state["kolom_terpilih"] = []

    kolom_terpilih = st.multiselect(
        "Pilih kolom yang ingin disertakan dalam hasil konversi:",
        options=list(df.columns),
        default=st.session_state["kolom_terpilih"],
        key="kolom_multiselect"
    )

    # simpan pilihan agar persistent
    st.session_state["kolom_terpilih"] = kolom_terpilih

    # --- Step 3: Preview hasil kolom terpilih ---
    if kolom_terpilih:
        df_filtered = df[kolom_terpilih]
        st.markdown("### 🔍 Preview Hasil Kolom Terpilih (5 Baris Pertama)")
        st.dataframe(df_filtered.head(5))

        # --- Step 4: Konversi & Download ---
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
