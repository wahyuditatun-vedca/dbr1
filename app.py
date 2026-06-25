"""
app.py
Streamlit app: Sistem Penomoran & Cetak Label Daftar Barang Ruangan (DBR).

Alur:
1. Upload PDF DBR
2. Sistem coba ekstraksi otomatis (tabel -> teks -> OCR fallback)
3. Pengguna KONFIRMASI/PETAKAN kolom mana = Kode Barang / Nama Barang / NUP / Ruangan
   (wajib manual, karena hasil ekstraksi otomatis dari PDF pemerintah sangat
   bervariasi formatnya dan tidak bisa diasumsikan benar tanpa verifikasi)
4. Rekap per jenis barang (Kode Barang -> daftar NUP)
5. Cetak label: urutkan (Kode Barang, NUP), generate PDF grid 12 label/lembar
"""

import streamlit as st
import pandas as pd

from pdf_parser import auto_extract, guess_column_mapping
from label_pdf import generate_label_pdf

st.set_page_config(page_title="Penomoran Label Barang Ruangan", layout="wide")

st.title("📋 Sistem Penomoran & Cetak Label Daftar Barang Ruangan")
st.caption(
    "Upload PDF Daftar Barang Ruangan (DBR) → kelompokkan per Kode Barang & NUP "
    "→ cetak label fisik (12 label per lembar A4)."
)

# ---------------------------------------------------------------------------
# Session state init
# ---------------------------------------------------------------------------
if "raw_df" not in st.session_state:
    st.session_state.raw_df = pd.DataFrame()
if "extract_method" not in st.session_state:
    st.session_state.extract_method = None

# ---------------------------------------------------------------------------
# Step 1: Upload & ekstraksi
# ---------------------------------------------------------------------------
st.header("1️⃣ Upload PDF")
uploaded = st.file_uploader("Pilih file PDF Daftar Barang Ruangan", type=["pdf"])

col_a, col_b = st.columns([1, 1])
with col_a:
    force_ocr = st.checkbox(
        "Paksa pakai OCR (kalau hasil ekstraksi otomatis kosong/salah parah)",
        value=False,
        help="Aktifkan kalau PDF ini hasil scan/foto dan ekstraksi tabel/teks gagal.",
    )

if uploaded is not None:
    pdf_bytes = uploaded.read()

    if st.button("🔍 Ekstrak Data dari PDF", type="primary"):
        with st.spinner("Mengekstrak data dari PDF..."):
            if force_ocr:
                from pdf_parser import extract_via_ocr, extract_via_text_split
                ocr_text, err = extract_via_ocr(pdf_bytes)
                if err:
                    st.error(err)
                    result = {"df": pd.DataFrame(), "method": "gagal", "warning": err}
                else:
                    df_ocr = extract_via_text_split(ocr_text)
                    result = {
                        "df": df_ocr,
                        "method": "ocr (manual)",
                        "warning": "Mode OCR dipaksa aktif. Cek ulang semua kolom secara manual.",
                    }
            else:
                result = auto_extract(pdf_bytes)

        st.session_state.raw_df = result["df"]
        st.session_state.extract_method = result["method"]

        if result["warning"]:
            st.warning(result["warning"])

        if result["df"].empty:
            st.error(
                "Tidak ada data yang berhasil diekstrak. Kemungkinan PDF terenkripsi, "
                "kosong, atau formatnya tidak dikenali sama sekali."
            )
        else:
            st.success(
                f"Berhasil mengekstrak {len(result['df'])} baris menggunakan metode: "
                f"**{result['method']}**. Lanjutkan ke langkah 2 untuk verifikasi."
            )

# ---------------------------------------------------------------------------
# Step 2: Preview & mapping kolom (WAJIB dikonfirmasi manual)
# ---------------------------------------------------------------------------
raw_df = st.session_state.raw_df

if not raw_df.empty:
    st.header("2️⃣ Verifikasi Data Mentah & Petakan Kolom")
    st.info(
        "⚠️ Hasil ekstraksi otomatis BELUM tentu akurat, terutama kalau PDF hasil scan "
        "atau format tabel tidak standar. Periksa tabel di bawah, lalu tentukan kolom "
        "mana yang berisi Kode Barang, Nama Barang, dan NUP."
    )

    with st.expander("Lihat data mentah hasil ekstraksi", expanded=True):
        st.dataframe(raw_df, use_container_width=True, height=300)

    guess = guess_column_mapping(raw_df)
    options = ["(tidak ada)"] + list(raw_df.columns)

    def _idx(col_name):
        return options.index(col_name) if col_name in options else 0

    st.subheader("Pemetaan kolom")
    m1, m2, m3 = st.columns(3)
    with m1:
        col_kode = st.selectbox("Kolom = Kode Barang", options, index=_idx(guess["Kode Barang"]))
        col_nama = st.selectbox("Kolom = Nama Barang", options, index=_idx(guess["Nama Barang"]))
    with m2:
        col_nup = st.selectbox("Kolom = NUP", options, index=_idx(guess["NUP"]))
        col_ruangan = st.selectbox("Kolom = Ruangan (opsional)", options, index=_idx(guess["Ruangan"]))
    with m3:
        header_is_row0 = st.checkbox(
            "Baris pertama tabel adalah HEADER (bukan data barang)", value=True
        )
        st.caption(
            "Centang ini kalau baris pertama berisi judul kolom seperti "
            "'Kode Barang', 'Nama Barang', dst -- bukan data barang sungguhan."
        )

    if col_kode == "(tidak ada)" or col_nama == "(tidak ada)" or col_nup == "(tidak ada)":
        st.error("Kolom Kode Barang, Nama Barang, dan NUP wajib dipetakan sebelum lanjut.")
        st.stop()

    clean_df = raw_df.copy()
    if header_is_row0:
        clean_df = clean_df.iloc[1:].reset_index(drop=True)

    clean_df = clean_df.rename(columns={
        col_kode: "Kode Barang",
        col_nama: "Nama Barang",
        col_nup: "NUP",
        **({col_ruangan: "Ruangan"} if col_ruangan != "(tidak ada)" else {}),
    })

    # buang baris yang Kode Barang & NUP-nya kosong (baris sampah/footer)
    before_n = len(clean_df)
    clean_df = clean_df[
        (clean_df["Kode Barang"].astype(str).str.strip() != "") &
        (clean_df["NUP"].astype(str).str.strip() != "")
    ].reset_index(drop=True)
    dropped_n = before_n - len(clean_df)
    if dropped_n > 0:
        st.caption(f"({dropped_n} baris dibuang karena Kode Barang/NUP kosong -- kemungkinan baris header/footer/kosong)")

    if clean_df.empty:
        st.error(
            "Setelah pembersihan, tidak ada baris data valid yang tersisa. "
            "Cek kembali pemetaan kolom di atas."
        )
        st.stop()

    st.success(f"Data siap: {len(clean_df)} baris barang valid.")
    st.dataframe(clean_df, use_container_width=True, height=250)

    # -----------------------------------------------------------------------
    # Step 3: Rekap per jenis barang (Kode Barang -> daftar NUP)
    # -----------------------------------------------------------------------
    st.header("3️⃣ Rekap per Jenis Barang (Kode Barang)")

    recap = (
        clean_df.groupby(["Kode Barang", "Nama Barang"])
        .agg(
            Jumlah_Unit=("NUP", "count"),
            Daftar_NUP=("NUP", lambda x: ", ".join(sorted(set(map(str, x))))),
        )
        .reset_index()
        .sort_values(["Kode Barang"])
    )

    # cek duplikat NUP dalam kode barang yang sama (indikasi data error / nomor dobel)
    dup_check = clean_df.groupby(["Kode Barang", "NUP"]).size().reset_index(name="count")
    dup_issues = dup_check[dup_check["count"] > 1]
    if not dup_issues.empty:
        st.warning(
            f"⚠️ Ditemukan {len(dup_issues)} kombinasi Kode Barang + NUP yang duplikat "
            "(NUP seharusnya unik per jenis barang dalam satu ruangan). Periksa data sumber."
        )
        with st.expander("Lihat detail duplikat NUP"):
            st.dataframe(dup_issues, use_container_width=True)

    st.dataframe(recap, use_container_width=True, height=300)

    rekap_csv = recap.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Unduh Rekap (CSV)", rekap_csv, file_name="rekap_per_jenis_barang.csv", mime="text/csv"
    )

    # -----------------------------------------------------------------------
    # Step 4: Cetak label (grid 12 per lembar)
    # -----------------------------------------------------------------------
    st.header("4️⃣ Cetak Label Barang (12 label per lembar A4)")

    c1, c2, c3 = st.columns(3)
    with c1:
        instansi = st.text_input("Nama Instansi/Ruangan (tampil di setiap label)", value="")
    with c2:
        grid_choice = st.radio(
            "Layout grid per lembar",
            ["3 kolom x 4 baris (12)", "4 kolom x 3 baris (12)"],
            horizontal=False,
        )
        if grid_choice.startswith("3"):
            grid_cols, grid_rows = 3, 4
        else:
            grid_cols, grid_rows = 4, 3
    with c3:
        include_ruangan_label = (
            st.checkbox("Tampilkan kolom Ruangan di label", value=("Ruangan" in clean_df.columns))
            if "Ruangan" in clean_df.columns else False
        )

    st.markdown("**Urutan cetak:** diurutkan otomatis berdasarkan *Kode Barang*, lalu *NUP* — "
                "supaya label barang sejenis tercetak berurutan dan mudah dipilah saat ditempel.")

    print_df = clean_df.sort_values(["Kode Barang", "NUP"]).reset_index(drop=True)

    n_pages = -(-len(print_df) // (grid_cols * grid_rows))  # ceiling division
    st.caption(f"Total {len(print_df)} label → akan tercetak dalam {n_pages} halaman A4.")

    if st.button("🖨️ Generate PDF Label", type="primary"):
        pdf_bytes_out = generate_label_pdf(
            print_df,
            col_kode="Kode Barang",
            col_nama="Nama Barang",
            col_nup="NUP",
            col_ruangan="Ruangan" if include_ruangan_label and "Ruangan" in print_df.columns else None,
            instansi=instansi,
            cols=grid_cols,
            rows=grid_rows,
        )
        st.success("PDF label berhasil dibuat.")
        st.download_button(
            "⬇️ Unduh PDF Label (siap print)",
            pdf_bytes_out,
            file_name="label_barang_ruangan.pdf",
            mime="application/pdf",
        )

else:
    st.info("Upload PDF dan klik 'Ekstrak Data dari PDF' untuk mulai.")
