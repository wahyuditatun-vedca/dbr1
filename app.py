"""
app.py
Sistem Penomoran & Pengelompokan Label Daftar Barang Ruangan (DBR).

Mode B (default, sesuai contoh data nyata): gabungkan beberapa PDF label
yang SUDAH dalam format resmi (header Kementerian + Kode Barang + NUP +
Nama Barang/Merk + QR), kelompokkan ulang berdasarkan Kode Barang & NUP,
lalu cetak ulang 12 label/lembar -- KONTEN setiap label tidak diubah sama
sekali, hanya posisi/urutannya.

Mode A (cadangan): kalau Anda punya data mentah Daftar Barang Ruangan
(tabel No/Kode/Nama/NUP biasa, BELUM berbentuk label dengan QR), mode ini
men-generate label dari nol.
"""

import streamlit as st
import pandas as pd

from label_merge import detect_cells, choose_template, find_duplicates, find_nup_gaps, build_merged_pdf
from pdf_parser import auto_extract, guess_column_mapping
from label_pdf import generate_label_pdf

st.set_page_config(page_title="Penomoran Label Barang Ruangan", layout="wide")

st.title("📋 Sistem Penomoran & Pengelompokan Label Barang Ruangan")

mode = st.sidebar.radio(
    "Pilih mode kerja",
    [
        "B. Gabung & Kelompokkan PDF Label (sudah ada label/QR)",
        "A. Generate Label dari Daftar Barang Ruangan mentah (belum ada label)",
    ],
    index=0,
)

# =============================================================================
# MODE B — Gabung & kelompokkan label yang sudah jadi
# =============================================================================
if mode.startswith("B"):
    st.caption(
        "Upload 1 atau beberapa PDF label resmi (yang sudah ada logo, Kode Barang, "
        "NUP, dan QR). Sistem akan memotong tiap kotak label persis seperti aslinya, "
        "mengelompokkan berdasarkan Kode Barang lalu NUP, dan menyusunnya ulang "
        "12 label per lembar. **Konten tiap label tidak diubah sama sekali.**"
    )

    uploaded_files = st.file_uploader(
        "Upload PDF label (boleh lebih dari satu)", type=["pdf"], accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("🔍 Deteksi & Proses Label", type="primary"):
            source_pdfs = {}
            all_cells = []
            all_pages_rects = []
            all_unmatched = []
            per_file_summary = []

            with st.spinner("Membaca dan mendeteksi kotak label di setiap PDF..."):
                for uf in uploaded_files:
                    data = uf.read()
                    source_pdfs[uf.name] = data
                    res = detect_cells(data, uf.name)
                    all_cells.extend(res["cells"])
                    all_pages_rects.extend(res["pages_rects"])
                    all_unmatched.extend(res["unmatched_cells"])
                    per_file_summary.append({
                        "File": uf.name,
                        "Label terdeteksi": len(res["cells"]),
                        "Gagal dibaca": len(res["unmatched_cells"]),
                    })

            st.session_state["mode_b_source_pdfs"] = source_pdfs
            st.session_state["mode_b_all_cells"] = all_cells
            st.session_state["mode_b_all_pages_rects"] = all_pages_rects
            st.session_state["mode_b_all_unmatched"] = all_unmatched
            st.session_state["mode_b_summary"] = per_file_summary

    if st.session_state.get("mode_b_all_cells"):
        all_cells = st.session_state["mode_b_all_cells"]
        all_unmatched = st.session_state["mode_b_all_unmatched"]
        source_pdfs = st.session_state["mode_b_source_pdfs"]
        all_pages_rects = st.session_state["mode_b_all_pages_rects"]

        st.header("1️⃣ Ringkasan Deteksi")
        st.dataframe(pd.DataFrame(st.session_state["mode_b_summary"]), use_container_width=True)
        st.success(f"Total {len(all_cells)} label berhasil dibaca dari {len(source_pdfs)} file.")

        if all_unmatched:
            st.warning(
                f"⚠️ {len(all_unmatched)} kotak terdeteksi tapi GAGAL dibaca Kode Barang/NUP-nya "
                "(format teksnya beda dari pola 'KodeBarang NUP: angka'). Kotak ini TIDAK ikut "
                "diproses/dicetak ulang -- cek detail di bawah, mungkin perlu ditangani manual."
            )
            with st.expander("Lihat detail kotak yang gagal dibaca"):
                st.dataframe(pd.DataFrame(all_unmatched)[["source_name", "page_idx", "raw_text"]],
                             use_container_width=True)

        # ---------------------------------------------------------------
        # 2. Duplikat & gap NUP (pengecekan, bukan pengubahan data)
        # ---------------------------------------------------------------
        st.header("2️⃣ Pengecekan Data")
        dup = find_duplicates(all_cells)
        if dup:
            st.warning(f"⚠️ Ditemukan {len(dup)} kombinasi Kode Barang + NUP yang DUPLIKAT "
                       "(kemungkinan file yang diupload tumpang tindih).")
            st.dataframe(pd.DataFrame(dup), use_container_width=True)
        else:
            st.caption("✅ Tidak ada duplikat Kode Barang + NUP antar file yang diupload.")

        gaps = find_nup_gaps(all_cells)
        if gaps:
            with st.expander(f"ℹ️ {len(gaps)} kode barang punya nomor NUP yang belum lengkap (info, bukan error)"):
                st.caption(
                    "Ini hanya menandakan ada NUP di tengah rentang yang TIDAK ada di file yang "
                    "diupload -- mungkin labelnya belum diupload, bukan berarti datanya salah."
                )
                st.dataframe(pd.DataFrame(gaps), use_container_width=True)

        # ---------------------------------------------------------------
        # 3. Rekap per jenis barang
        # ---------------------------------------------------------------
        st.header("3️⃣ Rekap per Jenis Barang (Kode Barang)")
        df_cells = pd.DataFrame(all_cells)
        recap = (
            df_cells.groupby(["kode", "nama_barang"])
            .agg(Jumlah_Unit=("nup", "count"),
                 Daftar_NUP=("nup", lambda x: ", ".join(str(v) for v in sorted(x))))
            .reset_index()
            .rename(columns={"kode": "Kode Barang", "nama_barang": "Nama Barang"})
            .sort_values("Kode Barang")
        )
        st.dataframe(recap, use_container_width=True, height=250)
        st.download_button(
            "⬇️ Unduh Rekap (CSV)",
            recap.to_csv(index=False).encode("utf-8"),
            file_name="rekap_per_jenis_barang.csv",
            mime="text/csv",
        )

        # ---------------------------------------------------------------
        # 4. Generate PDF gabungan, dikelompokkan, 12/lembar
        # ---------------------------------------------------------------
        st.header("4️⃣ Generate PDF Label Gabungan (Dikelompokkan)")

        sort_choice = st.radio(
            "Urutan pengelompokan label",
            ["Kode Barang lalu NUP (disarankan)", "NUP saja (semua kode digabung urut NUP)"],
            horizontal=False,
        )

        if sort_choice.startswith("Kode"):
            cells_sorted = sorted(all_cells, key=lambda c: (c["kode"], c["nup"]))
        else:
            cells_sorted = sorted(all_cells, key=lambda c: c["nup"])

        template, used_fallback = choose_template(all_pages_rects)
        if used_fallback:
            st.info(
                "ℹ️ Tidak ada satupun halaman input dengan 12 label penuh untuk kalibrasi grid "
                "otomatis -- sistem memakai grid acuan standar (2 kolom x 6 baris, A4). Kalau "
                "hasil cetak sedikit bergeser, upload minimal 1 halaman label yang penuh (12 label)."
            )

        n_per_page = len(template)
        n_pages = -(-len(cells_sorted) // n_per_page)
        st.caption(f"Total {len(cells_sorted)} label → {n_pages} halaman A4 ({n_per_page} label/lembar).")

        if st.button("🖨️ Generate PDF Gabungan", type="primary"):
            with st.spinner("Menyusun ulang label (copy konten asli, tanpa diubah)..."):
                merged_pdf = build_merged_pdf(cells_sorted, template, source_pdfs)

            manifest = pd.DataFrame([{
                "Urutan_Baru": i + 1,
                "Halaman_Baru": (i // n_per_page) + 1,
                "Kode Barang": c["kode"],
                "NUP": c["nup"],
                "Nama Barang": c["nama_barang"],
                "File_Asal": c["source_name"],
                "Halaman_Asal": c["page_idx"] + 1,
            } for i, c in enumerate(cells_sorted)])

            st.success("PDF gabungan berhasil dibuat.")
            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇️ Unduh PDF Label Gabungan (siap print)",
                    merged_pdf,
                    file_name="label_gabungan_terkelompok.pdf",
                    mime="application/pdf",
                )
            with d2:
                st.download_button(
                    "⬇️ Unduh Manifest/Audit Trail (CSV)",
                    manifest.to_csv(index=False).encode("utf-8"),
                    file_name="manifest_label.csv",
                    mime="text/csv",
                )
    else:
        st.info("Upload PDF label lalu klik 'Deteksi & Proses Label' untuk mulai.")

# =============================================================================
# MODE A — Generate label dari data DBR mentah (belum berbentuk label/QR)
# =============================================================================
else:
    st.caption(
        "Gunakan mode ini HANYA kalau data Anda masih berupa tabel mentah "
        "(No, Kode Barang, Nama Barang, NUP, dst) dan BELUM ada label dengan QR. "
        "Mode ini men-generate label baru dari nol (bukan meng-copy label yang sudah ada)."
    )

    if "raw_df" not in st.session_state:
        st.session_state.raw_df = pd.DataFrame()

    st.header("1️⃣ Upload PDF Daftar Barang Ruangan (tabel mentah)")
    uploaded = st.file_uploader("Pilih file PDF DBR", type=["pdf"], key="mode_a_upload")
    force_ocr = st.checkbox("Paksa pakai OCR (kalau PDF hasil scan/foto)", value=False)

    if uploaded is not None and st.button("🔍 Ekstrak Data dari PDF"):
        pdf_bytes = uploaded.read()
        with st.spinner("Mengekstrak data dari PDF..."):
            if force_ocr:
                from pdf_parser import extract_via_ocr, extract_via_text_split
                ocr_text, err = extract_via_ocr(pdf_bytes)
                if err:
                    st.error(err)
                    result = {"df": pd.DataFrame(), "method": "gagal", "warning": err}
                else:
                    df_ocr = extract_via_text_split(ocr_text)
                    result = {"df": df_ocr, "method": "ocr (manual)",
                              "warning": "Mode OCR dipaksa aktif. Cek ulang semua kolom secara manual."}
            else:
                result = auto_extract(pdf_bytes)

        st.session_state.raw_df = result["df"]
        if result["warning"]:
            st.warning(result["warning"])
        if result["df"].empty:
            st.error("Tidak ada data yang berhasil diekstrak.")
        else:
            st.success(f"Berhasil mengekstrak {len(result['df'])} baris ({result['method']}).")

    raw_df = st.session_state.raw_df
    if not raw_df.empty:
        st.header("2️⃣ Verifikasi & Petakan Kolom")
        with st.expander("Lihat data mentah hasil ekstraksi", expanded=True):
            st.dataframe(raw_df, use_container_width=True, height=250)

        guess = guess_column_mapping(raw_df)
        options = ["(tidak ada)"] + list(raw_df.columns)

        def _idx(col_name):
            return options.index(col_name) if col_name in options else 0

        m1, m2, m3 = st.columns(3)
        with m1:
            col_kode = st.selectbox("Kolom = Kode Barang", options, index=_idx(guess["Kode Barang"]))
            col_nama = st.selectbox("Kolom = Nama Barang", options, index=_idx(guess["Nama Barang"]))
        with m2:
            col_nup = st.selectbox("Kolom = NUP", options, index=_idx(guess["NUP"]))
            col_ruangan = st.selectbox("Kolom = Ruangan (opsional)", options, index=_idx(guess["Ruangan"]))
        with m3:
            header_is_row0 = st.checkbox("Baris pertama adalah HEADER", value=True)

        if col_kode != "(tidak ada)" and col_nama != "(tidak ada)" and col_nup != "(tidak ada)":
            clean_df = raw_df.copy()
            if header_is_row0:
                clean_df = clean_df.iloc[1:].reset_index(drop=True)
            clean_df = clean_df.rename(columns={
                col_kode: "Kode Barang", col_nama: "Nama Barang", col_nup: "NUP",
                **({col_ruangan: "Ruangan"} if col_ruangan != "(tidak ada)" else {}),
            })
            clean_df = clean_df[
                (clean_df["Kode Barang"].astype(str).str.strip() != "") &
                (clean_df["NUP"].astype(str).str.strip() != "")
            ].reset_index(drop=True)

            st.success(f"Data siap: {len(clean_df)} baris.")
            st.dataframe(clean_df, use_container_width=True, height=200)

            st.header("3️⃣ Cetak Label (12/lembar)")
            instansi = st.text_input("Nama Instansi (tampil di label)", value="")
            print_df = clean_df.sort_values(["Kode Barang", "NUP"]).reset_index(drop=True)

            if st.button("🖨️ Generate PDF Label", type="primary"):
                pdf_out = generate_label_pdf(
                    print_df, col_kode="Kode Barang", col_nama="Nama Barang", col_nup="NUP",
                    col_ruangan="Ruangan" if "Ruangan" in print_df.columns else None,
                    instansi=instansi, cols=3, rows=4,
                )
                st.download_button("⬇️ Unduh PDF Label", pdf_out,
                                    file_name="label_baru.pdf", mime="application/pdf")
        else:
            st.error("Petakan kolom Kode Barang, Nama Barang, dan NUP terlebih dahulu.")
    else:
        st.info("Upload PDF DBR mentah lalu klik 'Ekstrak Data dari PDF' untuk mulai.")
