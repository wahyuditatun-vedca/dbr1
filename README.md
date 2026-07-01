# Sistem Cetak Label Barang Ruangan

Aplikasi Streamlit untuk mem-filter dan mencetak label BMN per ruangan
berdasarkan Daftar Barang Ruangan (DBR) — BBPPMPV Pertanian Cianjur.

## Alur Kerja

1. **Langkah 1 — Master Label**: Unggah 1 atau lebih PDF label BMN, klik
   `JADIKAN DATA UTAMA`. Sistem mengekstrak katalog label (kode, NUP,
   tahun) dari seluruh halaman.
2. **Langkah 2 — DBR**: Unggah file Excel DBR. Setiap sheet dianggap
   satu ruangan. Sistem otomatis mendeteksi baris header. Sheet non-ruangan
   (mis. `Master Aset`) di-exclude by default. Klik `PROSES`.
3. **Langkah 3 — Hasil**: Preview per ruangan (jumlah label match, tidak
   ditemukan, daftar detail). Unduh PDF per ruangan, atau ZIP semua.

## Aturan Matching

- Match dilakukan berdasarkan `(Kode Barang, NUP)`. Tahun perolehan
  ditampilkan untuk konteks tetapi tidak dijadikan validator ketat
  karena DBR sering menuliskan tahun single-value untuk banyak NUP.
- Kolom NUP di DBR bisa berisi multi-value: `"9, 10"`, `"'23"`,
  `"100, 99"`, `"4, 1, 3"` — semua di-parse sebagai list angka.
- Kode barang di-normalisasi (buang `'`, `.0`, whitespace).

## Instalasi

```bash
pip install -r requirements.txt
streamlit run app.py
```

Buka `http://localhost:8501` di browser.

## Batasan yang Diketahui

- **PDF label harus vector-text**, bukan hasil scan/screenshot. Jika PDF
  adalah scan, teks tidak dapat diekstrak (perlu OCR — belum
  diimplementasi).
- **Grid label 2 kolom** diasumsikan (sesuai format label BMN standar).
  Layout lain (1 kolom, 3 kolom, dst.) memerlukan penyesuaian di
  `extract_labels_from_pdf`.
- **Header DBR di-auto-detect** berdasarkan keyword `Nomor Urut
  Pendaftaran` + `Kode Barang`. Kalau nama kolom di file DBR sangat
  berbeda, deteksi bisa gagal → tambahkan keyword di fungsi `find_col`
  atau lakukan preprocessing manual.
- **Output PDF**: menggunakan `Page.show_pdf_page()` (vector-preserving)
  sehingga QR code tetap tajam saat dicetak. Layout output: A4 portrait,
  2 kolom.

## File Structure

- `app.py` — aplikasi Streamlit utama (single-file)
- `requirements.txt` — dependencies
- `README.md` — file ini

## Dependencies

- `streamlit` — UI framework
- `PyMuPDF` (fitz) — parsing & manipulasi PDF (vector-preserving crop)
- `pandas` + `openpyxl` — baca file Excel/CSV DBR
