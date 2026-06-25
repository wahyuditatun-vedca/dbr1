# Sistem Penomoran & Cetak Label Daftar Barang Ruangan (DBR)

Aplikasi Streamlit untuk: upload PDF DBR → ekstrak data → kelompokkan per
Kode Barang & NUP → cetak label fisik barang (12 label per lembar A4).

## Instalasi

```bash
pip install -r requirements.txt
```

### Dependency tambahan untuk OCR (WAJIB jika PDF Anda hasil scan/foto)

`pytesseract` dan `pdf2image` di atas hanya wrapper Python — keduanya
membutuhkan **binary sistem** yang harus diinstall terpisah di server/komputer
Anda:

- **Tesseract OCR**
  - Ubuntu/Debian: `sudo apt-get install tesseract-ocr tesseract-ocr-ind`
  - Windows: install dari https://github.com/UB-Mannheim/tesseract/wiki
- **Poppler** (untuk render PDF → gambar)
  - Ubuntu/Debian: `sudo apt-get install poppler-utils`
  - Windows: download poppler, tambahkan ke PATH

Tanpa dua binary ini, fitur OCR akan menampilkan pesan error yang jelas
(bukan crash diam-diam) — aplikasi tetap bisa dipakai untuk PDF teks biasa.

## Menjalankan

```bash
streamlit run app.py
```

## Alur Pakai

1. **Upload** PDF Daftar Barang Ruangan.
2. Klik **Ekstrak Data dari PDF**. Sistem mencoba 3 metode berurutan:
   tabel bergaris → teks biasa → OCR (otomatis fallback kalau tidak ada teks).
   Centang **"Paksa pakai OCR"** kalau Anda sudah tahu PDF-nya hasil scan/foto.
3. **Verifikasi & petakan kolom** — sistem menebak kolom mana yang berisi
   Kode Barang/Nama Barang/NUP, tapi **WAJIB dicek manual**. Format DBR
   pemerintah tidak seragam, jadi tebakan otomatis bisa salah.
4. **Rekap per jenis barang** otomatis muncul (Kode Barang → daftar NUP,
   jumlah unit), termasuk peringatan kalau ada NUP duplikat dalam kode
   barang yang sama (indikasi data sumber bermasalah).
5. **Cetak label**: pilih layout grid (3x4 atau 4x3 = 12/lembar), isi nama
   instansi/ruangan (opsional), lalu generate & unduh PDF siap print.
   Urutan cetak otomatis diurutkan berdasarkan Kode Barang lalu NUP, supaya
   label barang sejenis tercetak berurutan dan mudah dipilah saat ditempel.

## Keterbatasan yang Perlu Disadari (bukan diasumsikan beres)

- **Heuristik parsing teks/OCR rapuh.** Untuk PDF tanpa garis tabel, kolom
  dideteksi dari spasi/tab ganda — ini bisa salah kalau nama barang punya
  spasi tidak konsisten. Karena itu langkah verifikasi manual (langkah 3)
  tidak bisa dilewati/diotomasi penuh.
- **OCR tidak 100% akurat**, terutama untuk digit (NUP, Kode Barang yang
  berupa angka panjang). Setelah OCR, cek ulang dua kolom ini secara teliti
  sebelum cetak — kesalahan satu digit pada NUP berarti label salah tempel.
- **Belum diuji dengan PDF DBR asli** — sistem ini sudah diuji penuh
  (ekstraksi tabel → mapping kolom → rekap → generate PDF label, termasuk
  verifikasi visual hasil cetak) menggunakan PDF dummy yang dibuat sendiri
  untuk simulasi, karena belum ada sampel PDF asli yang diterima. Struktur
  DBR asli BBPPMPV Anda mungkin punya kolom tambahan (Merk/Tipe, Asal
  Perolehan, dll) atau urutan kolom berbeda — App ini dirancang fleksibel
  (mapping kolom manual) justru karena ketidakpastian ini, tapi tetap
  perlu dicoba dengan file asli untuk memastikan ekstraksi tabel/teks
  bekerja sesuai layout PDF Anda yang sebenarnya.
