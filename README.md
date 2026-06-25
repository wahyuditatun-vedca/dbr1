# Sistem Penomoran & Pengelompokan Label Daftar Barang Ruangan (DBR)

Aplikasi Streamlit dengan 2 mode kerja.

## Mode B — Gabung & Kelompokkan PDF Label (mode utama, sesuai data nyata Anda)

Untuk PDF label yang **SUDAH** dalam format resmi (header Kementerian, logo
Garuda, Kode Barang, NUP, Nama Barang/Merk, QR code) — seperti contoh file
yang Anda kirim. Sistem ini:

1. Memotong setiap kotak label sebagai **konten PDF vektor utuh** (pakai
   PyMuPDF) — bukan menggambar ulang. Logo, QR, font, semuanya copy persis
   dari file asli.
2. Mengelompokkan ulang berdasarkan **Kode Barang lalu NUP**.
3. Menyusun ulang ke lembar baru, **12 label per lembar A4** (grid dikalibrasi
   otomatis dari halaman input yang punya 12 label penuh; kalau tidak ada,
   pakai grid acuan standar 2 kolom x 6 baris).
4. **Konten setiap label tidak diubah sama sekali** — hanya urutan/posisi
   penempatannya di lembar baru yang berubah.

Bisa upload beberapa file PDF sekaligus (misal label yang sudah dicetak
sebagian-sebagian dari rentang NUP berbeda) — semua digabung jadi satu PDF
yang rapi dan padat (tidak ada lembar yang cuma berisi 1-2 label seperti di
file asal).

Fitur tambahan:
- Deteksi **duplikat** (kombinasi Kode Barang + NUP yang muncul >1x antar file
  yang diupload — indikasi file tumpang tindih).
- Deteksi **gap NUP** (nomor yang hilang di tengah rentang per Kode Barang —
  info, bukan error, mungkin labelnya belum diupload).
- Rekap per jenis barang (CSV).
- **Manifest/audit trail** (CSV): mencatat tiap label pindah dari file+halaman
  asal ke halaman+urutan baru — penting untuk verifikasi di lingkungan kerja
  pemerintahan.

## Mode A — Generate Label dari Data Mentah (cadangan)

Untuk kasus Anda **belum** punya label sama sekali, hanya tabel mentah
(No, Kode Barang, Nama Barang, NUP, dst, tanpa QR). Mode ini men-generate
label baru dari nol (font sederhana, tanpa QR) — beda dari Mode B yang
mengopi label yang sudah ada.

## Instalasi

```bash
pip install -r requirements.txt
```

### Dependency tambahan untuk OCR (HANYA dipakai di Mode A, untuk PDF hasil scan/foto)

- **Tesseract OCR**: `sudo apt-get install tesseract-ocr tesseract-ocr-ind`
- **Poppler**: `sudo apt-get install poppler-utils`

Mode B tidak membutuhkan OCR sama sekali (bekerja di level vektor PDF asli).

## Menjalankan

```bash
streamlit run app.py
```

## Keterbatasan yang Perlu Disadari

- **Deteksi kotak label di Mode B mengasumsikan pola teks** `"<KodeBarang> NUP: <angka>"`
  di dalam tiap kotak (sesuai template Kemendikdasmen/SIMAK-BMN yang Anda
  contohkan). Kalau institusi lain pakai format teks berbeda, regex di
  `label_merge.py` (`NUP_PATTERN`) perlu disesuaikan.
- **Kalibrasi grid otomatis** paling akurat kalau salah satu halaman input
  punya 12 label penuh. Kalau semua file yang diupload kebetulan halaman
  terakhir/parsial saja, sistem jatuh ke grid acuan standar (margin hasil
  pengukuran dari contoh file Anda) — kemungkinan kecil bergeser kalau
  template institusi lain berbeda margin.
- **Sudah diuji penuh** menggunakan 2 file PDF label asli yang Anda kirim
  (23 label, 2 Kode Barang berbeda, lintas file) — termasuk verifikasi visual
  hasil cetak (render ke gambar, dicek manual tidak ada konten yang berubah/
  bergeser/overlap). Belum diuji dengan: PDF berisi Merk/Tipe yang sangat
  panjang (kemungkinan terpotong di tampilan kotak, tapi karena ini hasil
  COPY bukan re-render, tampilannya akan tetap sama seperti di PDF asli),
  atau jumlah file input yang sangat banyak (puluhan file sekaligus —
  performanya belum diukur).
