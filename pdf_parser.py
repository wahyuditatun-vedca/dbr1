"""
pdf_parser.py
Ekstraksi data dari PDF Daftar Barang Ruangan (DBR) SIMAK-BMN.

Strategi (berurutan, berhenti di langkah pertama yang berhasil):
1. pdfplumber.extract_tables()  -> kalau PDF hasil export langsung dan ada garis tabel.
2. pdfplumber.extract_text() lalu di-split berdasarkan spasi/tab ganda
   -> kalau PDF teks tapi tanpa garis tabel (rata kolom dengan spasi).
3. OCR (pytesseract + pdf2image) lalu di-split sama seperti #2
   -> kalau PDF hasil scan/foto (tidak ada teks sama sekali).

CATATAN PENTING (keterbatasan yang harus disadari, bukan disembunyikan):
- Heuristik split-by-whitespace pada langkah 2 & 3 RAPUH kalau nama barang
  mengandung spasi ganda yang tidak konsisten, atau kolom kosong di tengah baris.
  Karena itu hasil ekstraksi SELALU ditampilkan mentah ke pengguna untuk
  dikonfirmasi/dipetakan manual sebelum dipakai -- jangan percaya hasil
  parsing otomatis 100% tanpa verifikasi visual.
- OCR tidak pernah 100% akurat, terutama untuk angka (NUP, Kode Barang).
  Setelah OCR, pengguna WAJIB mengecek ulang kolom NUP & Kode Barang.
"""

import io
import re
from typing import List, Tuple

import pandas as pd
import pdfplumber


MIN_CHARS_FOR_TEXT_PDF = 30  # ambang batas: di bawah ini dianggap "tidak ada teks" -> coba OCR


def _pad_rows(rows: List[List[str]]) -> List[List[str]]:
    """Samakan jumlah kolom semua baris (isi kosong jika kurang)."""
    max_cols = max((len(r) for r in rows), default=0)
    return [r + [""] * (max_cols - len(r)) for r in rows]


def extract_via_tables(pdf_bytes: bytes) -> pd.DataFrame:
    """Ekstraksi via deteksi garis tabel pdfplumber. Return DataFrame kosong jika gagal."""
    all_rows: List[List[str]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    cleaned = [(c or "").strip() for c in row]
                    if any(cleaned):
                        all_rows.append(cleaned)

    if not all_rows or max(len(r) for r in all_rows) <= 1:
        return pd.DataFrame()

    all_rows = _pad_rows(all_rows)
    n_cols = len(all_rows[0])
    columns = [f"Kolom_{i+1}" for i in range(n_cols)]
    return pd.DataFrame(all_rows, columns=columns)


def extract_via_text_split(raw_text: str) -> pd.DataFrame:
    """
    Fallback: pisahkan setiap baris teks jadi kolom berdasarkan 2+ spasi atau tab.
    Dipakai untuk PDF teks tanpa garis tabel, ATAU hasil OCR.
    """
    rows: List[List[str]] = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        cells = [c.strip() for c in re.split(r"\t+|\s{2,}", line) if c.strip()]
        if len(cells) >= 2:
            rows.append(cells)

    if not rows:
        return pd.DataFrame()

    rows = _pad_rows(rows)
    n_cols = len(rows[0])
    columns = [f"Kolom_{i+1}" for i in range(n_cols)]
    return pd.DataFrame(rows, columns=columns)


def get_pdf_text(pdf_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            text_parts.append(t)
    return "\n".join(text_parts)


def extract_via_ocr(pdf_bytes: bytes) -> Tuple[str, str]:
    """
    OCR fallback. Mengembalikan (raw_text, error_message).
    error_message kosong jika sukses.
    Membutuhkan binary sistem: tesseract-ocr dan poppler-utils
    (TIDAK cukup hanya `pip install`, harus terinstall di OS).
    """
    try:
        import pytesseract
        from pdf2image import convert_from_bytes
    except ImportError as e:
        return "", f"Library OCR tidak terpasang: {e}"

    try:
        images = convert_from_bytes(pdf_bytes, dpi=300)
    except Exception as e:
        return "", (
            "Gagal merender PDF ke gambar untuk OCR. Kemungkinan poppler-utils "
            f"belum terinstall di sistem ini. Detail error: {e}"
        )

    try:
        text_parts = []
        for img in images:
            text_parts.append(pytesseract.image_to_string(img, lang="ind+eng"))
        return "\n".join(text_parts), ""
    except Exception as e:
        return "", (
            "Gagal menjalankan OCR. Kemungkinan tesseract-ocr belum terinstall "
            f"di sistem ini. Detail error: {e}"
        )


def auto_extract(pdf_bytes: bytes) -> dict:
    """
    Jalankan strategi ekstraksi berurutan. Return dict berisi:
    - df: DataFrame mentah (kolom generik Kolom_1, Kolom_2, ...)
    - method: metode yang berhasil dipakai ("tabel", "teks", "ocr", atau "gagal")
    - warning: pesan peringatan untuk ditampilkan ke pengguna (boleh kosong)
    """
    # 1. Coba deteksi tabel bergaris
    df = extract_via_tables(pdf_bytes)
    if not df.empty and len(df) >= 1:
        return {"df": df, "method": "tabel", "warning": ""}

    # 2. Coba ekstraksi teks biasa
    raw_text = get_pdf_text(pdf_bytes)
    if len(raw_text.strip()) >= MIN_CHARS_FOR_TEXT_PDF:
        df = extract_via_text_split(raw_text)
        if not df.empty:
            return {"df": df, "method": "teks", "warning": ""}

    # 3. Fallback OCR (PDF kemungkinan hasil scan/foto)
    ocr_text, err = extract_via_ocr(pdf_bytes)
    if err:
        return {
            "df": pd.DataFrame(),
            "method": "gagal",
            "warning": (
                "Ekstraksi otomatis gagal total. PDF ini sepertinya hasil scan/foto "
                f"dan OCR tidak bisa dijalankan di server ini. Detail: {err}"
            ),
        }

    df = extract_via_text_split(ocr_text)
    if df.empty:
        return {
            "df": pd.DataFrame(),
            "method": "gagal",
            "warning": "OCR berjalan tapi tidak ada baris data yang berhasil dikenali. "
                       "Coba unggah PDF dengan resolusi lebih tinggi.",
        }
    return {
        "df": df,
        "method": "ocr",
        "warning": "PDF ini diproses dengan OCR (hasil scan/foto terdeteksi). "
                   "Akurasi OCR tidak 100% -- WAJIB cek ulang kolom Kode Barang dan NUP "
                   "secara manual sebelum mencetak label.",
    }


def guess_column_mapping(df: pd.DataFrame) -> dict:
    """
    Tebak kolom mana yang berisi Kode Barang / Nama Barang / NUP / Ruangan / Kondisi,
    dengan mencocokkan kata kunci pada baris pertama (asumsi baris pertama = header)
    atau pada isi kolom jika header tidak terdeteksi sebagai kata kunci.
    Hasil tebakan ini HANYA default awal -- pengguna tetap harus mengonfirmasi.
    """
    mapping = {"Kode Barang": None, "Nama Barang": None, "NUP": None,
               "Ruangan": None, "Kondisi": None}
    if df.empty:
        return mapping

    header_row = [str(c).lower() for c in df.iloc[0].tolist()]

    keyword_map = {
        "Kode Barang": ["kode barang", "kode brg", "kode"],
        "Nama Barang": ["nama barang", "nama brg", "nama"],
        "NUP": ["nup", "nomor urut pendaftaran"],
        "Ruangan": ["ruang", "lokasi"],
        "Kondisi": ["kondisi"],
    }

    for field, keywords in keyword_map.items():
        for i, cell in enumerate(header_row):
            if any(kw in cell for kw in keywords):
                mapping[field] = df.columns[i]
                break

    return mapping
