"""
label_merge.py
Menggabungkan & mengelompokkan ulang PDF label barang yang SUDAH dalam format
resmi (header Kementerian + Kode Barang + NUP + Nama Barang + Merk/Tipe + QR),
TANPA mengubah konten tiap label sama sekali.

Teknik: setiap kotak label di-crop sebagai konten PDF vektor utuh (teks, logo,
QR) menggunakan PyMuPDF, lalu ditempel ulang ke halaman baru pada grid yang
sama (2 kolom x 6 baris = 12/lembar). Yang berubah HANYA urutan/posisi mana
yang dipasangkan jadi satu halaman -- bukan isi label itu sendiri.

Kalau ada perubahan pada teks, font, atau gambar di label, berarti ada bug --
tujuan modul ini adalah copy-paste vektor 1:1.
"""

import io
import re
from collections import defaultdict
from typing import List, Dict, Tuple

import pdfplumber
import fitz  # PyMuPDF


PAGE_W, PAGE_H = 595.28, 841.89  # A4 dalam pt

# Grid fallback (HANYA dipakai jika tidak ada satupun halaman input yang berisi
# 12 sel lengkap untuk dijadikan kalibrasi). Nilai ini diambil dari template
# resmi yang sudah diverifikasi cocok (BBPPMPV Pertanian, Kemendikdasmen).
# Kalau institusi/format lain punya margin berbeda, hasil dengan fallback ini
# BISA SEDIKIT BERGESER -- selalu lebih baik kalau ada minimal 1 halaman penuh
# (12 label) di salah satu file yang diupload untuk kalibrasi otomatis.
FALLBACK_TEMPLATE = [
    (5.7, 28.4, 294.8, 141.7), (300.5, 28.4, 589.6, 141.7),
    (5.7, 147.4, 294.8, 260.8), (300.5, 147.4, 589.6, 260.8),
    (5.7, 266.5, 294.8, 379.9), (300.5, 266.5, 589.6, 379.9),
    (5.7, 385.5, 294.8, 498.9), (300.5, 385.5, 589.6, 498.9),
    (5.7, 504.6, 294.8, 618.0), (300.5, 504.6, 589.6, 618.0),
    (5.7, 623.6, 294.8, 737.0), (300.5, 623.6, 589.6, 737.0),
]

NUP_PATTERN = re.compile(r"(\d+)\s*NUP\s*:\s*(\d+)")


def detect_cells(pdf_bytes: bytes, source_name: str) -> Dict:
    """
    Pindai satu file PDF label, kembalikan:
    - cells: list of dict {kode, nup, nama_barang (untuk info saja), source_name,
             page_idx, rect}
    - pages_rects: list per halaman, masing-masing list rect (untuk kalibrasi grid)
    - unmatched_cells: kotak yang terdeteksi tapi GAGAL dibaca Kode+NUP-nya
                        (format beda/rusak) -- harus ditampilkan ke pengguna,
                        jangan didiamkan/dianggap tidak ada.
    """
    cells = []
    pages_rects = []
    unmatched_cells = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            page_rect_list = []
            # urutkan rect berdasarkan posisi (atas->bawah, kiri->kanan) supaya
            # konsisten kalau nanti dipakai sebagai kandidat template
            rects_sorted = sorted(page.rects, key=lambda r: (round(r["top"]), r["x0"]))
            for r in rects_sorted:
                rect_tuple = (r["x0"], r["top"], r["x1"], r["bottom"])
                page_rect_list.append(rect_tuple)

                crop = page.crop(rect_tuple)
                text = (crop.extract_text() or "").replace("\n", " ")
                m = NUP_PATTERN.search(text)
                if not m:
                    unmatched_cells.append({
                        "source_name": source_name,
                        "page_idx": page_idx,
                        "rect": rect_tuple,
                        "raw_text": text[:120],
                    })
                    continue

                # nama barang: baris setelah baris "<kode> NUP: <nup>" pada teks asli
                nama_barang = ""
                lines = (crop.extract_text() or "").split("\n")
                for i, line in enumerate(lines):
                    if "NUP" in line and m.group(2) in line:
                        if i + 1 < len(lines):
                            nama_barang = lines[i + 1].strip()
                        break

                cells.append({
                    "kode": m.group(1),
                    "nup": int(m.group(2)),
                    "nama_barang": nama_barang,
                    "source_name": source_name,
                    "page_idx": page_idx,
                    "rect": rect_tuple,
                })
            pages_rects.append(page_rect_list)

    return {"cells": cells, "pages_rects": pages_rects, "unmatched_cells": unmatched_cells}


def choose_template(all_pages_rects: List[List[Tuple]]) -> Tuple[List[Tuple], bool]:
    """
    Pilih grid template untuk halaman output:
    - Prioritas: halaman INPUT manapun yang punya tepat 12 sel (1 lembar penuh)
      -> dipakai sebagai kalibrasi (paling akurat, sesuai file asli pengguna).
    - Fallback: kalau tidak ada satupun halaman 12-sel, pakai FALLBACK_TEMPLATE.

    Return: (template_positions, used_fallback: bool)
    """
    full_pages = [pr for pr in all_pages_rects if len(pr) == 12]
    if full_pages:
        return full_pages[0], False
    # ambil halaman dengan sel terbanyak sebagai upaya terbaik kedua
    best = max(all_pages_rects, key=len, default=[])
    if len(best) >= 12:
        return best[:12], False
    return FALLBACK_TEMPLATE, True


def find_duplicates(cells: List[Dict]) -> List[Dict]:
    """Cari kombinasi (kode, nup) yang muncul lebih dari sekali (kemungkinan file overlap)."""
    seen = defaultdict(list)
    for c in cells:
        seen[(c["kode"], c["nup"])].append(c)
    return [{"kode": k, "nup": n, "jumlah": len(v),
              "sumber": [f"{x['source_name']} (hal. {x['page_idx']+1})" for x in v]}
             for (k, n), v in seen.items() if len(v) > 1]


def find_nup_gaps(cells: List[Dict]) -> List[Dict]:
    """
    Per Kode Barang, cek apakah ada lompatan nomor NUP (indikasi label yang
    belum tercetak / belum diupload, BUKAN error -- hanya info untuk dicek).
    """
    by_kode = defaultdict(set)
    for c in cells:
        by_kode[c["kode"]].add(c["nup"])

    gaps = []
    for kode, nups in by_kode.items():
        if len(nups) < 2:
            continue
        lo, hi = min(nups), max(nups)
        missing = sorted(set(range(lo, hi + 1)) - nups)
        if missing:
            gaps.append({"kode": kode, "rentang": f"{lo}-{hi}", "nup_hilang": missing})
    return gaps


def build_merged_pdf(
    cells_sorted: List[Dict],
    template_positions: List[Tuple],
    source_pdfs: Dict[str, bytes],
) -> bytes:
    """
    Tempel ulang setiap cell (sudah diurutkan sesuai keinginan pengguna) ke
    halaman baru, per `len(template_positions)` label per halaman.
    Konten setiap label di-copy persis (vektor) dari PDF sumbernya -- TIDAK
    digambar ulang.
    """
    src_docs = {name: fitz.open(stream=data, filetype="pdf") for name, data in source_pdfs.items()}

    out_doc = fitz.open()
    per_page = len(template_positions)

    for start in range(0, len(cells_sorted), per_page):
        chunk = cells_sorted[start:start + per_page]
        new_page = out_doc.new_page(width=PAGE_W, height=PAGE_H)
        for cell, target_rect in zip(chunk, template_positions):
            src_doc = src_docs[cell["source_name"]]
            clip = fitz.Rect(*cell["rect"])
            target = fitz.Rect(*target_rect)
            new_page.show_pdf_page(target, src_doc, cell["page_idx"], clip=clip)

    result = out_doc.tobytes()
    out_doc.close()
    for d in src_docs.values():
        d.close()
    return result
