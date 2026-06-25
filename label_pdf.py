"""
label_pdf.py
Generate PDF label barang (untuk ditempel di barang fisik), grid N label per
lembar A4 (default 12 = 3 kolom x 4 baris).

Urutan label di-print SESUAI urutan baris pada DataFrame yang diberikan --
pengelompokan (sort by Kode Barang lalu NUP) dilakukan di app.py SEBELUM
DataFrame ini dikirim ke sini, supaya label barang yang sejenis tercetak
berurutan dan mudah dipilah saat ditempel.
"""

from io import BytesIO
from typing import List

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor


PAGE_W, PAGE_H = A4


def _wrap_text(text: str, max_chars: int) -> List[str]:
    """Word-wrap sederhana untuk teks di dalam kotak label (tanpa dependensi tambahan)."""
    words = str(text).split()
    lines, current = [], ""
    for w in words:
        trial = (current + " " + w).strip()
        if len(trial) <= max_chars:
            current = trial
        else:
            if current:
                lines.append(current)
            current = w
    if current:
        lines.append(current)
    return lines or [""]


def generate_label_pdf(
    df: pd.DataFrame,
    col_kode: str,
    col_nama: str,
    col_nup: str,
    col_ruangan: str = None,
    instansi: str = "",
    cols: int = 3,
    rows: int = 4,
    margin_mm: float = 8,
) -> bytes:
    """
    df            : DataFrame yang SUDAH bersih & SUDAH diurutkan sesuai keinginan cetak.
    col_kode/col_nama/col_nup/col_ruangan : nama kolom di df untuk tiap field.
    cols, rows    : jumlah kolom & baris grid per halaman (default 3x4 = 12).
    """
    per_page = cols * rows
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    margin = margin_mm * mm
    usable_w = PAGE_W - 2 * margin
    usable_h = PAGE_H - 2 * margin
    cell_w = usable_w / cols
    cell_h = usable_h / rows

    border_color = HexColor("#444444")

    records = df.to_dict("records")
    total = len(records)

    for start in range(0, total, per_page):
        chunk = records[start:start + per_page]

        for idx, rec in enumerate(chunk):
            col_i = idx % cols
            row_i = idx // cols

            x = margin + col_i * cell_w
            # y dihitung dari atas halaman ke bawah
            y_top = PAGE_H - margin - row_i * cell_h
            y_bottom = y_top - cell_h

            pad = 2.5 * mm

            # Kotak border label (garis putus-putus tipis sebagai panduan gunting)
            c.setStrokeColor(border_color)
            c.setLineWidth(0.4)
            c.setDash(2, 2)
            c.rect(x, y_bottom, cell_w, cell_h, stroke=1, fill=0)
            c.setDash()  # reset ke garis solid untuk teks

            text_x = x + pad
            text_y = y_top - pad - 3 * mm
            max_text_w = cell_w - 2 * pad

            # Nama instansi (opsional, font kecil di atas)
            if instansi:
                c.setFont("Helvetica", 6.5)
                c.setFillColor(HexColor("#000000"))
                c.drawString(text_x, text_y, instansi[:40])
                text_y -= 3.4 * mm

            # Kode Barang (tebal, paling menonjol)
            kode_val = str(rec.get(col_kode, "")).strip()
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y, f"Kode: {kode_val}")
            text_y -= 4 * mm

            # NUP
            nup_val = str(rec.get(col_nup, "")).strip()
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y, f"NUP : {nup_val}")
            text_y -= 4.2 * mm

            # Nama Barang (boleh wrap 2 baris)
            nama_val = str(rec.get(col_nama, "")).strip()
            approx_chars_per_line = max(8, int(max_text_w / (1.7 * mm)))
            wrapped = _wrap_text(nama_val, approx_chars_per_line)[:2]
            c.setFont("Helvetica", 7.5)
            for line in wrapped:
                if text_y < y_bottom + pad:
                    break
                c.drawString(text_x, text_y, line)
                text_y -= 3.2 * mm

            # Ruangan (opsional, baris kecil paling bawah jika muat)
            if col_ruangan and col_ruangan in rec:
                ruangan_val = str(rec.get(col_ruangan, "")).strip()
                if ruangan_val and text_y >= y_bottom + pad:
                    c.setFont("Helvetica-Oblique", 6.5)
                    c.drawString(text_x, text_y, f"Ruang: {ruangan_val[:30]}")

        c.showPage()

    c.save()
    buf.seek(0)
    return buf.read()
