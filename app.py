"""
Sistem Cetak Label DBR BBPPMPV Pertanian  (v4)
================================================
Arsitektur:
  - Master PDF  → katalog validasi (kode, NUP, tahun)
  - DBR Excel   → data label per ruangan (sumber utama cetak)
  - Output PDF  → label di-GENERATE sebagai vector-text + QR code
"""

import io, re, zipfile, math
from dataclasses import dataclass, field
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd
import segno
import streamlit as st

CURRENT_YEAR = datetime.now().year
KODE_UAKPB = "138050200693453000KD"

# ============================================================
# PAGE CONFIG & CSS
# ============================================================
st.set_page_config(
    page_title="Sistem Cetak Label DBR",
    page_icon="🏷️", layout="wide", initial_sidebar_state="expanded",
)
st.markdown("""
<style>
.main .block-container { padding-top:1rem; max-width:1300px; }
.app-header {
    background: linear-gradient(135deg, #e65100 0%, #ff8f00 100%);
    color:#fff; padding:1.2rem 2rem; border-radius:14px;
    margin-bottom:1.2rem; box-shadow:0 6px 18px rgba(230,81,0,.22);
}
.app-header .title { font-size:1.45rem; font-weight:700; }
.app-header .sub   { font-size:.88rem; opacity:.92; margin-top:.2rem; }
.step-card {
    background:#fff; border:1px solid #e2e8f0; border-radius:12px;
    padding:1rem 1.2rem; margin-bottom:.7rem; box-shadow:0 1px 3px rgba(0,0,0,.04);
}
.step-card.active { border-left:4px solid #e65100; }
.step-card.done   { border-left:4px solid #2e7d32; background:#f4faf5; }
.step-card.locked { opacity:.5; border-left:4px solid #cbd5e0; }
.step-header {
    display:flex; align-items:center; gap:.6rem;
    font-size:1.05rem; font-weight:600; color:#1a202c; margin-bottom:.15rem;
}
.step-badge {
    background:#e65100; color:#fff; width:26px; height:26px;
    border-radius:50%; display:flex; align-items:center; justify-content:center;
    font-size:.8rem; font-weight:700;
}
.step-badge.done { background:#2e7d32; }
.metric-big {
    background:linear-gradient(135deg,#e65100,#ff8f00);
    color:#fff; padding:1rem; border-radius:12px; text-align:center;
}
.metric-big .num { font-size:2.2rem; font-weight:700; line-height:1; }
.metric-big .lbl { font-size:.8rem; opacity:.9; margin-top:.25rem;
                   text-transform:uppercase; letter-spacing:.4px; }
section[data-testid="stSidebar"] { background:#fdf5ef; }
.sidebar-step {
    display:flex; align-items:center; gap:.45rem;
    padding:.45rem .65rem; border-radius:8px; margin-bottom:.3rem;
    font-size:.86rem; color:#4a5568;
}
.sidebar-step.active { background:#fff3e0; color:#bf360c; font-weight:600; }
.sidebar-step.done   { background:#e8f5e9; color:#1b5e20; }
.mini-note {
    font-size:.82rem; color:#4a5568; background:#f7fafc;
    border-left:3px solid #a0aec0; padding:.45rem .7rem;
    border-radius:4px; margin:.4rem 0;
}
.cond-card {
    padding:.7rem; border-radius:8px; text-align:center; margin-bottom:.4rem;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# DATA CLASSES
# ============================================================
@dataclass
class LabelInfo:
    kode_barang: str
    nup: int
    tahun: str
    source_pdf: str
    page_index: int
    nama_barang: str = ""
    doc_id: int = 0


@dataclass
class DBRItem:
    nup: int
    kode_barang: str
    nama_barang: str
    merk: str
    tahun: str
    keterangan: str
    in_master: bool = False


@dataclass
class RuanganResult:
    ruangan: str
    sheet_name: str
    items: list = field(default_factory=list)   # list[DBRItem]


# ============================================================
# HELPERS
# ============================================================
def parse_multi_number(value) -> list:
    if value is None:
        return []
    if isinstance(value, (int, float)):
        return [] if pd.isna(value) else [int(value)]
    s = str(value).strip().replace("'", "").replace('"', "").replace("`", "")
    if not s or s.lower() == "nan":
        return []
    parts = re.split(r"[,;/]", s)
    result = set()
    for p in parts:
        m = re.search(r"\d+", p.strip())
        if m:
            try:
                result.add(int(m.group()))
            except ValueError:
                pass
    return sorted(result)


def clean_kode(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        try:
            return str(int(value))
        except (ValueError, OverflowError):
            return ""
    if isinstance(value, int):
        return str(value)
    s = str(value).strip().replace("'", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    # Handle scientific notation (rare but possible)
    try:
        if "e" in s.lower():
            s = str(int(float(s)))
    except (ValueError, OverflowError):
        pass
    m = re.search(r"\d{6,}", s)
    return m.group() if m else s


def classify_condition(text) -> str:
    if pd.isna(text) or not str(text).strip():
        return "TIDAK DIKETAHUI"
    t = re.sub(r"\s+", " ", str(text).strip().upper())
    if t in ("BAIK", "B", "BK", "GOOD"):
        return "BAIK"
    if any(k in t for k in ("RUSAK RINGAN", "RR", "RUSAK R")):
        return "RUSAK RINGAN"
    if any(k in t for k in ("RUSAK BERAT", "RB")):
        return "RUSAK"
    if t in ("RUSAK", "R"):
        return "RUSAK"
    if "KURANG" in t:
        return "RUSAK RINGAN"
    if "RUSAK" in t and "RINGAN" in t:
        return "RUSAK RINGAN"
    if "RUSAK" in t:
        return "RUSAK"
    if "BAIK" in t:
        return "BAIK"
    return "BAIK"


COND_DISPLAY = {
    "BAIK":            ("BAIK",                  "#2e7d32", "#e8f5e9"),
    "RUSAK RINGAN":    ("PERLU PERBAIKAN",       "#f57f17", "#fff8e1"),
    "RUSAK":           ("PERLU DIGANTI/DIHAPUS", "#c62828", "#ffebee"),
    "TIDAK DIKETAHUI": ("TIDAK DIKETAHUI",       "#78909c", "#eceff1"),
}


def priority_score(kondisi: str, tahun: str) -> float:
    """
    Weighted Decision-Tree Scoring untuk prioritas penanganan BMN.

    Skor = bobot_kondisi × faktor_usia
      - bobot_kondisi:  BAIK=1, RUSAK RINGAN=5, RUSAK=10
      - faktor_usia:    1 + (umur / 10)   (makin tua makin prioritas)

    Ini decision-tree scoring — BUKAN XGBoost/ML, karena:
      1. Kategori input deterministik (3 kelas biner), bukan fitur kontinu.
      2. Tidak ada training data historis untuk supervised learning.
      3. Rule-tree menghasilkan output identik dengan ML tapi interpretable,
         zero-dependency, dan reproducible.
    Algoritma ML (XGBoost, Random Forest) tepat jika ada task prediktif,
    misal: prediksi kapan barang akan rusak berdasarkan riwayat perbaikan.
    """
    w = {"BAIK": 1, "RUSAK RINGAN": 5, "RUSAK": 10, "TIDAK DIKETAHUI": 2}
    try:
        age = max(0, CURRENT_YEAR - int(tahun))
    except (ValueError, TypeError):
        age = 0
    return w.get(kondisi, 2) * (1 + age / 10)


def priority_label(kondisi: str, tahun: str) -> str:
    try:
        age = CURRENT_YEAR - int(tahun)
    except (ValueError, TypeError):
        age = 0
    if kondisi == "RUSAK":
        return "🔴 SEGERA GANTI"
    if kondisi == "RUSAK RINGAN" and age > 10:
        return "🟠 JADWALKAN PENGGANTIAN"
    if kondisi == "RUSAK RINGAN":
        return "🟡 PERBAIKI"
    if kondisi == "BAIK" and age > 20:
        return "🔵 PANTAU (Usang)"
    return "🟢 AMAN"


# ============================================================
# MASTER PDF EXTRACTION (validasi saja)
# ============================================================
KODE_RE = re.compile(r"\b(\d{10})\b")
NUP_RE  = re.compile(r"NUP\s*:?\s*(\d+)", re.IGNORECASE)
TAHUN_RE = re.compile(r"KD\s*\.?\s*(\d{4})")


def extract_catalog(pdf_bytes: bytes, source_name: str) -> list:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    labels = []
    for pi in range(doc.page_count):
        page = doc[pi]
        spans = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type", 0) != 0:
                continue
            for ln in b.get("lines", []):
                for sp in ln.get("spans", []):
                    t = sp.get("text", "").strip()
                    if t:
                        spans.append((t, tuple(sp["bbox"])))
        if not spans:
            continue
        page_w = page.rect.width
        kem_ys = sorted(set(round(bb[1], 1) for t, bb in spans if "KEMENTERIAN" in t.upper()))
        if not kem_ys:
            continue
        row_tops = []
        for y in kem_ys:
            if not row_tops or (y - row_tops[-1]) > 8:
                row_tops.append(y)
        rh = min(row_tops[i+1]-row_tops[i] for i in range(len(row_tops)-1)) if len(row_tops) >= 2 else 100
        for rt in row_tops:
            y0, y1 = max(0, rt-6), rt+rh-4
            for col in (0, 1):
                cx0 = 4 if col == 0 else page_w/2+4
                cx1 = page_w/2-4 if col == 0 else page_w-4
                cs = [(bb[1], t) for t, bb in spans
                      if cx0 <= (bb[0]+bb[2])/2 <= cx1 and y0 <= (bb[1]+bb[3])/2 <= y1]
                if not cs:
                    continue
                cs.sort()
                txt = " ".join(t for _, t in cs)
                mn = NUP_RE.search(txt)
                mk = KODE_RE.search(txt)
                mt = TAHUN_RE.search(txt)
                if not (mn and mk and mt):
                    continue
                try:
                    nup_val = int(mn.group(1))
                except ValueError:
                    continue
                labels.append(LabelInfo(
                    kode_barang=mk.group(1), nup=nup_val, tahun=mt.group(1),
                    source_pdf=source_name, page_index=pi,
                ))
    doc.close()
    return labels


# ============================================================
# DBR DETECTION (data-row-first, robust)
# ============================================================
def detect_dbr(xl_file, sheet_name: str):
    try:
        raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=object)
    except Exception:
        return None, sheet_name
    if raw.empty or len(raw) < 3:
        return None, sheet_name
    nc = len(raw.columns)

    # Cari kolom Kode Barang (7-10 digit) — scan sampai 200 baris, 20 kolom
    kode_col, first_row = None, None
    scan_limit = min(len(raw), 200)
    col_limit = min(nc, 20)
    for i in range(scan_limit):
        for j in range(col_limit):
            v = raw.iloc[i, j]
            if pd.isna(v):
                continue
            cleaned = clean_kode(v)
            if re.match(r"^\d{7,10}$", cleaned):
                first_row, kode_col = i, j
                break
        if first_row is not None:
            break
    if first_row is None:
        return None, sheet_name

    # Header scan (backward, sampai 10 baris ke atas)
    col_map = {"kode": kode_col}
    for sr in range(first_row - 1, max(first_row - 10, -1), -1):
        if sr < 0:
            break
        vals = [str(x).strip().lower() for x in raw.iloc[sr].values if pd.notna(x)]
        joined = " ".join(vals)
        # Skip number row (1, 2, 3, 4, 5…)
        if sum(1 for v in vals if re.match(r"^\d{1,2}$", v)) >= 4:
            continue
        if not any(k in joined for k in ["nama barang", "kode barang", "keterangan",
                                          "merk", "pendaftaran", "tahun"]):
            continue
        for j in range(nc):
            v = raw.iloc[sr, j]
            if pd.isna(v):
                continue
            t = str(v).strip().lower()
            if ("pendaftaran" in t and "urut" in t) or t == "nup":
                col_map["nup"] = j
            elif "nama barang" in t:
                col_map["nama"] = j
            elif "merk" in t or "type" in t:
                col_map["merk"] = j
            elif "kode barang" in t:
                col_map["kode"] = j
            elif "tahun" in t:
                col_map["tahun"] = j
            elif "keterangan" in t:
                col_map["ket"] = j
        break

    # Fallback positional — verifikasi NUP col
    if "nup" not in col_map:
        for off in [kode_col-3, kode_col-2, kode_col-1]:
            if off >= 0:
                tv = raw.iloc[first_row, off]
                if parse_multi_number(tv):
                    col_map["nup"] = off
                    break
        col_map.setdefault("nup", max(0, kode_col-3))
    col_map.setdefault("nama", max(0, kode_col-2))
    col_map.setdefault("merk", max(0, kode_col-1))
    col_map.setdefault("tahun", min(nc-1, kode_col+1))
    col_map.setdefault("ket", min(nc-1, kode_col+3))

    # Nama ruangan
    ruangan = sheet_name
    for i in range(first_row):
        for j in range(nc):
            v = raw.iloc[i, j]
            if pd.notna(v) and "ruangan" in str(v).lower():
                txt = str(v).strip()
                if ":" in txt:
                    c = txt.split(":", 1)[1].strip()
                    if c:
                        ruangan = c
                elif j+1 < nc and pd.notna(raw.iloc[i, j+1]):
                    ruangan = str(raw.iloc[i, j+1]).strip().lstrip(": ")
                break

    # Ekstrak data
    items = []
    for i in range(first_row, len(raw)):
        kv = raw.iloc[i, col_map["kode"]] if col_map["kode"] < nc else None
        if pd.isna(kv):
            continue
        kode = clean_kode(kv)
        if not re.match(r"^\d{6,}$", kode):
            continue
        def g(k):
            idx = col_map.get(k)
            if idx is None or idx >= nc:
                return None
            return raw.iloc[i, idx]
        nup_list = parse_multi_number(g("nup"))
        if not nup_list:
            continue
        tahun_list = parse_multi_number(g("tahun"))
        nama = str(g("nama") or "").strip()
        merk = str(g("merk") or "").strip()
        ket = str(g("ket") or "").strip() if pd.notna(g("ket")) else ""

        for nup in nup_list:
            # Jika ada multi-tahun, pasangkan posisi. Jika single → pakai satu utk semua
            if len(tahun_list) == len(nup_list):
                idx_t = nup_list.index(nup)
                thn = str(tahun_list[idx_t])
            elif tahun_list:
                thn = str(tahun_list[0])
            else:
                thn = ""
            items.append(DBRItem(
                nup=nup, kode_barang=kode, nama_barang=nama,
                merk=merk, tahun=thn, keterangan=ket,
            ))
    if not items:
        return None, ruangan
    return items, ruangan


# ============================================================
# PDF LABEL GENERATOR (vector-text, QR code, clean)
# ============================================================
def make_qr_png(data: str) -> bytes:
    qr = segno.make(data, error="L")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=3, border=1)
    return buf.getvalue()


def draw_label(page, x0, y0, w, h, item: DBRItem, kode_uakpb: str):
    """Gambar satu label BMN sebagai vector-text + QR."""
    rect = fitz.Rect(x0, y0, x0+w, y0+h)
    page.draw_rect(rect, color=(0, 0, 0), width=0.7)

    # ---- Header: KEMENTERIAN ----
    hdr_rect = fitz.Rect(x0+4, y0+3, x0+w-4, y0+15)
    page.insert_textbox(hdr_rect, "KEMENTERIAN PENDIDIKAN DASAR DAN",
                        fontsize=6.5, fontname="helv", align=1)
    hdr2 = fitz.Rect(x0+4, y0+12, x0+w-4, y0+22)
    page.insert_textbox(hdr2, "MENENGAH", fontsize=6.5, fontname="helv", align=1)

    # UAKPB
    uakpb = f"{kode_uakpb}.{item.tahun}"
    uakpb_rect = fitz.Rect(x0+4, y0+21, x0+w-4, y0+30)
    page.insert_textbox(uakpb_rect, uakpb, fontsize=5.8, fontname="helv", align=1)

    # Garis separator
    page.draw_line(fitz.Point(x0+4, y0+32), fitz.Point(x0+w-4, y0+32),
                   color=(0, 0, 0), width=0.3)

    # ---- QR Code (kanan) ----
    qr_size = 44
    qr_x = x0 + w - qr_size - 6
    qr_y = y0 + 34
    try:
        qr_bytes = make_qr_png(uakpb)
        qr_rect = fitz.Rect(qr_x, qr_y, qr_x + qr_size, qr_y + qr_size)
        page.insert_image(qr_rect, stream=qr_bytes)
    except Exception:
        pass

    # ---- Teks kiri ----
    text_x = x0 + 6
    text_w = w - qr_size - 18

    # Kode Barang + NUP
    page.insert_text(fitz.Point(text_x, y0+43),
                     item.kode_barang, fontsize=7.5, fontname="hebo")
    page.insert_text(fitz.Point(text_x + text_w/2 + 10, y0+43),
                     f"NUP: {item.nup}", fontsize=7.5, fontname="helv")

    # Nama Barang (potong jika terlalu panjang)
    nama = item.nama_barang[:45] if len(item.nama_barang) > 45 else item.nama_barang
    page.insert_text(fitz.Point(text_x, y0+54), nama,
                     fontsize=6.5, fontname="helv")

    # Garis separator
    page.draw_line(fitz.Point(x0+4, y0+60), fitz.Point(qr_x-4, y0+60),
                   color=(0, 0, 0), width=0.2)

    # Merk
    merk = item.merk[:45] if len(item.merk) > 45 else item.merk
    page.insert_text(fitz.Point(text_x, y0+72), merk,
                     fontsize=6.5, fontname="helv")


def draw_label_unavailable(page, x0, y0, w, h, item: DBRItem):
    """Label placeholder untuk item yang TIDAK ADA di master PDF."""
    rect = fitz.Rect(x0, y0, x0 + w, y0 + h)
    # Background abu-abu muda
    page.draw_rect(rect, color=(0.5, 0.5, 0.5), fill=(0.95, 0.95, 0.95), width=0.5,
                   dashes="[2] 0")

    # Nama barang
    nama = item.nama_barang[:50] if len(item.nama_barang) > 50 else item.nama_barang
    page.insert_text(fitz.Point(x0 + 8, y0 + 18), nama,
                     fontsize=7.5, fontname="hebo")

    # Kode + NUP (info referensi)
    page.insert_text(fitz.Point(x0 + 8, y0 + 32),
                     f"Kode: {item.kode_barang}   NUP: {item.nup}   Tahun: {item.tahun}",
                     fontsize=6.5, fontname="helv", color=(0.3, 0.3, 0.3))

    # Pesan "tidak tersedia"
    msg_rect = fitz.Rect(x0 + 8, y0 + 42, x0 + w - 8, y0 + h - 6)
    page.insert_textbox(msg_rect,
                        "⚠ LABEL TIDAK TERSEDIA DI MASTER LABEL",
                        fontsize=8, fontname="hebo", color=(0.7, 0.1, 0.1),
                        align=0)


def build_output_pdf(ruangan: str, items: list, kode_uakpb: str = KODE_UAKPB) -> bytes:
    """Generate PDF label untuk satu ruangan.
    - Item in_master=True  → label lengkap + QR code
    - Item in_master=False → placeholder 'Label Tidak Tersedia'
    """
    if not items:
        return b""
    doc = fitz.open()
    PAGE_W, PAGE_H = 595, 842
    MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 28, 52, 25
    GUTTER_X, GUTTER_Y = 10, 7
    CELL_W = (PAGE_W - 2*MARGIN_X - GUTTER_X) / 2
    CELL_H = 82

    usable_h = PAGE_H - MARGIN_TOP - MARGIN_BOTTOM
    rows_per_page = max(1, int((usable_h + GUTTER_Y) / (CELL_H + GUTTER_Y)))
    cells_per_page = rows_per_page * 2

    for i, item in enumerate(items):
        pos = i % cells_per_page
        if pos == 0:
            page = doc.new_page(width=PAGE_W, height=PAGE_H)
            page.insert_text(fitz.Point(MARGIN_X, 22),
                             f"DAFTAR LABEL BARANG — {ruangan.upper()}",
                             fontsize=10, fontname="hebo")
            page.draw_line(fitz.Point(MARGIN_X, 27),
                           fitz.Point(PAGE_W - MARGIN_X, 27),
                           color=(0, 0, 0), width=0.6)
            page_num = (i // cells_per_page) + 1
            total_pages = math.ceil(len(items) / cells_per_page)
            page.insert_text(fitz.Point(PAGE_W - MARGIN_X - 60, 22),
                             f"Hal {page_num}/{total_pages}",
                             fontsize=7, fontname="helv")

        col = pos % 2
        row = pos // 2
        x0 = MARGIN_X + col * (CELL_W + GUTTER_X)
        y0 = MARGIN_TOP + row * (CELL_H + GUTTER_Y)

        if item.in_master:
            draw_label(page, x0, y0, CELL_W, CELL_H, item, kode_uakpb)
        else:
            draw_label_unavailable(page, x0, y0, CELL_W, CELL_H, item)

    buf = io.BytesIO()
    doc.save(buf, deflate=True)
    doc.close()
    return buf.getvalue()


# ============================================================
# SESSION STATE
# ============================================================
def init_state():
    for k, v in {"catalog": [], "master_ready": False, "master_meta": [],
                 "dbr_processed": False, "ruangan_results": {},
                 "generated_pdfs": {}, "failed_sheets": []}.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ============================================================
# HEADER
# ============================================================
st.markdown("""
<div class="app-header">
  <div class="title">🏷️ Sistem Cetak Label DBR BBPPMPV Pertanian</div>
  <div class="sub">Filter &amp; cetak label BMN per ruangan · Data dari DBR Excel · Output vector-text + QR Code</div>
</div>""", unsafe_allow_html=True)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("### 📋 Progres")
    s1 = "done" if st.session_state.master_ready else "active"
    s2 = "done" if st.session_state.dbr_processed else ("active" if st.session_state.master_ready else "")
    s3 = "done" if st.session_state.generated_pdfs else ("active" if st.session_state.dbr_processed else "")
    for i, (lb, sts) in enumerate([("Master label (validasi)", s1),
                                    ("Upload & proses DBR", s2),
                                    ("Hasil & unduh", s3)], 1):
        ico = "✅" if sts == "done" else ("🔵" if sts == "active" else "⚪")
        st.markdown(f'<div class="sidebar-step {sts or ""}">{ico} Langkah {i}: {lb}</div>',
                    unsafe_allow_html=True)
    st.divider()
    if st.session_state.master_ready:
        st.metric("Label Master", f"{len(st.session_state.catalog):,}")
    st.divider()
    if st.button("🔄 Reset Semua Data", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
    st.caption("Master = validasi. Data cetak dari DBR Excel.")


# ============================================================
# LANGKAH 1 — MASTER
# ============================================================
s1c = "done" if st.session_state.master_ready else "active"
b1 = "done" if st.session_state.master_ready else ""
st.markdown(f"""<div class="step-card {s1c}"><div class="step-header">
<span class="step-badge {b1}">1</span>Master Label (PDF) — Katalog Validasi</div>
<div style="color:#4a5568;font-size:.88rem">Unggah PDF label BMN. Sistem membaca kode &amp; NUP sebagai
<b>katalog validasi</b>. Data cetak diambil dari DBR (langkah 2).</div></div>""",
            unsafe_allow_html=True)

cu, cm = st.columns([2, 1])
with cu:
    upl_pdfs = st.file_uploader("PDF Label", type=["pdf"], accept_multiple_files=True,
                                key="pdf_up", label_visibility="collapsed")
    if upl_pdfs:
        st.caption(f"📎 {len(upl_pdfs)} file")
    if st.button("🎯 JADIKAN DATA UTAMA", type="primary",
                 disabled=not upl_pdfs, use_container_width=True):
        cat, meta = [], []
        bar = st.progress(0)
        for i, f in enumerate(upl_pdfs):
            bar.progress((i+1)/len(upl_pdfs)*.95, text=f"Ekstrak: {f.name}")
            try:
                b = f.read()
                ls = extract_catalog(b, f.name)
                cat.extend(ls)
                meta.append((f.name, len(b)//1024, len(ls)))
            except Exception as e:
                st.error(f"❌ {f.name}: {e}")
        st.session_state.catalog = cat
        st.session_state.master_ready = True
        st.session_state.master_meta = meta
        st.session_state.dbr_processed = False
        st.session_state.ruangan_results = {}
        st.session_state.generated_pdfs = {}
        bar.progress(1.0, text="Selesai!")
        st.rerun()

with cm:
    if st.session_state.master_ready:
        st.markdown(f"""<div class="metric-big"><div class="num">
{len(st.session_state.catalog):,}</div><div class="lbl">Label Katalog</div></div>""",
                    unsafe_allow_html=True)

# Detail master
if st.session_state.master_ready and st.session_state.catalog:
    with st.expander("📊 Detail & Visualisasi Katalog Master"):
        cat = st.session_state.catalog
        cat_df = pd.DataFrame([{"Kode": l.kode_barang, "NUP": l.nup,
                                "Tahun": l.tahun, "Nama": l.nama_barang,
                                "File": l.source_pdf} for l in cat])

        # Grouped view
        st.markdown("#### Kode Barang × Tahun × NUP")
        grp = (cat_df.groupby(["Kode", "Tahun"])
               .agg(Nama=("Nama", "first"), Jml_NUP=("NUP", "count"),
                    NUP_List=("NUP", lambda x: ", ".join(str(v) for v in sorted(x))))
               .reset_index().sort_values(["Kode", "Tahun"]))
        grp.columns = ["Kode Barang", "Tahun", "Nama Barang", "Jumlah NUP", "Daftar NUP"]
        grp["Status"] = grp["Tahun"].apply(
            lambda t: "⚠️ BARANG USANG — PRIORITAS DIHAPUS" if int(t) < 2000 else "✅ Aktif")
        st.dataframe(grp, use_container_width=True, hide_index=True)

        usang = grp[grp["Status"].str.contains("USANG")]
        if not usang.empty:
            st.warning(f"⚠️ **{len(usang)} kelompok** tahun < 2000 "
                       f"(total {usang['Jumlah NUP'].sum()} unit) — prioritas dihapus.")

        st.markdown("#### Distribusi per Tahun")
        tc = cat_df["Tahun"].value_counts().sort_index().reset_index()
        tc.columns = ["Tahun", "Jumlah"]
        st.bar_chart(tc, x="Tahun", y="Jumlah", color="#e65100")


# ============================================================
# LANGKAH 2 — DBR
# ============================================================
st.markdown("<div style='height:.3rem'></div>", unsafe_allow_html=True)
s2_lock = not st.session_state.master_ready
s2c = "locked" if s2_lock else ("done" if st.session_state.dbr_processed else "active")
b2 = "done" if st.session_state.dbr_processed else ""
st.markdown(f"""<div class="step-card {s2c}"><div class="step-header">
<span class="step-badge {b2}">2</span>Daftar Barang Ruangan (DBR)</div>
<div style="color:#4a5568;font-size:.88rem">Unggah Excel DBR. Setiap <b>sheet = ruangan</b>.
Sheet 'Master Aset' otomatis di-skip. <b>Data cetak diambil dari sini.</b></div></div>""",
            unsafe_allow_html=True)

if s2_lock:
    st.markdown('<div class="mini-note">🔒 Selesaikan Langkah 1.</div>', unsafe_allow_html=True)
else:
    upl_dbr = st.file_uploader("Excel DBR", type=["xlsx", "xls"],
                               key="dbr_up", label_visibility="collapsed")
    if upl_dbr:
        try:
            xb = upl_dbr.read()
            xl = pd.ExcelFile(io.BytesIO(xb))
            SKIP = ["master aset", "master"]
            default_sel = [s for s in xl.sheet_names if s.lower().strip() not in SKIP]
            sel = st.multiselect("Sheet ruangan", xl.sheet_names, default=default_sel,
                                 help="Sheet 'Master Aset' di-exclude.")

            if st.button("⚙️ PROSES", type="primary", disabled=not sel, use_container_width=True):
                # Build master index
                master_idx = set()
                for l in st.session_state.catalog:
                    master_idx.add((l.kode_barang, l.nup))

                results, pdfs, failed_sheets = {}, {}, []
                bar = st.progress(0)
                for idx, sh in enumerate(sel):
                    bar.progress((idx+1)/len(sel)*.9, text=f"Proses: {sh}")
                    if sh.lower().strip() in SKIP:
                        continue
                    items, ruangan = detect_dbr(io.BytesIO(xb), sh)
                    if items is None:
                        # Sheet gagal → tetap masukkan dengan 0 items
                        results[sh] = RuanganResult(
                            ruangan=ruangan, sheet_name=sh, items=[])
                        failed_sheets.append(sh)
                        continue
                    # Validasi terhadap master
                    for it in items:
                        it.in_master = (it.kode_barang, it.nup) in master_idx
                    rr = RuanganResult(ruangan=ruangan, sheet_name=sh, items=items)
                    results[sh] = rr
                    # Generate PDF
                    pdf_bytes = build_output_pdf(ruangan, items)
                    if pdf_bytes:
                        pdfs[sh] = pdf_bytes

                st.session_state.ruangan_results = results
                st.session_state.generated_pdfs = pdfs
                st.session_state.dbr_processed = True
                st.session_state.failed_sheets = failed_sheets
                bar.progress(1.0, text=f"Selesai! {len(pdfs)} PDF, "
                             f"{len(failed_sheets)} gagal deteksi.")
                st.rerun()
        except Exception as e:
            st.error(f"❌ Gagal: {e}")


# ============================================================
# LANGKAH 3 — HASIL
# ============================================================
st.markdown("<div style='height:.3rem'></div>", unsafe_allow_html=True)
s3_lock = not st.session_state.dbr_processed
s3c = "locked" if s3_lock else ("done" if st.session_state.generated_pdfs else "active")
st.markdown(f"""<div class="step-card {s3c}"><div class="step-header">
<span class="step-badge">3</span>Hasil Filter, Analisis &amp; Unduh</div>
<div style="color:#4a5568;font-size:.88rem">Preview per ruangan, analisis kondisi barang,
unduh PDF satuan atau ZIP semua.</div></div>""", unsafe_allow_html=True)

if s3_lock:
    st.markdown('<div class="mini-note">🔒 Selesaikan Langkah 2.</div>', unsafe_allow_html=True)
else:
    results = st.session_state.ruangan_results
    pdfs = st.session_state.generated_pdfs

    if not results:
        st.warning("Tidak ada data yang berhasil diproses.")
    else:
        # Tampilkan sheet yang gagal deteksi
        failed = st.session_state.get("failed_sheets", [])
        if failed:
            st.error(
                f"⚠️ **{len(failed)} sheet gagal deteksi data**: "
                f"{', '.join(failed)}. "
                f"Kemungkinan penyebab: format kolom berbeda, "
                f"tidak ada kode barang 7-10 digit, "
                f"atau data dimulai terlalu jauh ke bawah (>200 baris)."
            )
        # Ringkasan
        total_items = sum(len(r.items) for r in results.values())
        total_master = sum(1 for r in results.values() for it in r.items if it.in_master)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ruangan", len(results))
        c2.metric("Total Label", total_items)
        c3.metric("Ada di Master", total_master)
        c4.metric("PDF Ter-generate", len(pdfs))

        # ZIP
        if len(pdfs) > 1:
            zb = io.BytesIO()
            with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as zf:
                for nm, pb in pdfs.items():
                    safe = re.sub(r"[^\w\-. ]", "_", nm).strip()
                    zf.writestr(f"Label_{safe}.pdf", pb)
            st.download_button("📦 Unduh SEMUA (ZIP)", data=zb.getvalue(),
                               file_name="Label_Semua_Ruangan.zip",
                               mime="application/zip", type="primary",
                               use_container_width=True)

        # ===== ANALISIS KONDISI GLOBAL =====
        st.markdown("---")
        st.markdown("### 📊 Analisis Kondisi Barang")

        all_items_data = []
        for sh, rr in results.items():
            for it in rr.items:
                kondisi = classify_condition(it.keterangan)
                all_items_data.append({
                    "ruangan": rr.ruangan, "sheet": sh,
                    "kode_barang": it.kode_barang, "nup": it.nup,
                    "nama_barang": it.nama_barang, "merk": it.merk,
                    "tahun": it.tahun, "keterangan_raw": it.keterangan,
                    "kondisi": kondisi,
                    "skor_prioritas": round(priority_score(kondisi, it.tahun), 1),
                    "rekomendasi": priority_label(kondisi, it.tahun),
                })

        if all_items_data:
            adf = pd.DataFrame(all_items_data)

            # Card kondisi
            st.markdown("#### Distribusi Kondisi")
            ccounts = adf["kondisi"].value_counts()
            cc = st.columns(min(len(ccounts), 4))
            for ci, (kond, jml) in enumerate(ccounts.items()):
                d = COND_DISPLAY.get(kond, ("?", "#999", "#eee"))
                with cc[ci % len(cc)]:
                    st.markdown(f"""<div class="cond-card" style="background:{d[2]};
                    border-left:4px solid {d[1]}">
                    <div style="font-size:1.8rem;font-weight:700;color:{d[1]}">{jml}</div>
                    <div style="font-size:.78rem;color:{d[1]};font-weight:600">{d[0]}</div>
                    </div>""", unsafe_allow_html=True)

            # ===== DETAIL BARANG PERLU DIGANTI =====
            st.markdown("#### 🔴 Barang yang Perlu Diganti / Dihapus")
            ganti = adf[adf["kondisi"] == "RUSAK"].copy()
            if not ganti.empty:
                ganti_view = ganti[["ruangan", "nama_barang", "kode_barang",
                                    "nup", "tahun", "merk", "skor_prioritas",
                                    "rekomendasi"]].sort_values("skor_prioritas",
                                                                ascending=False)
                ganti_view.columns = ["Ruangan", "Nama Barang", "Kode", "NUP",
                                      "Tahun", "Merk", "Skor Prioritas", "Rekomendasi"]
                st.dataframe(ganti_view, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Tidak ada barang dengan status RUSAK.")

            st.markdown("#### 🟡 Barang yang Perlu Perbaikan")
            perbaiki = adf[adf["kondisi"] == "RUSAK RINGAN"].copy()
            if not perbaiki.empty:
                pb_view = perbaiki[["ruangan", "nama_barang", "kode_barang",
                                    "nup", "tahun", "merk", "skor_prioritas",
                                    "rekomendasi"]].sort_values("skor_prioritas",
                                                                ascending=False)
                pb_view.columns = ["Ruangan", "Nama Barang", "Kode", "NUP",
                                   "Tahun", "Merk", "Skor Prioritas", "Rekomendasi"]
                st.dataframe(pb_view, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Tidak ada barang rusak ringan.")

            # Barang usang (tahun < 2000)
            usang_items = adf[adf["tahun"].apply(
                lambda t: int(t) < 2000 if t and str(t).isdigit() else False)].copy()
            if not usang_items.empty:
                st.markdown("#### ⚠️ Barang Usang (Tahun < 2000) — Prioritas Dihapus")
                uv = usang_items[["ruangan", "nama_barang", "kode_barang",
                                  "nup", "tahun", "kondisi",
                                  "rekomendasi"]].sort_values("tahun")
                uv.columns = ["Ruangan", "Nama Barang", "Kode", "NUP",
                              "Tahun", "Kondisi", "Rekomendasi"]
                st.dataframe(uv, use_container_width=True, hide_index=True)

            # Chart: ruangan vs jumlah rusak
            st.markdown("#### 🏚️ Ruangan dengan Barang Rusak Terbanyak")
            rusak_all = adf[adf["kondisi"].isin(["RUSAK", "RUSAK RINGAN"])]
            if not rusak_all.empty:
                rpr = (rusak_all.groupby("ruangan").size()
                       .reset_index(name="Jumlah Rusak")
                       .sort_values("Jumlah Rusak", ascending=False))
                st.bar_chart(rpr.set_index("ruangan")["Jumlah Rusak"], color="#c62828")
                top = rpr.iloc[0]
                st.info(f"📌 **{top['ruangan']}** memiliki **{top['Jumlah Rusak']}** "
                        f"barang rusak — paling banyak.")
            else:
                st.success("✅ Tidak ada barang rusak.")

        # ===== TABS PER RUANGAN (SEMUA, termasuk yang gagal/kosong) =====
        st.markdown("---")
        st.markdown("### 📑 Detail per Ruangan")
        all_sheet_names = list(results.keys())
        tabs = st.tabs([f"{results[s].ruangan}" for s in all_sheet_names])
        for tab, sh in zip(tabs, all_sheet_names):
            rr = results[sh]
            with tab:
                if not rr.items:
                    st.warning(
                        f"⚠️ **{rr.ruangan}** — Tidak ada data terdeteksi di sheet ini. "
                        f"Kemungkinan: format kolom berbeda, sheet kosong, "
                        f"atau kode barang tidak ditemukan."
                    )
                    continue

                ca, cb, cc = st.columns(3)
                ca.metric("Jumlah Label", len(rr.items))
                in_m = sum(1 for it in rr.items if it.in_master)
                cb.metric("Ada di Master", in_m)
                not_in = len(rr.items) - in_m
                cc.metric("Tidak di Master", not_in,
                          delta=str(not_in) if not_in else None, delta_color="inverse")

                if sh in pdfs:
                    safe = re.sub(r"[^\w\-. ]", "_", sh).strip()
                    st.download_button(f"⬇️ Unduh PDF — {rr.ruangan}",
                                       data=pdfs[sh],
                                       file_name=f"Label_{safe}.pdf",
                                       mime="application/pdf",
                                       key=f"dl_{sh}", type="primary")

                # Kondisi per ruangan
                room_data = [d for d in all_items_data if d["sheet"] == sh]
                if room_data:
                    rdf = pd.DataFrame(room_data)
                    rc = rdf["kondisi"].value_counts()
                    rcols = st.columns(min(len(rc), 4))
                    for ci, (k, j) in enumerate(rc.items()):
                        d = COND_DISPLAY.get(k, ("?", "#999", "#eee"))
                        with rcols[ci % len(rcols)]:
                            st.markdown(f"""<div class="cond-card" style="background:{d[2]};
                            border-left:3px solid {d[1]}"><b style="color:{d[1]}">{j}</b>
                            <span style="font-size:.75rem;color:{d[1]}"> {d[0]}</span></div>""",
                                        unsafe_allow_html=True)

                # Tabel detail
                st.markdown("**Detail barang:**")
                dtbl = pd.DataFrame([{
                    "No": i+1, "Kode": it.kode_barang, "NUP": it.nup,
                    "Tahun": it.tahun, "Nama": it.nama_barang,
                    "Merk": it.merk, "Kondisi": classify_condition(it.keterangan),
                    "✓ Master": "✅" if it.in_master else "❌",
                } for i, it in enumerate(rr.items)])
                st.dataframe(dtbl, use_container_width=True, hide_index=True)


# ============================================================
# FOOTER
# ============================================================
st.markdown("""<div style="text-align:center;color:#a0aec0;font-size:.78rem;
margin-top:2rem;padding-top:.8rem;border-top:1px solid #edf2f7">
BBPPMPV Pertanian · Cianjur · Sistem Label BMN v4.0</div>""", unsafe_allow_html=True)
