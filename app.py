"""
Sistem Cetak Label Barang Ruangan
BBPPMPV Pertanian - Cianjur

Alur:
1. Upload master PDF label (bisa banyak file) → "JADIKAN DATA UTAMA"
   → sistem ekstrak katalog label (kode, NUP, tahun, bbox).
2. Upload DBR Excel/CSV → "PROSES"
   → sistem baca setiap sheet (= ruangan), match tiap baris ke label master.
3. Preview hasil per ruangan → download PDF per ruangan atau ZIP semua.
"""

import io
import re
import zipfile
from dataclasses import dataclass, field
from typing import Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# ============================================================
# KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(
    page_title="Sistem Cetak Label Barang Ruangan",
    page_icon="🏷️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CUSTOM CSS – UI PROFESIONAL
# ============================================================
st.markdown(
    """
    <style>
    /* ---------- Global ---------- */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1300px;
    }
    /* ---------- Header banner ---------- */
    .app-header {
        background: linear-gradient(135deg, #0f4c81 0%, #1976d2 100%);
        color: #fff;
        padding: 1.5rem 2rem;
        border-radius: 14px;
        margin-bottom: 1.5rem;
        box-shadow: 0 6px 18px rgba(15,76,129,0.18);
    }
    .app-header h1 { color: #fff; margin: 0; font-size: 1.6rem; font-weight: 700; }
    .app-header p  { color: #d6e6f5; margin: 0.35rem 0 0 0; font-size: 0.95rem; }

    /* ---------- Step card ---------- */
    .step-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .step-card.active   { border-left: 4px solid #1976d2; }
    .step-card.done     { border-left: 4px solid #2e7d32; background: #f4faf5; }
    .step-card.locked   { opacity: 0.55; border-left: 4px solid #cbd5e0; }
    .step-header {
        display: flex; align-items: center; gap: 0.75rem;
        font-size: 1.1rem; font-weight: 600; color: #1a202c; margin-bottom: 0.25rem;
    }
    .step-badge {
        background: #1976d2; color: #fff; width: 28px; height: 28px;
        border-radius: 50%; display: flex; align-items: center; justify-content: center;
        font-size: 0.85rem; font-weight: 700;
    }
    .step-badge.done { background: #2e7d32; }

    /* ---------- Metric ---------- */
    .metric-big {
        background: linear-gradient(135deg, #1976d2 0%, #1565c0 100%);
        color: #fff; padding: 1.25rem; border-radius: 12px; text-align: center;
    }
    .metric-big .num { font-size: 2.5rem; font-weight: 700; line-height: 1; }
    .metric-big .lbl { font-size: 0.85rem; opacity: 0.9; margin-top: 0.35rem;
                       text-transform: uppercase; letter-spacing: 0.5px; }

    /* ---------- Sidebar ---------- */
    section[data-testid="stSidebar"] { background: #f8fafc; }
    .sidebar-step {
        display: flex; align-items: center; gap: 0.5rem;
        padding: 0.55rem 0.75rem; border-radius: 8px; margin-bottom: 0.4rem;
        font-size: 0.9rem; color: #4a5568;
    }
    .sidebar-step.active { background: #e3f2fd; color: #0d47a1; font-weight: 600; }
    .sidebar-step.done   { background: #e8f5e9; color: #1b5e20; }
    .sidebar-step .ico   { font-size: 1.1rem; }

    /* ---------- Buttons ---------- */
    .stButton > button {
        border-radius: 8px; font-weight: 600; padding: 0.55rem 1.2rem;
        transition: all 0.15s ease;
    }
    .stButton > button[kind="primary"] {
        background: #1976d2; border-color: #1976d2;
    }
    .stButton > button[kind="primary"]:hover {
        background: #1565c0; border-color: #1565c0;
        transform: translateY(-1px); box-shadow: 0 4px 10px rgba(25,118,210,0.3);
    }

    /* ---------- File uploader kecil ---------- */
    [data-testid="stFileUploaderDropzone"] {
        padding: 1rem; border-radius: 10px; border-style: dashed;
    }

    /* ---------- Tab styling ---------- */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0; padding: 0.5rem 1rem;
    }

    /* ---------- Info box kecil ---------- */
    .mini-note {
        font-size: 0.82rem; color: #4a5568; background: #f7fafc;
        border-left: 3px solid #a0aec0; padding: 0.5rem 0.75rem;
        border-radius: 4px; margin: 0.5rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# DATA CLASSES
# ============================================================
@dataclass
class LabelInfo:
    """Katalog satu label dari master PDF."""
    kode_barang: str            # 10 digit, e.g. "3030212012"
    nup: int                    # e.g. 1
    tahun: str                  # e.g. "2019"
    source_pdf: str             # nama file PDF asal
    page_index: int             # 0-based
    bbox: tuple                 # (x0, y0, x1, y1) di halaman sumber
    nama_barang: str = ""       # opsional, dari teks label
    doc_id: int = 0             # index dokumen di catalog (untuk retrieval)

@dataclass
class DBRRow:
    """Satu baris DBR yang sudah di-parse."""
    no_urut: str
    nup_list: list
    nama_barang: str
    merk: str
    kode_barang: str
    tahun_list: list
    jumlah: str
    keterangan: str

@dataclass
class MatchResult:
    ruangan: str
    matched: list = field(default_factory=list)      # list[LabelInfo]
    not_found: list = field(default_factory=list)    # list[dict] – row info yang tidak ketemu

# ============================================================
# HELPER FUNCTIONS
# ============================================================
KODE_RE = re.compile(r"\b(\d{10})\b")
NUP_RE = re.compile(r"NUP\s*:?\s*(\d+)", re.IGNORECASE)
TAHUN_RE = re.compile(r"KD\s*\.\s*(\d{4})")


def parse_multi_number(value) -> list:
    """
    Parse cell yang bisa berisi angka tunggal atau list:
    '9, 10' → [9, 10] ; "'23" → [23] ; '100, 99' → [99, 100] ; 47 → [47]
    Return: list int terurut ascending, unique.
    """
    if value is None:
        return []
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return []
        return [int(value)]
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return []
    # bersihkan quote dan whitespace
    s = s.replace("'", "").replace('"', "").replace("`", "")
    # split by koma / titik-koma / slash
    parts = re.split(r"[,;/]", s)
    result = set()
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # bisa berupa "9.0" dari Excel numeric
        m = re.search(r"\d+", p)
        if m:
            try:
                result.add(int(m.group()))
            except ValueError:
                pass
    return sorted(result)


def clean_kode_barang(value) -> str:
    """Normalisasi kode barang ke string 10 digit (Excel sering muncul sebagai float)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip().replace("'", "").replace(" ", "")
    # buang trailing .0
    if s.endswith(".0"):
        s = s[:-2]
    # ambil digit saja
    m = re.search(r"\d{6,}", s)
    return m.group() if m else s


def extract_labels_from_pdf(pdf_bytes: bytes, source_name: str, doc_id: int) -> tuple:
    """
    Ekstrak seluruh label dari satu file PDF.

    Strategi grid-based (deterministik):
      1. Buka dengan PyMuPDF.
      2. Untuk setiap halaman, kumpulkan SEMUA text span.
      3. Deteksi ROW grid: cluster Y-position dari span "KEMENTERIAN"
         (setiap row = satu baris label).
      4. Tinggi cell = jarak antar row Y (konsisten untuk semua baris,
         termasuk baris terakhir).
      5. Setiap NUP di dalam row → satu label dengan bbox konsisten.

    Return: (list[LabelInfo], fitz.Document handle)
            Dokumen tetap dipertahankan karena bbox merujuk ke halaman-nya.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    labels: list[LabelInfo] = []

    for page_index in range(doc.page_count):
        page = doc[page_index]
        page_w = page.rect.width
        page_h = page.rect.height

        # Kumpulkan semua span (text, bbox)
        spans = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type", 0) != 0:
                continue
            for line in b.get("lines", []):
                for sp in line.get("spans", []):
                    txt = sp.get("text", "").strip()
                    if txt:
                        spans.append((txt, tuple(sp["bbox"])))

        if not spans:
            continue

        # --- 1. Deteksi ROW grid dari span "KEMENTERIAN" ---
        kementerian_ys = sorted(set(
            round(bb[1], 1) for t, bb in spans
            if "KEMENTERIAN" in t.upper()
        ))
        if not kementerian_ys:
            # halaman tanpa label
            continue

        # Cluster Y yang berdekatan (< 8 pt) → 1 row
        row_tops = []
        for y in kementerian_ys:
            if not row_tops or (y - row_tops[-1]) > 8:
                row_tops.append(y)

        # --- 2. Hitung tinggi row (konsisten) ---
        if len(row_tops) >= 2:
            diffs = [row_tops[i + 1] - row_tops[i] for i in range(len(row_tops) - 1)]
            row_height = min(diffs)  # gunakan gap terkecil (biasanya jarak antar row)
        else:
            # fallback: cari NUP terjauh dari row_tops[0], buffer 15%
            row_height = 100

        # padding di atas untuk logo/garis border
        top_pad = 6
        bottom_pad = 4

        # --- 3. Untuk tiap row, deteksi kolom dari X-position ---
        # 2 kolom: kiri (x < page_w/2), kanan
        for row_top in row_tops:
            cell_y0 = max(0, row_top - top_pad)
            cell_y1 = min(page_h, row_top + row_height - bottom_pad)

            for col in (0, 1):
                col_x0 = 0 if col == 0 else page_w / 2
                col_x1 = page_w / 2 if col == 0 else page_w
                # padding kiri/kanan
                cell_x0 = max(0, col_x0 + 4)
                cell_x1 = min(page_w, col_x1 - 4)
                cell_bbox = (cell_x0, cell_y0, cell_x1, cell_y1)

                # kumpulkan span dalam cell
                cell_spans = []
                for t, bb in spans:
                    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
                    if cell_x0 <= cx <= cell_x1 and cell_y0 <= cy <= cell_y1:
                        cell_spans.append((bb[1], bb[0], t, bb))
                if not cell_spans:
                    continue
                cell_spans.sort()
                cell_text = " ".join(t for _, _, t, _ in cell_spans)

                # parse identifier
                m_nup = NUP_RE.search(cell_text)
                m_kode = KODE_RE.search(cell_text)
                m_tahun = TAHUN_RE.search(cell_text)
                if not (m_nup and m_kode and m_tahun):
                    continue

                try:
                    nup_val = int(m_nup.group(1))
                except ValueError:
                    continue
                kode = m_kode.group(1)
                tahun = m_tahun.group(1)

                # nama barang: baris pertama SETELAH baris kode+NUP
                # baris kode+NUP dikenali dari Y span yg mengandung kode
                kode_y = None
                for y, _, t, bb in cell_spans:
                    if kode in t:
                        kode_y = y
                        break
                nama = ""
                if kode_y is not None:
                    for y, _, t, bb in cell_spans:
                        if y > kode_y + 3 and y < kode_y + 25:
                            # skip kalau baris ini masih berisi kode/NUP
                            if kode in t or "NUP" in t.upper():
                                continue
                            nama = t
                            break

                labels.append(LabelInfo(
                    kode_barang=kode,
                    nup=nup_val,
                    tahun=tahun,
                    source_pdf=source_name,
                    page_index=page_index,
                    bbox=cell_bbox,
                    nama_barang=nama,
                    doc_id=doc_id,
                ))

    return labels, doc


def detect_dbr_data(xl_file, sheet_name: str) -> Optional[pd.DataFrame]:
    """
    Auto-detect baris header di sheet DBR, lalu return DataFrame data barang.
    Baris header dikenali dari kemunculan kata kunci 'Nomor Urut' dan 'Kode Barang'.
    """
    try:
        raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=object)
    except Exception:
        return None

    header_row = None
    for i in range(min(len(raw), 40)):
        row_texts = [str(x).lower() for x in raw.iloc[i].values if pd.notna(x)]
        joined = " ".join(row_texts)
        if ("nomor urut" in joined or "no. urut" in joined) and "kode barang" in joined:
            header_row = i
            break
        # fallback: cari kombinasi "nama barang" + "kode barang"
        if "nama barang" in joined and "kode barang" in joined:
            header_row = i
            break

    if header_row is None:
        return None

    # Ambil data mulai baris setelah header. Untuk mengatasi multi-row header
    # (merged cells), coba geser 1-2 baris ke bawah sampai menemukan baris
    # numerik di kolom pertama.
    df = raw.iloc[header_row:].reset_index(drop=True)
    # Set header
    df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}"
                  for i, c in enumerate(df.iloc[0].values)]
    df = df.iloc[1:].reset_index(drop=True)

    # buang baris kosong
    df = df.dropna(how="all")
    return df


def find_col(df: pd.DataFrame, keywords: list) -> Optional[str]:
    """Cari nama kolom yang cocok dengan salah satu keyword (case-insensitive)."""
    for col in df.columns:
        col_low = str(col).lower()
        for kw in keywords:
            if kw in col_low:
                return col
    return None


def parse_dbr_dataframe(df: pd.DataFrame) -> list:
    """Ubah dataframe raw jadi list[DBRRow]."""
    if df is None or df.empty:
        return []

    col_no = find_col(df, ["no. urut", "no urut"])
    col_nup = find_col(df, ["nomor urut pendaftaran", "pendaftaran", "nup"])
    col_nama = find_col(df, ["nama barang"])
    col_merk = find_col(df, ["merk", "type"])
    col_kode = find_col(df, ["kode barang"])
    col_tahun = find_col(df, ["tahun perolehan", "tahun"])
    col_jml = find_col(df, ["jumlah"])
    col_ket = find_col(df, ["keterangan"])

    if not col_nup or not col_kode:
        return []

    rows = []
    for _, r in df.iterrows():
        kode = clean_kode_barang(r.get(col_kode))
        nup_list = parse_multi_number(r.get(col_nup))
        if not kode or not nup_list:
            continue
        tahun_list = parse_multi_number(r.get(col_tahun)) if col_tahun else []
        rows.append(DBRRow(
            no_urut=str(r.get(col_no, "")).strip() if col_no else "",
            nup_list=nup_list,
            nama_barang=str(r.get(col_nama, "")).strip() if col_nama else "",
            merk=str(r.get(col_merk, "")).strip() if col_merk else "",
            kode_barang=kode,
            tahun_list=[str(t) for t in tahun_list],
            jumlah=str(r.get(col_jml, "")).strip() if col_jml else "",
            keterangan=str(r.get(col_ket, "")).strip() if col_ket else "",
        ))
    return rows


def match_ruangan(dbr_rows: list, catalog: list) -> MatchResult:
    """
    Match tiap baris DBR ke label master.
    Strategi: (kode_barang, nup) exact match. Tahun tidak divalidasi ketat
    karena data DBR sering single-value padahal aktualnya multi.
    """
    # Index catalog by (kode, nup) → list[LabelInfo]  (bisa >1 kalau duplicate)
    index = {}
    for lbl in catalog:
        key = (lbl.kode_barang, lbl.nup)
        index.setdefault(key, []).append(lbl)

    result = MatchResult(ruangan="")
    for row in dbr_rows:
        for nup in row.nup_list:
            key = (row.kode_barang, nup)
            hits = index.get(key, [])
            if hits:
                # ambil kandidat terbaik: kalau ada preferensi tahun di DBR
                # dan salah satu label tahun-nya cocok, prioritaskan itu.
                chosen = None
                if row.tahun_list:
                    for h in hits:
                        if h.tahun in row.tahun_list:
                            chosen = h
                            break
                if chosen is None:
                    chosen = hits[0]
                result.matched.append(chosen)
            else:
                result.not_found.append({
                    "kode_barang": row.kode_barang,
                    "nup": nup,
                    "nama_barang": row.nama_barang,
                    "tahun": ", ".join(row.tahun_list),
                })
    return result


# ============================================================
# PDF GENERATOR
# ============================================================
def build_output_pdf(matched: list, docs_by_id: dict) -> bytes:
    """
    Buat PDF baru berisi label yang di-crop dari master.
    Layout: 2 kolom, ukuran cell mengikuti proporsi ~ label master.
    Menggunakan Page.show_pdf_page() → vector preserved (QR code tetap tajam).
    """
    if not matched:
        return b""

    out = fitz.open()

    # Ukuran halaman A4 portrait
    PAGE_W, PAGE_H = 595, 842   # points
    MARGIN_X = 30
    MARGIN_TOP = 40
    MARGIN_BOTTOM = 30
    GUTTER_X = 12
    GUTTER_Y = 10

    # 2 kolom
    cell_w = (PAGE_W - 2 * MARGIN_X - GUTTER_X) / 2

    # ukuran label master ~ (bbox lebar & tinggi) → hitung tinggi cell
    # pakai contoh dari label pertama
    sample = matched[0]
    src_bbox = sample.bbox
    src_w = src_bbox[2] - src_bbox[0]
    src_h = src_bbox[3] - src_bbox[1]
    # scale factor untuk fit ke cell_w
    scale = cell_w / src_w
    cell_h = src_h * scale

    # jumlah row per halaman
    usable_h = PAGE_H - MARGIN_TOP - MARGIN_BOTTOM
    rows_per_page = max(1, int((usable_h + GUTTER_Y) // (cell_h + GUTTER_Y)))
    cells_per_page = rows_per_page * 2

    for i, lbl in enumerate(matched):
        pos_in_page = i % cells_per_page
        col = pos_in_page % 2
        row = pos_in_page // 2

        if pos_in_page == 0:
            page = out.new_page(width=PAGE_W, height=PAGE_H)

        x0 = MARGIN_X + col * (cell_w + GUTTER_X)
        y0 = MARGIN_TOP + row * (cell_h + GUTTER_Y)
        target_rect = fitz.Rect(x0, y0, x0 + cell_w, y0 + cell_h)

        src_doc = docs_by_id.get(lbl.doc_id)
        if src_doc is None:
            continue

        # Vector-preserving crop
        page.show_pdf_page(
            target_rect,
            src_doc,
            lbl.page_index,
            clip=fitz.Rect(*lbl.bbox),
        )

    buf = io.BytesIO()
    out.save(buf, deflate=True)
    out.close()
    return buf.getvalue()


# ============================================================
# SESSION STATE
# ============================================================
def init_state():
    defaults = {
        "catalog": [],           # list[LabelInfo]
        "docs_by_id": {},        # doc_id → fitz.Document
        "master_ready": False,   # sudah "Jadikan Data Utama"?
        "master_files_meta": [], # [(name, size_kb, label_count), ...]
        "dbr_processed": False,
        "match_results": {},     # ruangan → MatchResult
        "generated_pdfs": {},    # ruangan → bytes
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ============================================================
# HEADER
# ============================================================
st.markdown(
    """
    <div class="app-header">
      <h1>🏷️ Sistem Cetak Label Barang Ruangan</h1>
      <p>BBPPMPV Pertanian · Cianjur — Otomatisasi filter & cetak label BMN per ruangan berbasis DBR</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# SIDEBAR – STATUS LANGKAH
# ============================================================
with st.sidebar:
    st.markdown("### 📋 Progres")

    step1_status = "done" if st.session_state.master_ready else "active"
    step2_status = (
        "done" if st.session_state.dbr_processed
        else ("active" if st.session_state.master_ready else "")
    )
    step3_status = (
        "done" if st.session_state.generated_pdfs
        else ("active" if st.session_state.dbr_processed else "")
    )

    for i, (label, status) in enumerate([
        ("Upload & set master label", step1_status),
        ("Upload & proses DBR", step2_status),
        ("Preview & unduh hasil", step3_status),
    ], start=1):
        icon = "✅" if status == "done" else ("🔵" if status == "active" else "⚪")
        css = status if status else ""
        st.markdown(
            f'<div class="sidebar-step {css}"><span class="ico">{icon}</span>'
            f'<span>Langkah {i}: {label}</span></div>',
            unsafe_allow_html=True,
        )

    st.divider()

    if st.session_state.master_ready:
        st.metric("Total Label Master", f"{len(st.session_state.catalog):,}")
        st.caption(f"dari {len(st.session_state.master_files_meta)} file PDF")

    st.divider()
    if st.button("🔄 Reset Semua Data", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    st.caption(
        "💡 Master label = referensi utama. DBR = filter per ruangan. "
        "Sistem match berdasarkan (Kode Barang, NUP)."
    )


# ============================================================
# LANGKAH 1 – UPLOAD & JADIKAN DATA UTAMA
# ============================================================
step1_class = "done" if st.session_state.master_ready else "active"
st.markdown(
    f'<div class="step-card {step1_class}">'
    f'<div class="step-header"><span class="step-badge {"done" if st.session_state.master_ready else ""}">1</span>'
    f'Master Label Barang (PDF)</div>'
    f'<div style="color:#4a5568;font-size:0.9rem">'
    f'Unggah satu atau lebih file PDF label barang. Setelah klik <b>JADIKAN DATA UTAMA</b>, '
    f'sistem akan mengekstrak katalog label yang bisa difilter oleh DBR.'
    f'</div></div>',
    unsafe_allow_html=True,
)

col_up, col_metric = st.columns([2, 1])

with col_up:
    uploaded_pdfs = st.file_uploader(
        "Pilih file PDF label",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_uploader",
        label_visibility="collapsed",
    )
    if uploaded_pdfs:
        st.caption(f"📎 {len(uploaded_pdfs)} file terpilih: "
                   + ", ".join(f.name for f in uploaded_pdfs))

    btn_disabled = not uploaded_pdfs
    if st.button("🎯 JADIKAN DATA UTAMA", type="primary",
                 disabled=btn_disabled, use_container_width=True):
        # ekstrak semua file
        catalog = []
        docs_by_id = {}
        meta = []
        progress = st.progress(0, text="Memulai ekstraksi...")
        for i, f in enumerate(uploaded_pdfs):
            progress.progress((i + 1) / len(uploaded_pdfs) * 0.95,
                              text=f"Mengekstrak: {f.name}")
            try:
                pdf_bytes = f.read()
                labels, doc = extract_labels_from_pdf(pdf_bytes, f.name, doc_id=i)
                catalog.extend(labels)
                docs_by_id[i] = doc
                meta.append((f.name, len(pdf_bytes) // 1024, len(labels)))
            except Exception as e:
                st.error(f"❌ Gagal memproses `{f.name}`: {e}")

        st.session_state.catalog = catalog
        st.session_state.docs_by_id = docs_by_id
        st.session_state.master_files_meta = meta
        st.session_state.master_ready = True
        # reset downstream
        st.session_state.dbr_processed = False
        st.session_state.match_results = {}
        st.session_state.generated_pdfs = {}
        progress.progress(1.0, text="Selesai!")
        st.rerun()

with col_metric:
    if st.session_state.master_ready:
        st.markdown(
            f'<div class="metric-big">'
            f'<div class="num">{len(st.session_state.catalog):,}</div>'
            f'<div class="lbl">Total Label Terdeteksi</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

# Detail per file
if st.session_state.master_ready and st.session_state.master_files_meta:
    with st.expander("📄 Detail file master", expanded=False):
        meta_df = pd.DataFrame(
            st.session_state.master_files_meta,
            columns=["Nama File", "Ukuran (KB)", "Jumlah Label"],
        )
        st.dataframe(meta_df, use_container_width=True, hide_index=True)

        # sample label
        if st.session_state.catalog:
            st.markdown("**Sample katalog (10 label pertama):**")
            sample = pd.DataFrame([
                {"Kode Barang": l.kode_barang, "NUP": l.nup, "Tahun": l.tahun,
                 "Nama": l.nama_barang, "Sumber": l.source_pdf, "Hal.": l.page_index + 1}
                for l in st.session_state.catalog[:10]
            ])
            st.dataframe(sample, use_container_width=True, hide_index=True)


st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ============================================================
# LANGKAH 2 – UPLOAD & PROSES DBR
# ============================================================
step2_locked = not st.session_state.master_ready
step2_class = (
    "locked" if step2_locked
    else ("done" if st.session_state.dbr_processed else "active")
)
st.markdown(
    f'<div class="step-card {step2_class}">'
    f'<div class="step-header"><span class="step-badge {"done" if st.session_state.dbr_processed else ""}">2</span>'
    f'Daftar Barang Ruangan (DBR)</div>'
    f'<div style="color:#4a5568;font-size:0.9rem">'
    f'Unggah file DBR (Excel). Setiap <b>sheet</b> mewakili satu ruangan. '
    f'Klik <b>PROSES</b> untuk mem-filter label master sesuai NUP & kode barang di tiap ruangan.'
    f'</div></div>',
    unsafe_allow_html=True,
)

if step2_locked:
    st.markdown(
        '<div class="mini-note">🔒 Selesaikan Langkah 1 terlebih dahulu.</div>',
        unsafe_allow_html=True,
    )
else:
    uploaded_dbr = st.file_uploader(
        "Pilih file DBR (Excel)",
        type=["xlsx", "xls"],
        key="dbr_uploader",
        label_visibility="collapsed",
    )

    if uploaded_dbr:
        try:
            xl_bytes = uploaded_dbr.read()
            xl = pd.ExcelFile(io.BytesIO(xl_bytes))
            all_sheets = xl.sheet_names
            st.caption(f"📑 File: **{uploaded_dbr.name}** — {len(all_sheets)} sheet ditemukan")

            # Pilih sheet mana yang mau diproses (default: semua kecuali yang kelihatan bukan ruangan)
            EXCLUDE_KW = ["master aset", "master", "dbr", "index", "cover", "kop"]
            default_selected = [
                s for s in all_sheets
                if not any(k in s.lower() for k in EXCLUDE_KW)
            ]
            if not default_selected:
                default_selected = all_sheets

            selected_sheets = st.multiselect(
                "Sheet ruangan yang akan diproses",
                options=all_sheets,
                default=default_selected,
                help="Kecualikan sheet 'Master Aset' atau sheet non-ruangan lain.",
            )

            if st.button("⚙️ PROSES", type="primary",
                         disabled=not selected_sheets, use_container_width=True):
                match_results = {}
                generated_pdfs = {}
                progress = st.progress(0, text="Memulai pemrosesan...")

                for idx, sheet in enumerate(selected_sheets):
                    progress.progress(
                        (idx + 1) / len(selected_sheets) * 0.9,
                        text=f"Memproses sheet: {sheet}",
                    )
                    df = detect_dbr_data(io.BytesIO(xl_bytes), sheet)
                    if df is None or df.empty:
                        continue
                    dbr_rows = parse_dbr_dataframe(df)
                    if not dbr_rows:
                        continue
                    result = match_ruangan(dbr_rows, st.session_state.catalog)
                    result.ruangan = sheet
                    match_results[sheet] = result

                    # generate PDF
                    if result.matched:
                        pdf_bytes = build_output_pdf(
                            result.matched, st.session_state.docs_by_id
                        )
                        if pdf_bytes:
                            generated_pdfs[sheet] = pdf_bytes

                st.session_state.match_results = match_results
                st.session_state.generated_pdfs = generated_pdfs
                st.session_state.dbr_processed = True
                progress.progress(1.0, text="Selesai!")
                st.rerun()

        except Exception as e:
            st.error(f"❌ Gagal membaca file DBR: {e}")


st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ============================================================
# LANGKAH 3 – PREVIEW & UNDUH HASIL
# ============================================================
step3_locked = not st.session_state.dbr_processed
step3_class = (
    "locked" if step3_locked
    else ("done" if st.session_state.generated_pdfs else "active")
)
st.markdown(
    f'<div class="step-card {step3_class}">'
    f'<div class="step-header"><span class="step-badge">3</span>'
    f'Hasil Filter & Unduh</div>'
    f'<div style="color:#4a5568;font-size:0.9rem">'
    f'Preview hasil per ruangan. Setiap ruangan menghasilkan satu PDF berisi label yang '
    f'sudah difilter. Bisa unduh satuan atau semua sekaligus (ZIP).'
    f'</div></div>',
    unsafe_allow_html=True,
)

if step3_locked:
    st.markdown(
        '<div class="mini-note">🔒 Selesaikan Langkah 2 terlebih dahulu.</div>',
        unsafe_allow_html=True,
    )
else:
    results = st.session_state.match_results
    pdfs = st.session_state.generated_pdfs

    if not results:
        st.warning("Tidak ada data yang berhasil diproses dari DBR.")
    else:
        # Ringkasan
        total_matched = sum(len(r.matched) for r in results.values())
        total_not_found = sum(len(r.not_found) for r in results.values())
        total_ruangan = len(results)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ruangan Diproses", total_ruangan)
        c2.metric("Label Ditemukan", total_matched)
        c3.metric("Tidak Ditemukan", total_not_found,
                  delta=None if total_not_found == 0 else f"{total_not_found} item",
                  delta_color="inverse")
        c4.metric("PDF Ter-generate", len(pdfs))

        # Download semua ZIP
        if len(pdfs) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for ruangan, pdf_bytes in pdfs.items():
                    safe_name = re.sub(r"[^\w\-. ]", "_", ruangan).strip() or "ruangan"
                    zf.writestr(f"Label_{safe_name}.pdf", pdf_bytes)
            st.download_button(
                "📦 Unduh SEMUA (ZIP)",
                data=zip_buf.getvalue(),
                file_name="Label_Semua_Ruangan.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

        st.markdown("---")

        # Tab per ruangan
        tabs = st.tabs(list(results.keys()))
        for tab, (ruangan, result) in zip(tabs, results.items()):
            with tab:
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("Label Match", len(result.matched))
                col_b.metric("Tidak Ditemukan", len(result.not_found))
                col_c.metric("Kode Barang Unik",
                             len(set(l.kode_barang for l in result.matched)))

                if ruangan in pdfs:
                    safe_name = re.sub(r"[^\w\-. ]", "_", ruangan).strip() or "ruangan"
                    st.download_button(
                        f"⬇️ Unduh PDF — {ruangan}",
                        data=pdfs[ruangan],
                        file_name=f"Label_{safe_name}.pdf",
                        mime="application/pdf",
                        key=f"dl_{ruangan}",
                        type="primary",
                    )

                # Preview label yang match
                if result.matched:
                    st.markdown("**Label yang berhasil difilter:**")
                    matched_df = pd.DataFrame([
                        {
                            "No": i + 1,
                            "Kode Barang": l.kode_barang,
                            "NUP": l.nup,
                            "Tahun": l.tahun,
                            "Nama Barang": l.nama_barang,
                            "Sumber PDF": l.source_pdf,
                            "Hal.": l.page_index + 1,
                        }
                        for i, l in enumerate(result.matched)
                    ])
                    st.dataframe(matched_df, use_container_width=True, hide_index=True)

                # Label yang tidak ketemu
                if result.not_found:
                    with st.expander(f"⚠️ {len(result.not_found)} label tidak ditemukan",
                                     expanded=False):
                        nf_df = pd.DataFrame(result.not_found)
                        st.dataframe(nf_df, use_container_width=True, hide_index=True)
                        st.caption(
                            "Kemungkinan penyebab: kode/NUP tidak ada di master PDF, "
                            "atau perbedaan format kode barang. Periksa manual di DBR."
                        )


# ============================================================
# FOOTER
# ============================================================
st.markdown(
    """
    <div style="text-align:center; color:#a0aec0; font-size:0.8rem;
                margin-top:2.5rem; padding-top:1rem; border-top:1px solid #edf2f7;">
      BBPPMPV Pertanian · Cianjur · Sistem Label BMN v2.0
    </div>
    """,
    unsafe_allow_html=True,
)
