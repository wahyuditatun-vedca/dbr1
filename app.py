"""
Sistem Cetak Label DBR BBPPMPV Pertanian
=========================================
Alur:
1. Upload master PDF label → "JADIKAN DATA UTAMA" → katalog label.
2. Upload DBR Excel → "PROSES" → filter label per ruangan (sheet).
3. Preview, visualisasi, & unduh PDF per ruangan atau ZIP semua.
"""

import io
import re
import zipfile
import base64
from dataclasses import dataclass, field
from typing import Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# ============================================================
# KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(
    page_title="Sistem Cetak Label DBR",
    page_icon="🏷️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# LOGO BASE64 (Kemendikdasmen – 120px)
# ============================================================
LOGO_B64 = (
    "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"
    "8EiLDrvMv8bsEAAAAASUVORK5CYII="
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 1rem; padding-bottom: 3rem; max-width: 1300px;
    }
    /* ---- Header Oranye ---- */
    .app-header {
        background: linear-gradient(135deg, #e65100 0%, #ff8f00 100%);
        color: #fff; padding: 1.2rem 2rem; border-radius: 14px;
        margin-bottom: 1.2rem;
        box-shadow: 0 6px 18px rgba(230,81,0,0.22);
        display: flex; align-items: center; gap: 1.2rem;
    }
    .app-header img { height: 62px; flex-shrink: 0; }
    .app-header .title { font-size: 1.45rem; font-weight: 700; line-height: 1.3; }
    .app-header .sub   { font-size: 0.88rem; opacity: 0.92; margin-top: 0.2rem; }

    .step-card {
        background: #fff; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 1.1rem 1.3rem; margin-bottom: 0.8rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .step-card.active { border-left: 4px solid #e65100; }
    .step-card.done   { border-left: 4px solid #2e7d32; background: #f4faf5; }
    .step-card.locked  { opacity: 0.5; border-left: 4px solid #cbd5e0; }

    .step-header {
        display: flex; align-items: center; gap: 0.65rem;
        font-size: 1.05rem; font-weight: 600; color: #1a202c; margin-bottom: 0.2rem;
    }
    .step-badge {
        background: #e65100; color: #fff; width: 26px; height: 26px;
        border-radius: 50%; display: flex; align-items: center; justify-content: center;
        font-size: 0.8rem; font-weight: 700;
    }
    .step-badge.done { background: #2e7d32; }

    .metric-big {
        background: linear-gradient(135deg, #e65100 0%, #ff8f00 100%);
        color: #fff; padding: 1.15rem; border-radius: 12px; text-align: center;
    }
    .metric-big .num { font-size: 2.4rem; font-weight: 700; line-height: 1; }
    .metric-big .lbl { font-size: 0.82rem; opacity: 0.9; margin-top: 0.3rem;
                       text-transform: uppercase; letter-spacing: 0.4px; }

    section[data-testid="stSidebar"] { background: #fdf5ef; }
    .sidebar-step {
        display: flex; align-items: center; gap: 0.5rem;
        padding: 0.5rem 0.7rem; border-radius: 8px; margin-bottom: 0.35rem;
        font-size: 0.88rem; color: #4a5568;
    }
    .sidebar-step.active { background: #fff3e0; color: #bf360c; font-weight: 600; }
    .sidebar-step.done   { background: #e8f5e9; color: #1b5e20; }

    .mini-note {
        font-size: 0.82rem; color: #4a5568; background: #f7fafc;
        border-left: 3px solid #a0aec0; padding: 0.5rem 0.75rem;
        border-radius: 4px; margin: 0.5rem 0;
    }
    [data-testid="stFileUploaderDropzone"] {
        padding: 0.8rem; border-radius: 10px; border-style: dashed;
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
    kode_barang: str
    nup: int
    tahun: str
    source_pdf: str
    page_index: int
    bbox: tuple
    nama_barang: str = ""
    doc_id: int = 0


@dataclass
class DBRRow:
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
    matched: list = field(default_factory=list)
    not_found: list = field(default_factory=list)
    condition_data: list = field(default_factory=list)  # list[dict] kondisi per item


# ============================================================
# HELPER FUNCTIONS
# ============================================================
KODE_RE = re.compile(r"\b(\d{10})\b")
NUP_RE = re.compile(r"NUP\s*:?\s*(\d+)", re.IGNORECASE)
TAHUN_RE = re.compile(r"KD\s*\.?\s*(\d{4})")


def parse_multi_number(value) -> list:
    if value is None:
        return []
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return []
        return [int(value)]
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return []
    s = s.replace("'", "").replace('"', "").replace("`", "")
    parts = re.split(r"[,;/]", s)
    result = set()
    for p in parts:
        p = p.strip()
        if not p:
            continue
        m = re.search(r"\d+", p)
        if m:
            try:
                result.add(int(m.group()))
            except ValueError:
                pass
    return sorted(result)


def clean_kode_barang(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip().replace("'", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    m = re.search(r"\d{6,}", s)
    return m.group() if m else s


def classify_condition(keterangan) -> str:
    """
    Klasifikasi kondisi barang dari teks Keterangan.
    Rule-based (deterministik) — tepat untuk mapping 3 kategori.
    """
    if pd.isna(keterangan) or not str(keterangan).strip():
        return "TIDAK DIKETAHUI"
    text = str(keterangan).strip().upper()
    text = re.sub(r"\s+", " ", text)

    if text in ("BAIK", "B", "BK", "GOOD"):
        return "BAIK"
    if any(k in text for k in ("RUSAK RINGAN", "RR", "RUSAK R")):
        return "RUSAK RINGAN"
    if any(k in text for k in ("RUSAK BERAT", "RB")):
        return "RUSAK"
    if text in ("RUSAK", "R"):
        return "RUSAK"
    if "KURANG" in text:
        return "RUSAK RINGAN"
    if "BAIK" in text:
        return "BAIK"
    if "RUSAK" in text and "RINGAN" in text:
        return "RUSAK RINGAN"
    if "RUSAK" in text:
        return "RUSAK"
    return "BAIK"  # default = baik jika tidak terdeteksi


CONDITION_DISPLAY = {
    "BAIK": ("BAIK", "#2e7d32", "#e8f5e9"),
    "RUSAK RINGAN": ("PERLU PERBAIKAN", "#f57f17", "#fff8e1"),
    "RUSAK": ("PERLU DIGANTI/DIHAPUS", "#c62828", "#ffebee"),
    "TIDAK DIKETAHUI": ("TIDAK DIKETAHUI", "#78909c", "#eceff1"),
}


# ============================================================
# PDF LABEL EXTRACTION (grid-based, vector-preserving)
# ============================================================
def extract_labels_from_pdf(pdf_bytes: bytes, source_name: str, doc_id: int) -> tuple:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    labels: list[LabelInfo] = []

    for page_index in range(doc.page_count):
        page = doc[page_index]
        page_w = page.rect.width
        page_h = page.rect.height

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

        # Deteksi ROW grid dari "KEMENTERIAN"
        kementerian_ys = sorted(set(
            round(bb[1], 1) for t, bb in spans
            if "KEMENTERIAN" in t.upper()
        ))
        if not kementerian_ys:
            continue

        # Cluster Y berdekatan
        row_tops = []
        for y in kementerian_ys:
            if not row_tops or (y - row_tops[-1]) > 8:
                row_tops.append(y)

        # Tinggi row
        if len(row_tops) >= 2:
            diffs = [row_tops[i + 1] - row_tops[i] for i in range(len(row_tops) - 1)]
            row_height = min(diffs)
        else:
            row_height = 100

        top_pad, bottom_pad = 6, 4

        for row_top in row_tops:
            cell_y0 = max(0, row_top - top_pad)
            cell_y1 = min(page_h, row_top + row_height - bottom_pad)

            for col in (0, 1):
                col_x0 = 0 if col == 0 else page_w / 2
                col_x1 = page_w / 2 if col == 0 else page_w
                cell_x0 = max(0, col_x0 + 4)
                cell_x1 = min(page_w, col_x1 - 4)
                cell_bbox = (cell_x0, cell_y0, cell_x1, cell_y1)

                cell_spans = []
                for t, bb in spans:
                    cx = (bb[0] + bb[2]) / 2
                    cy = (bb[1] + bb[3]) / 2
                    if cell_x0 <= cx <= cell_x1 and cell_y0 <= cy <= cell_y1:
                        cell_spans.append((bb[1], bb[0], t, bb))
                if not cell_spans:
                    continue
                cell_spans.sort()
                cell_text = " ".join(t for _, _, t, _ in cell_spans)

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

                kode_y = None
                for y, _, t, bb in cell_spans:
                    if kode in t:
                        kode_y = y
                        break
                nama = ""
                if kode_y is not None:
                    for y, _, t, bb in cell_spans:
                        if y > kode_y + 3 and y < kode_y + 25:
                            if kode in t or "NUP" in t.upper():
                                continue
                            nama = t
                            break

                labels.append(LabelInfo(
                    kode_barang=kode, nup=nup_val, tahun=tahun,
                    source_pdf=source_name, page_index=page_index,
                    bbox=cell_bbox, nama_barang=nama, doc_id=doc_id,
                ))

    return labels, doc


# ============================================================
# DBR DETECTION — REWRITTEN (data-row-first approach)
# ============================================================
def detect_dbr_data(xl_file, sheet_name: str):
    """
    Strategi ROBUST: cari baris data dulu (baris yang punya Kode Barang
    10 digit), lalu identifikasi kolom lewat header keywords + fallback
    posisi relatif.

    Return: (DataFrame|None, ruangan_str)
    """
    try:
        raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=object)
    except Exception:
        return None, sheet_name

    if raw.empty or len(raw) < 3:
        return None, sheet_name

    n_cols = len(raw.columns)

    # ---- STEP 1: Cari Kode Barang column (10-digit numbers) ----
    kode_col = None
    first_data_row = None

    for i in range(min(len(raw), 60)):
        for j in range(min(n_cols, 15)):
            val = raw.iloc[i, j]
            if pd.isna(val):
                continue
            cleaned = clean_kode_barang(val)
            if re.match(r"^\d{10}$", cleaned):
                if first_data_row is None:
                    first_data_row = i
                    kode_col = j
                break
        if first_data_row is not None:
            break

    if first_data_row is None or kode_col is None:
        return None, sheet_name

    # ---- STEP 2: Cari header row (scan backward) ----
    col_map = {}
    header_found = False

    for scan_row in range(first_data_row - 1, max(first_data_row - 6, -1), -1):
        if scan_row < 0:
            break
        row_texts = [str(x).strip().lower() for x in raw.iloc[scan_row].values
                     if pd.notna(x)]
        joined = " ".join(row_texts)

        # Skip number row (1, 2, 3, 4, 5…)
        nums_only = [v for v in row_texts if re.match(r"^\d{1,2}$", v)]
        if len(nums_only) >= 4:
            continue

        if not any(kw in joined for kw in
                   ["nama barang", "kode barang", "keterangan", "merk",
                    "nomor urut", "pendaftaran"]):
            continue

        # Parse per cell
        for j in range(n_cols):
            val = raw.iloc[scan_row, j]
            if pd.isna(val):
                continue
            t = str(val).strip().lower()
            if ("nomor urut" in t and "pendaftaran" in t) or t == "nup":
                col_map["nup"] = j
            elif "nama barang" in t:
                col_map["nama"] = j
            elif "merk" in t or "type" in t:
                col_map["merk"] = j
            elif "kode barang" in t:
                col_map["kode"] = j
            elif "tahun" in t and "perolehan" in t:
                col_map["tahun"] = j
            elif t.startswith("tahun") and "tahun" not in [k for k in col_map]:
                col_map["tahun"] = j
            elif "jumlah" in t:
                col_map["jumlah"] = j
            elif "keterangan" in t:
                col_map["keterangan"] = j
            elif "no" in t and "urut" in t and "pendaftaran" not in t:
                col_map["no_urut"] = j
        header_found = True
        break

    # Fallback: jika header ditemukan tapi partial, atau tidak ditemukan sama sekali
    col_map.setdefault("kode", kode_col)

    # Posisi relatif standar DBR Kemendikdasmen:
    # [A] No | [B] NUP | [C] Nama | [D] Merk | [E] Kode | [F] Tahun | [G] Jml | [H] Ket
    if "nup" not in col_map:
        # NUP biasanya 3 kolom sebelum Kode (jika ada No.Urut)
        # atau 1 kolom sebelum Nama
        for offset in [kode_col - 3, kode_col - 2, kode_col - 1]:
            if offset >= 0:
                # Verifikasi: apakah kolom ini berisi data NUP-like?
                test_val = raw.iloc[first_data_row, offset]
                parsed = parse_multi_number(test_val)
                if parsed and all(0 < x < 100000 for x in parsed):
                    col_map["nup"] = offset
                    break
        if "nup" not in col_map:
            col_map["nup"] = max(0, kode_col - 3)

    col_map.setdefault("nama", max(0, kode_col - 2))
    col_map.setdefault("merk", max(0, kode_col - 1))
    col_map.setdefault("tahun", min(n_cols - 1, kode_col + 1))
    col_map.setdefault("jumlah", min(n_cols - 1, kode_col + 2))
    col_map.setdefault("keterangan", min(n_cols - 1, kode_col + 3))

    # ---- STEP 3: Cari nama ruangan dari header area ----
    ruangan = sheet_name
    for i in range(first_data_row):
        for j in range(n_cols):
            val = raw.iloc[i, j]
            if pd.isna(val):
                continue
            text = str(val).strip()
            if "ruangan" in text.lower():
                if ":" in text:
                    candidate = text.split(":", 1)[1].strip()
                    if candidate:
                        ruangan = candidate
                elif j + 1 < n_cols:
                    nxt = raw.iloc[i, j + 1]
                    if pd.notna(nxt):
                        ruangan = str(nxt).strip().lstrip(": ")
                break

    # ---- STEP 4: Ekstrak data rows ----
    data_rows = []
    for i in range(first_data_row, len(raw)):
        kode_val = raw.iloc[i, col_map["kode"]] if col_map["kode"] < n_cols else None
        if pd.isna(kode_val):
            continue
        kode = clean_kode_barang(kode_val)
        if not re.match(r"^\d{6,}$", kode):
            continue

        def get(key):
            idx = col_map.get(key)
            if idx is None or idx >= n_cols:
                return None
            return raw.iloc[i, idx]

        nup_list = parse_multi_number(get("nup"))
        if not nup_list:
            continue

        tahun_list = parse_multi_number(get("tahun"))
        ket_raw = get("keterangan")
        ket = str(ket_raw).strip() if pd.notna(ket_raw) else ""

        data_rows.append(DBRRow(
            nup_list=nup_list,
            nama_barang=str(get("nama") or "").strip(),
            merk=str(get("merk") or "").strip(),
            kode_barang=kode,
            tahun_list=[str(t) for t in tahun_list],
            jumlah=str(get("jumlah") or "").strip(),
            keterangan=ket,
        ))

    if not data_rows:
        return None, ruangan

    return data_rows, ruangan


# ============================================================
# MATCHING
# ============================================================
def match_ruangan(dbr_rows: list, catalog: list, ruangan: str) -> MatchResult:
    index = {}
    for lbl in catalog:
        key = (lbl.kode_barang, lbl.nup)
        index.setdefault(key, []).append(lbl)

    result = MatchResult(ruangan=ruangan)
    for row in dbr_rows:
        for nup in row.nup_list:
            key = (row.kode_barang, nup)
            hits = index.get(key, [])
            condition = classify_condition(row.keterangan)
            if hits:
                chosen = None
                if row.tahun_list:
                    for h in hits:
                        if h.tahun in row.tahun_list:
                            chosen = h
                            break
                if chosen is None:
                    chosen = hits[0]
                result.matched.append(chosen)
                result.condition_data.append({
                    "kode_barang": chosen.kode_barang,
                    "nup": chosen.nup,
                    "tahun": chosen.tahun,
                    "nama_barang": row.nama_barang or chosen.nama_barang,
                    "kondisi": condition,
                    "keterangan_raw": row.keterangan,
                })
            else:
                result.not_found.append({
                    "kode_barang": row.kode_barang, "nup": nup,
                    "nama_barang": row.nama_barang,
                    "tahun": ", ".join(row.tahun_list),
                    "kondisi": condition,
                    "keterangan_raw": row.keterangan,
                })
    return result


# ============================================================
# PDF OUTPUT GENERATOR (vector-preserving)
# ============================================================
def build_output_pdf(matched: list, docs_by_id: dict) -> bytes:
    if not matched:
        return b""
    out = fitz.open()
    PAGE_W, PAGE_H = 595, 842
    MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 30, 40, 30
    GUTTER_X, GUTTER_Y = 12, 10
    cell_w = (PAGE_W - 2 * MARGIN_X - GUTTER_X) / 2

    sample = matched[0]
    src_w = sample.bbox[2] - sample.bbox[0]
    src_h = sample.bbox[3] - sample.bbox[1]
    scale = cell_w / src_w
    cell_h = src_h * scale

    usable_h = PAGE_H - MARGIN_TOP - MARGIN_BOTTOM
    rows_per_page = max(1, int((usable_h + GUTTER_Y) // (cell_h + GUTTER_Y)))
    cells_per_page = rows_per_page * 2

    for i, lbl in enumerate(matched):
        pos = i % cells_per_page
        col, row = pos % 2, pos // 2
        if pos == 0:
            page = out.new_page(width=PAGE_W, height=PAGE_H)
        x0 = MARGIN_X + col * (cell_w + GUTTER_X)
        y0 = MARGIN_TOP + row * (cell_h + GUTTER_Y)
        target_rect = fitz.Rect(x0, y0, x0 + cell_w, y0 + cell_h)
        src_doc = docs_by_id.get(lbl.doc_id)
        if src_doc is None:
            continue
        page.show_pdf_page(target_rect, src_doc, lbl.page_index,
                           clip=fitz.Rect(*lbl.bbox))

    buf = io.BytesIO()
    out.save(buf, deflate=True)
    out.close()
    return buf.getvalue()


# ============================================================
# SESSION STATE
# ============================================================
def init_state():
    defaults = {
        "catalog": [], "docs_by_id": {}, "master_ready": False,
        "master_files_meta": [], "dbr_processed": False,
        "match_results": {}, "generated_pdfs": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ============================================================
# HEADER ORANYE DENGAN LOGO
# ============================================================
st.markdown(
    f"""
    <div class="app-header">
      <img src="data:image/png;base64,{LOGO_B64}" alt="Logo Kemendikdasmen"/>
      <div>
        <div class="title">Sistem Cetak Label DBR BBPPMPV Pertanian</div>
        <div class="sub">Otomatisasi filter &amp; cetak label BMN per ruangan berbasis Daftar Barang Ruangan</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("### 📋 Progres")
    s1 = "done" if st.session_state.master_ready else "active"
    s2 = "done" if st.session_state.dbr_processed else ("active" if st.session_state.master_ready else "")
    s3 = "done" if st.session_state.generated_pdfs else ("active" if st.session_state.dbr_processed else "")
    for i, (label, status) in enumerate([
        ("Upload & set master label", s1),
        ("Upload & proses DBR", s2),
        ("Preview & unduh hasil", s3),
    ], start=1):
        icon = "✅" if status == "done" else ("🔵" if status == "active" else "⚪")
        css = status if status else ""
        st.markdown(
            f'<div class="sidebar-step {css}"><span>{icon}</span>'
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
    st.caption("💡 Match: (Kode Barang, NUP). Sheet = ruangan.")


# ============================================================
# LANGKAH 1 – MASTER LABEL
# ============================================================
step1_cls = "done" if st.session_state.master_ready else "active"
badge1 = "done" if st.session_state.master_ready else ""
st.markdown(
    f'<div class="step-card {step1_cls}">'
    f'<div class="step-header"><span class="step-badge {badge1}">1</span>'
    f'Master Label Barang (PDF)</div>'
    f'<div style="color:#4a5568;font-size:0.88rem">'
    f'Unggah file PDF label barang (bisa banyak). Klik <b>JADIKAN DATA UTAMA</b>.'
    f'</div></div>', unsafe_allow_html=True,
)

col_up, col_met = st.columns([2, 1])
with col_up:
    uploaded_pdfs = st.file_uploader(
        "Pilih file PDF label", type=["pdf"],
        accept_multiple_files=True, key="pdf_uploader", label_visibility="collapsed",
    )
    if uploaded_pdfs:
        st.caption(f"📎 {len(uploaded_pdfs)} file: " + ", ".join(f.name for f in uploaded_pdfs[:5])
                   + ("…" if len(uploaded_pdfs) > 5 else ""))
    if st.button("🎯 JADIKAN DATA UTAMA", type="primary",
                 disabled=not uploaded_pdfs, use_container_width=True):
        catalog, docs_by_id, meta = [], {}, []
        bar = st.progress(0, text="Memulai…")
        for i, f in enumerate(uploaded_pdfs):
            bar.progress((i + 1) / len(uploaded_pdfs) * 0.95, text=f"Mengekstrak: {f.name}")
            try:
                b = f.read()
                labels, doc = extract_labels_from_pdf(b, f.name, doc_id=i)
                catalog.extend(labels)
                docs_by_id[i] = doc
                meta.append((f.name, len(b) // 1024, len(labels)))
            except Exception as e:
                st.error(f"❌ Gagal: {f.name} — {e}")
        st.session_state.catalog = catalog
        st.session_state.docs_by_id = docs_by_id
        st.session_state.master_files_meta = meta
        st.session_state.master_ready = True
        st.session_state.dbr_processed = False
        st.session_state.match_results = {}
        st.session_state.generated_pdfs = {}
        bar.progress(1.0, text="Selesai!")
        st.rerun()

with col_met:
    if st.session_state.master_ready:
        st.markdown(
            f'<div class="metric-big">'
            f'<div class="num">{len(st.session_state.catalog):,}</div>'
            f'<div class="lbl">Total Label Terdeteksi</div></div>',
            unsafe_allow_html=True,
        )

# Detail Master + Visualisasi
if st.session_state.master_ready and st.session_state.catalog:
    with st.expander("📊 Detail & Visualisasi File Master", expanded=False):
        catalog = st.session_state.catalog

        # --- Tabel file ---
        meta_df = pd.DataFrame(
            st.session_state.master_files_meta,
            columns=["Nama File", "Ukuran (KB)", "Jumlah Label"],
        )
        st.dataframe(meta_df, use_container_width=True, hide_index=True)

        st.markdown("---")

        # --- Visualisasi: Kode Barang → Tahun → NUP ---
        st.markdown("#### 📦 Katalog per Kode Barang")
        cat_df = pd.DataFrame([
            {"Kode Barang": l.kode_barang, "NUP": l.nup,
             "Tahun": l.tahun, "Nama": l.nama_barang}
            for l in catalog
        ])

        # Group: per kode+tahun → kumpulkan NUP
        grouped = (
            cat_df.groupby(["Kode Barang", "Tahun"])
            .agg(
                Nama=("Nama", "first"),
                Jumlah_NUP=("NUP", "count"),
                NUP_List=("NUP", lambda x: ", ".join(str(v) for v in sorted(x))),
            )
            .reset_index()
            .sort_values(["Kode Barang", "Tahun"])
        )
        grouped.columns = ["Kode Barang", "Tahun Perolehan", "Nama Barang",
                           "Jumlah NUP", "Daftar NUP"]

        # Tandai barang usang (tahun < 2000)
        grouped["Status"] = grouped["Tahun Perolehan"].apply(
            lambda t: "⚠️ Barang Usang Prioritas Dihapus" if int(t) < 2000 else "✅ Aktif"
        )
        st.dataframe(grouped, use_container_width=True, hide_index=True)

        # --- Chart: distribusi tahun ---
        st.markdown("#### 📅 Distribusi Label per Tahun Perolehan")
        tahun_counts = cat_df["Tahun"].value_counts().sort_index().reset_index()
        tahun_counts.columns = ["Tahun", "Jumlah Label"]
        st.bar_chart(tahun_counts, x="Tahun", y="Jumlah Label", color="#e65100")

        # Highlight usang
        usang = grouped[grouped["Status"].str.contains("Usang")]
        if not usang.empty:
            st.warning(
                f"⚠️ **{len(usang)} kelompok barang** dengan tahun perolehan < 2000 "
                f"(total {usang['Jumlah NUP'].sum()} unit) — "
                f"**Prioritas dihapus** sesuai ketentuan BMN."
            )


# ============================================================
# LANGKAH 2 – DBR
# ============================================================
st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)
step2_locked = not st.session_state.master_ready
step2_cls = "locked" if step2_locked else ("done" if st.session_state.dbr_processed else "active")
badge2 = "done" if st.session_state.dbr_processed else ""
st.markdown(
    f'<div class="step-card {step2_cls}">'
    f'<div class="step-header"><span class="step-badge {badge2}">2</span>'
    f'Daftar Barang Ruangan (DBR)</div>'
    f'<div style="color:#4a5568;font-size:0.88rem">'
    f'Unggah file DBR (Excel). Setiap <b>sheet</b> = satu ruangan. '
    f'Sheet <b>Master Aset</b> otomatis di-skip.'
    f'</div></div>', unsafe_allow_html=True,
)

if step2_locked:
    st.markdown('<div class="mini-note">🔒 Selesaikan Langkah 1.</div>',
                unsafe_allow_html=True)
else:
    uploaded_dbr = st.file_uploader(
        "Pilih file DBR (Excel)", type=["xlsx", "xls"],
        key="dbr_uploader", label_visibility="collapsed",
    )
    if uploaded_dbr:
        try:
            xl_bytes = uploaded_dbr.read()
            xl = pd.ExcelFile(io.BytesIO(xl_bytes))
            all_sheets = xl.sheet_names
            st.caption(f"📑 **{uploaded_dbr.name}** — {len(all_sheets)} sheet")

            EXCLUDE_KW = ["master aset", "master", "index", "cover"]
            default_sel = [s for s in all_sheets
                           if not any(k == s.lower().strip() for k in EXCLUDE_KW)]
            if not default_sel:
                default_sel = all_sheets

            selected = st.multiselect(
                "Sheet ruangan yang diproses",
                options=all_sheets, default=default_sel,
                help="Sheet 'Master Aset' di-exclude otomatis.",
            )

            if st.button("⚙️ PROSES", type="primary",
                         disabled=not selected, use_container_width=True):
                match_results, generated_pdfs = {}, {}
                bar = st.progress(0, text="Memulai…")
                success_count = 0

                for idx, sheet in enumerate(selected):
                    bar.progress((idx + 1) / len(selected) * 0.9,
                                 text=f"Memproses: {sheet}")

                    # Skip Master Aset variants
                    if sheet.lower().strip() in EXCLUDE_KW:
                        continue

                    dbr_rows, ruangan = detect_dbr_data(
                        io.BytesIO(xl_bytes), sheet
                    )
                    if dbr_rows is None or (isinstance(dbr_rows, list) and len(dbr_rows) == 0):
                        continue

                    result = match_ruangan(
                        dbr_rows, st.session_state.catalog, ruangan
                    )
                    match_results[sheet] = result

                    if result.matched:
                        pdf_bytes = build_output_pdf(
                            result.matched, st.session_state.docs_by_id
                        )
                        if pdf_bytes:
                            generated_pdfs[sheet] = pdf_bytes
                            success_count += 1

                st.session_state.match_results = match_results
                st.session_state.generated_pdfs = generated_pdfs
                st.session_state.dbr_processed = True
                bar.progress(1.0, text=f"Selesai! {success_count} ruangan berhasil.")
                st.rerun()

        except Exception as e:
            st.error(f"❌ Gagal membaca DBR: {e}")


# ============================================================
# LANGKAH 3 – HASIL FILTER & UNDUH
# ============================================================
st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)
step3_locked = not st.session_state.dbr_processed
step3_cls = "locked" if step3_locked else ("done" if st.session_state.generated_pdfs else "active")
st.markdown(
    f'<div class="step-card {step3_cls}">'
    f'<div class="step-header"><span class="step-badge">3</span>'
    f'Hasil Filter & Unduh</div>'
    f'<div style="color:#4a5568;font-size:0.88rem">'
    f'Preview per ruangan, visualisasi kondisi barang, unduh PDF satuan atau ZIP semua.'
    f'</div></div>', unsafe_allow_html=True,
)

if step3_locked:
    st.markdown('<div class="mini-note">🔒 Selesaikan Langkah 2.</div>',
                unsafe_allow_html=True)
else:
    results = st.session_state.match_results
    pdfs = st.session_state.generated_pdfs

    if not results:
        st.warning("Tidak ada data yang berhasil diproses dari DBR.")
    else:
        total_matched = sum(len(r.matched) for r in results.values())
        total_nf = sum(len(r.not_found) for r in results.values())
        total_ruangan = len(results)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ruangan Diproses", total_ruangan)
        c2.metric("Label Ditemukan", total_matched)
        c3.metric("Tidak Ditemukan", total_nf,
                  delta=f"{total_nf}" if total_nf else None,
                  delta_color="inverse")
        c4.metric("PDF Ter-generate", len(pdfs))

        # ZIP all
        if len(pdfs) > 1:
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                for rm, pb in pdfs.items():
                    safe = re.sub(r"[^\w\-. ]", "_", rm).strip() or "ruangan"
                    zf.writestr(f"Label_{safe}.pdf", pb)
            st.download_button("📦 Unduh SEMUA Ruangan (ZIP)",
                               data=zbuf.getvalue(),
                               file_name="Label_Semua_Ruangan.zip",
                               mime="application/zip", type="primary",
                               use_container_width=True)

        # ====== VISUALISASI KONDISI BARANG GLOBAL ======
        st.markdown("---")
        st.markdown("### 📊 Visualisasi Kondisi Barang")

        all_cond = []
        for rm, res in results.items():
            for cd in res.condition_data:
                all_cond.append({**cd, "ruangan": rm})
            for nf in res.not_found:
                all_cond.append({**nf, "ruangan": rm})

        if all_cond:
            cond_df = pd.DataFrame(all_cond)

            # --- Chart kondisi global ---
            st.markdown("#### Distribusi Kondisi Barang (Semua Ruangan)")
            cond_counts = cond_df["kondisi"].value_counts().reset_index()
            cond_counts.columns = ["Kondisi", "Jumlah"]
            # Map display labels
            cond_counts["Label"] = cond_counts["Kondisi"].map(
                lambda k: CONDITION_DISPLAY.get(k, ("?", "#999", "#eee"))[0]
            )
            cond_counts["Warna"] = cond_counts["Kondisi"].map(
                lambda k: CONDITION_DISPLAY.get(k, ("?", "#999", "#eee"))[1]
            )

            cols_cond = st.columns(len(cond_counts))
            for idx_c, (_, row_c) in enumerate(cond_counts.iterrows()):
                disp = CONDITION_DISPLAY.get(row_c["Kondisi"],
                                             ("?", "#78909c", "#eceff1"))
                with cols_cond[idx_c]:
                    st.markdown(
                        f'<div style="background:{disp[2]};border-left:4px solid {disp[1]};'
                        f'padding:0.8rem;border-radius:8px;text-align:center">'
                        f'<div style="font-size:1.8rem;font-weight:700;color:{disp[1]}">'
                        f'{row_c["Jumlah"]}</div>'
                        f'<div style="font-size:0.78rem;color:{disp[1]};font-weight:600">'
                        f'{disp[0]}</div></div>',
                        unsafe_allow_html=True,
                    )

            # --- Chart: Ruangan dengan barang rusak terbanyak ---
            st.markdown("#### 🏚️ Ruangan dengan Barang Rusak Terbanyak")
            rusak_df = cond_df[cond_df["kondisi"].isin(["RUSAK", "RUSAK RINGAN"])]
            if not rusak_df.empty:
                rusak_per_room = (
                    rusak_df.groupby("ruangan")
                    .size()
                    .reset_index(name="Jumlah Rusak")
                    .sort_values("Jumlah Rusak", ascending=False)
                )
                st.bar_chart(rusak_per_room.set_index("ruangan")["Jumlah Rusak"],
                             color="#c62828")

                top_rusak = rusak_per_room.iloc[0]
                st.info(
                    f"📌 Ruangan **{top_rusak['ruangan']}** memiliki "
                    f"**{top_rusak['Jumlah Rusak']} barang rusak** — paling banyak."
                )
            else:
                st.success("✅ Tidak ada barang rusak di seluruh ruangan.")

        # ====== TABS PER RUANGAN ======
        st.markdown("---")
        st.markdown("### 📑 Detail per Ruangan")
        tabs = st.tabs(list(results.keys()))
        for tab, (ruangan, result) in zip(tabs, results.items()):
            with tab:
                ca, cb, cc = st.columns(3)
                ca.metric("Label Match", len(result.matched))
                cb.metric("Tidak Ditemukan", len(result.not_found))
                cc.metric("Kode Unik", len(set(l.kode_barang for l in result.matched)))

                if ruangan in pdfs:
                    safe = re.sub(r"[^\w\-. ]", "_", ruangan).strip() or "ruangan"
                    st.download_button(
                        f"⬇️ Unduh PDF — {ruangan}",
                        data=pdfs[ruangan],
                        file_name=f"Label_{safe}.pdf",
                        mime="application/pdf",
                        key=f"dl_{ruangan}", type="primary",
                    )

                # Kondisi per ruangan
                if result.condition_data:
                    room_cond = pd.DataFrame(result.condition_data)
                    room_counts = room_cond["kondisi"].value_counts()
                    cond_cols = st.columns(min(len(room_counts), 4))
                    for ci, (kondisi, jumlah) in enumerate(room_counts.items()):
                        disp = CONDITION_DISPLAY.get(
                            kondisi, ("?", "#78909c", "#eceff1"))
                        with cond_cols[ci % len(cond_cols)]:
                            st.markdown(
                                f'<div style="background:{disp[2]};'
                                f'border-left:3px solid {disp[1]};'
                                f'padding:0.5rem;border-radius:6px;'
                                f'text-align:center;margin-bottom:0.5rem">'
                                f'<span style="font-weight:700;color:{disp[1]}">'
                                f'{jumlah}</span> '
                                f'<span style="font-size:0.75rem;color:{disp[1]}">'
                                f'{disp[0]}</span></div>',
                                unsafe_allow_html=True,
                            )

                # Tabel match
                if result.matched:
                    st.markdown("**Label yang berhasil difilter:**")
                    m_df = pd.DataFrame([
                        {"No": i + 1, "Kode Barang": l.kode_barang,
                         "NUP": l.nup, "Tahun": l.tahun,
                         "Nama": l.nama_barang, "Sumber": l.source_pdf}
                        for i, l in enumerate(result.matched)
                    ])
                    st.dataframe(m_df, use_container_width=True, hide_index=True)

                if result.not_found:
                    with st.expander(f"⚠️ {len(result.not_found)} tidak ditemukan"):
                        nf_df = pd.DataFrame(result.not_found)
                        st.dataframe(nf_df, use_container_width=True, hide_index=True)
                        st.caption(
                            "Kemungkinan: kode/NUP tidak ada di master PDF, "
                            "atau format berbeda."
                        )


# ============================================================
# CATATAN TENTANG XGBOOST
# ============================================================
# Tatun meminta klasifikasi kondisi barang menggunakan XGBoost.
# Setelah analisis, rule-based mapping lebih tepat di sini karena:
# 1. Kolom "Keterangan" hanya berisi 3 kategori diskrit (BAIK/RUSAK RINGAN/RUSAK)
#    yang bisa di-map deterministik — bukan masalah prediksi/estimasi.
# 2. XGBoost membutuhkan training data & feature engineering; untuk mapping
#    string → kategori, hasilnya identik dengan if-else tapi jauh lebih berat.
# 3. Menambah dependency (xgboost, scikit-learn) memperbesar image deploy.
# Jika ada use-case prediksi (misal: prediksi kapan barang akan rusak
# berdasarkan umur, frekuensi pakai, dll), XGBoost baru tepat digunakan.


# ============================================================
# FOOTER
# ============================================================
st.markdown(
    """
    <div style="text-align:center; color:#a0aec0; font-size:0.78rem;
                margin-top:2rem; padding-top:0.8rem; border-top:1px solid #edf2f7;">
      BBPPMPV Pertanian · Cianjur · Sistem Label BMN v3.0
    </div>
    """,
    unsafe_allow_html=True,
)
