"""
Sistem Cetak Label DBR BBPPMPV Pertanian (v5)
==============================================
Arsitektur CROP-BASED:
  - Master PDF → katalog (kode,NUP,tahun) + bbox untuk crop
  - DBR Excel  → filter per ruangan
  - Output     → label di-CROP utuh dari master (logo,QR,teks semua ada)
  - Item DBR yang TIDAK ADA di master → TIDAK dicetak (hanya ditampilkan di tabel)
"""
import io, re, zipfile, math
from dataclasses import dataclass, field
from datetime import datetime
import fitz
import pandas as pd
import streamlit as st

CURRENT_YEAR = datetime.now().year

st.set_page_config(page_title="Sistem Cetak Label DBR", page_icon="🏷️",
                   layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.main .block-container{padding-top:1rem;max-width:1300px}
.app-header{background:linear-gradient(135deg,#e65100,#ff8f00);color:#fff;padding:1.2rem 2rem;border-radius:14px;margin-bottom:1.2rem;box-shadow:0 6px 18px rgba(230,81,0,.22)}
.app-header .title{font-size:1.45rem;font-weight:700}
.app-header .sub{font-size:.88rem;opacity:.92;margin-top:.2rem}
.step-card{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:1rem 1.2rem;margin-bottom:.7rem;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.step-card.active{border-left:4px solid #e65100}.step-card.done{border-left:4px solid #2e7d32;background:#f4faf5}.step-card.locked{opacity:.5;border-left:4px solid #cbd5e0}
.step-header{display:flex;align-items:center;gap:.6rem;font-size:1.05rem;font-weight:600;color:#1a202c;margin-bottom:.15rem}
.step-badge{background:#e65100;color:#fff;width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:.8rem;font-weight:700}
.step-badge.done{background:#2e7d32}
.metric-big{background:linear-gradient(135deg,#e65100,#ff8f00);color:#fff;padding:1rem;border-radius:12px;text-align:center}
.metric-big .num{font-size:2.2rem;font-weight:700;line-height:1}.metric-big .lbl{font-size:.8rem;opacity:.9;margin-top:.25rem;text-transform:uppercase;letter-spacing:.4px}
section[data-testid="stSidebar"]{background:#fdf5ef}
.cond-card{padding:.7rem;border-radius:8px;text-align:center;margin-bottom:.4rem}
</style>""", unsafe_allow_html=True)

# ============================================================
@dataclass
class LabelInfo:
    kode_barang: str; nup: int; tahun: str
    source_pdf: str; page_index: int; bbox: tuple
    nama_barang: str = ""; doc_id: int = 0

@dataclass
class DBRItem:
    nup: int; kode_barang: str; nama_barang: str
    merk: str; tahun: str; keterangan: str
    in_master: bool = False

@dataclass
class RuanganResult:
    ruangan: str; sheet_name: str
    items: list = field(default_factory=list)
    matched_labels: list = field(default_factory=list)

# ============================================================
KODE_RE = re.compile(r"\b(\d{10})\b")
NUP_RE = re.compile(r"NUP\s*:?\s*(\d+)", re.IGNORECASE)
TAHUN_RE = re.compile(r"KD\s*\.?\s*(\d{4})")

def parse_multi_number(value):
    if value is None: return []
    if isinstance(value,(int,float)):
        return [] if pd.isna(value) else [int(value)]
    s = str(value).strip().replace("'","").replace('"',"").replace("`","")
    if not s or s.lower()=="nan": return []
    result = set()
    for p in re.split(r"[,;/]", s):
        m = re.search(r"\d+", p.strip())
        if m:
            try: result.add(int(m.group()))
            except ValueError: pass
    return sorted(result)

def clean_kode(value):
    if value is None: return ""
    if isinstance(value,float):
        if pd.isna(value): return ""
        try: return str(int(value))
        except: return ""
    if isinstance(value,int): return str(value)
    s = str(value).strip().replace("'","").replace(" ","")
    if s.endswith(".0"): s = s[:-2]
    try:
        if "e" in s.lower(): s = str(int(float(s)))
    except: pass
    m = re.search(r"\d{6,}", s)
    return m.group() if m else s

def classify_condition(text):
    if pd.isna(text) or not str(text).strip(): return "TIDAK DIKETAHUI"
    t = re.sub(r"\s+"," ",str(text).strip().upper())
    if t in ("BAIK","B","BK","GOOD"): return "BAIK"
    if any(k in t for k in ("RUSAK RINGAN","RR","RUSAK R")): return "RUSAK RINGAN"
    if any(k in t for k in ("RUSAK BERAT","RB")): return "RUSAK"
    if t in ("RUSAK","R"): return "RUSAK"
    if "KURANG" in t: return "RUSAK RINGAN"
    if "RUSAK" in t and "RINGAN" in t: return "RUSAK RINGAN"
    if "RUSAK" in t: return "RUSAK"
    if "BAIK" in t: return "BAIK"
    return "BAIK"

COND_DISPLAY = {
    "BAIK":("BAIK","#2e7d32","#e8f5e9"),
    "RUSAK RINGAN":("PERLU PERBAIKAN","#f57f17","#fff8e1"),
    "RUSAK":("PERLU DIGANTI/DIHAPUS","#c62828","#ffebee"),
    "TIDAK DIKETAHUI":("TIDAK DIKETAHUI","#78909c","#eceff1"),
}

def priority_score(kondisi, tahun):
    w = {"BAIK":1,"RUSAK RINGAN":5,"RUSAK":10,"TIDAK DIKETAHUI":2}
    try: age = max(0, CURRENT_YEAR - int(tahun))
    except: age = 0
    return w.get(kondisi,2) * (1 + age/10)

def priority_label(kondisi, tahun):
    try: age = CURRENT_YEAR - int(tahun)
    except: age = 0
    if kondisi=="RUSAK": return "🔴 SEGERA GANTI"
    if kondisi=="RUSAK RINGAN" and age>10: return "🟠 JADWALKAN PENGGANTIAN"
    if kondisi=="RUSAK RINGAN": return "🟡 PERBAIKI"
    if kondisi=="BAIK" and age>20: return "🔵 PANTAU (Usang)"
    return "🟢 AMAN"

# ============================================================
# MASTER PDF EXTRACTION — crop-based, grid detection
# ============================================================
def extract_labels_from_pdf(pdf_bytes, source_name, doc_id):
    """Extract label catalog WITH bbox for cropping. Returns (labels, fitz.Document)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    labels = []
    for pi in range(doc.page_count):
        page = doc[pi]
        pw, ph = page.rect.width, page.rect.height
        spans = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type",0)!=0: continue
            for ln in b.get("lines",[]):
                for sp in ln.get("spans",[]):
                    t = sp.get("text","").strip()
                    if t: spans.append((t, tuple(sp["bbox"])))
        if not spans: continue

        kem_ys = sorted(set(round(bb[1],1) for t,bb in spans if "KEMENTERIAN" in t.upper()))
        if not kem_ys: continue
        row_tops = []
        for y in kem_ys:
            if not row_tops or (y-row_tops[-1])>8: row_tops.append(y)
        if len(row_tops)>=2:
            diffs = [row_tops[i+1]-row_tops[i] for i in range(len(row_tops)-1)]
            row_h = sorted(diffs)[len(diffs)//2]  # median
        else:
            row_h = 119

        PAD_TOP = 16  # logo space above KEMENTERIAN text
        col_split = pw / 2

        for rt in row_tops:
            cell_y0 = max(0, rt - PAD_TOP)
            # Extend bottom to capture full label border (row_h only covers
            # to next KEMENTERIAN, but label border extends further)
            cell_y1 = min(ph, rt + row_h - 4)
            for col in (0,1):
                cx0 = 0 if col==0 else col_split
                cx1 = col_split if col==0 else pw
                bbox = (cx0, cell_y0, cx1, cell_y1)

                # Search text with extended Y range to catch bottom merk text
                cell_spans = [(bb[1],t) for t,bb in spans
                              if cx0<=(bb[0]+bb[2])/2<=cx1 and cell_y0-5<=(bb[1]+bb[3])/2<=cell_y1+10]
                if not cell_spans: continue
                cell_spans.sort()
                txt = " ".join(t for _,t in cell_spans)

                mn = NUP_RE.search(txt)
                mk = KODE_RE.search(txt)
                mt = TAHUN_RE.search(txt)
                if not (mn and mk and mt): continue
                try: nup_val = int(mn.group(1))
                except: continue

                # Extract nama barang
                kode = mk.group(1)
                nama = ""
                kode_y = None
                for y,t in cell_spans:
                    if kode in t: kode_y = y; break
                if kode_y:
                    for y,t in cell_spans:
                        if y>kode_y+3 and y<kode_y+25 and kode not in t and "NUP" not in t.upper():
                            nama = t; break

                labels.append(LabelInfo(
                    kode_barang=kode, nup=nup_val, tahun=mt.group(1),
                    source_pdf=source_name, page_index=pi, bbox=bbox,
                    nama_barang=nama, doc_id=doc_id))
    return labels, doc

# ============================================================
# DBR DETECTION
# ============================================================
def detect_dbr(xl_file, sheet_name):
    try: raw = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, dtype=object)
    except: return None, sheet_name
    if raw.empty or len(raw)<3: return None, sheet_name
    nc = len(raw.columns)

    kode_col, first_row = None, None
    for i in range(min(len(raw),200)):
        for j in range(min(nc,20)):
            v = raw.iloc[i,j]
            if pd.isna(v): continue
            if re.match(r"^\d{7,10}$", clean_kode(v)):
                first_row, kode_col = i, j; break
        if first_row is not None: break
    if first_row is None: return None, sheet_name

    col_map = {"kode": kode_col}
    for sr in range(first_row-1, max(first_row-10,-1), -1):
        if sr<0: break
        vals = [str(x).strip().lower() for x in raw.iloc[sr].values if pd.notna(x)]
        joined = " ".join(vals)
        if sum(1 for v in vals if re.match(r"^\d{1,2}$",v))>=4: continue
        if not any(k in joined for k in ["nama barang","kode barang","keterangan","merk","pendaftaran","tahun"]): continue
        for j in range(nc):
            v = raw.iloc[sr,j]
            if pd.isna(v): continue
            t = str(v).strip().lower()
            if ("pendaftaran" in t and "urut" in t) or t=="nup": col_map["nup"]=j
            elif "nama barang" in t: col_map["nama"]=j
            elif "merk" in t or "type" in t: col_map["merk"]=j
            elif "kode barang" in t: col_map["kode"]=j
            elif "tahun" in t: col_map["tahun"]=j
            elif "keterangan" in t: col_map["ket"]=j
        break

    if "nup" not in col_map:
        for off in [kode_col-3, kode_col-2, kode_col-1]:
            if off>=0:
                tv = raw.iloc[first_row, off]
                if parse_multi_number(tv): col_map["nup"]=off; break
        col_map.setdefault("nup", max(0,kode_col-3))
    col_map.setdefault("nama", max(0,kode_col-2))
    col_map.setdefault("merk", max(0,kode_col-1))
    col_map.setdefault("tahun", min(nc-1,kode_col+1))
    col_map.setdefault("ket", min(nc-1,kode_col+3))

    ruangan = sheet_name
    for i in range(first_row):
        for j in range(nc):
            v = raw.iloc[i,j]
            if pd.notna(v) and "ruangan" in str(v).lower():
                txt = str(v).strip()
                if ":" in txt:
                    c = txt.split(":",1)[1].strip()
                    if c: ruangan = c
                elif j+1<nc and pd.notna(raw.iloc[i,j+1]):
                    ruangan = str(raw.iloc[i,j+1]).strip().lstrip(": ")
                break

    items = []
    for i in range(first_row, len(raw)):
        kv = raw.iloc[i, col_map["kode"]] if col_map["kode"]<nc else None
        if pd.isna(kv): continue
        kode = clean_kode(kv)
        if not re.match(r"^\d{6,}$", kode): continue
        def g(k):
            idx=col_map.get(k)
            if idx is None or idx>=nc: return None
            return raw.iloc[i,idx]
        nup_list = parse_multi_number(g("nup"))
        if not nup_list: continue
        tahun_list = parse_multi_number(g("tahun"))
        nama = str(g("nama") or "").strip()
        merk = str(g("merk") or "").strip()
        ket = str(g("ket") or "").strip() if pd.notna(g("ket")) else ""
        for nup in nup_list:
            if len(tahun_list)==len(nup_list):
                thn = str(tahun_list[nup_list.index(nup)])
            elif tahun_list: thn = str(tahun_list[0])
            else: thn = ""
            items.append(DBRItem(nup=nup, kode_barang=kode, nama_barang=nama,
                                 merk=merk, tahun=thn, keterangan=ket))
    if not items: return None, ruangan
    return items, ruangan

# ============================================================
# PDF OUTPUT — CROP from master (vector-preserving, utuh)
# ============================================================
def build_output_pdf(ruangan, matched_labels, docs_by_id):
    """Print pages from master PDF 1:1 (no scaling). Only pages with matched labels."""
    if not matched_labels: return b""
    
    # Identify pages containing matched labels
    pages_to_print = sorted(set(lbl.page_index for lbl in matched_labels))
    src_doc = docs_by_id.get(matched_labels[0].doc_id)
    if src_doc is None: return b""
    
    doc = fitz.open()
    
    for idx, src_pi in enumerate(pages_to_print):
        if src_pi >= src_doc.page_count: continue
        
        # Copy page 1:1 from master (exact same size, no scaling)
        src_page = src_doc[src_pi]
        new_page = doc.new_page(width=src_page.rect.width, 
                                height=src_page.rect.height)
        
        # Draw master content verbatim
        new_page.show_pdf_page(src_page.rect, src_doc, src_pi)
        
        # Overlay: header (ruangan + page number)
        header_y = 8
        header_x = 20
        new_page.insert_text(fitz.Point(header_x, header_y),
                             f"DAFTAR LABEL BARANG — {ruangan.upper()}",
                             fontsize=8, fontname="hebo", color=(0,0,0))
        
        total_pages = len(pages_to_print)
        page_num_x = src_page.rect.width - 80
        new_page.insert_text(fitz.Point(page_num_x, header_y),
                             f"Hal {idx+1}/{total_pages}",
                             fontsize=7, fontname="helv", color=(0,0,0))
    
    buf = io.BytesIO()
    doc.save(buf, deflate=True)
    doc.close()
    return buf.getvalue()

# ============================================================
# SESSION STATE
# ============================================================
def init_state():
    for k,v in {"catalog":[],"docs_by_id":{},"master_ready":False,"master_meta":[],
                 "dbr_processed":False,"ruangan_results":{},"generated_pdfs":{},
                 "failed_sheets":[]}.items():
        if k not in st.session_state: st.session_state[k]=v
init_state()

# ============================================================
# HEADER
# ============================================================
st.markdown("""<div class="app-header">
<div class="title">🏷️ Sistem Cetak Label DBR BBPPMPV Pertanian</div>
<div class="sub">Filter &amp; cetak label BMN per ruangan · Label di-crop utuh dari Master PDF</div>
</div>""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("### 📋 Progres")
    s1="done" if st.session_state.master_ready else "active"
    s2="done" if st.session_state.dbr_processed else ("active" if st.session_state.master_ready else "")
    s3="done" if st.session_state.generated_pdfs else ("active" if st.session_state.dbr_processed else "")
    for i,(lb,sts) in enumerate([("Master label","active" if not st.session_state.master_ready else "done"),
                                   ("Proses DBR",s2),("Hasil & unduh",s3)],1):
        ico="✅" if sts=="done" else ("🔵" if sts=="active" else "⚪")
        st.markdown(f'<div style="padding:.4rem .6rem;border-radius:8px;margin-bottom:.3rem;font-size:.86rem">{ico} Langkah {i}: {lb}</div>', unsafe_allow_html=True)
    st.divider()
    if st.session_state.master_ready:
        st.metric("Label Master", f"{len(st.session_state.catalog):,}")
    st.divider()
    if st.button("🔄 Reset Semua Data", use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()

# ============================================================
# LANGKAH 1
# ============================================================
s1c = "done" if st.session_state.master_ready else "active"
st.markdown(f'<div class="step-card {s1c}"><div class="step-header"><span class="step-badge {"done" if st.session_state.master_ready else ""}">1</span>Master Label (PDF)</div><div style="color:#4a5568;font-size:.88rem">Upload PDF label BMN. Label akan di-<b>crop utuh</b> (logo, QR, teks semua terjaga).</div></div>', unsafe_allow_html=True)

cu,cm = st.columns([2,1])
with cu:
    upl_pdfs = st.file_uploader("PDF Label", type=["pdf"], accept_multiple_files=True, key="pdf_up", label_visibility="collapsed")
    if upl_pdfs: st.caption(f"📎 {len(upl_pdfs)} file")
    if st.button("🎯 JADIKAN DATA UTAMA", type="primary", disabled=not upl_pdfs, use_container_width=True):
        cat, docs, meta = [], {}, []
        bar = st.progress(0)
        for i,f in enumerate(upl_pdfs):
            bar.progress((i+1)/len(upl_pdfs)*.95, text=f"Ekstrak: {f.name}")
            try:
                b = f.read()
                ls, doc = extract_labels_from_pdf(b, f.name, doc_id=i)
                cat.extend(ls); docs[i] = doc
                meta.append((f.name, len(b)//1024, len(ls)))
            except Exception as e: st.error(f"❌ {f.name}: {e}")
        st.session_state.catalog = cat
        st.session_state.docs_by_id = docs
        st.session_state.master_ready = True
        st.session_state.master_meta = meta
        st.session_state.dbr_processed = False
        st.session_state.ruangan_results = {}
        st.session_state.generated_pdfs = {}
        bar.progress(1.0, text="Selesai!")
        st.rerun()

with cm:
    if st.session_state.master_ready:
        st.markdown(f'<div class="metric-big"><div class="num">{len(st.session_state.catalog):,}</div><div class="lbl">Label Katalog</div></div>', unsafe_allow_html=True)

if st.session_state.master_ready and st.session_state.catalog:
    with st.expander("📊 Detail Katalog Master"):
        cat = st.session_state.catalog
        cat_df = pd.DataFrame([{"Kode":l.kode_barang,"NUP":l.nup,"Tahun":l.tahun,"Nama":l.nama_barang,"File":l.source_pdf} for l in cat])
        grp = (cat_df.groupby(["Kode","Tahun"]).agg(Nama=("Nama","first"),Jml=("NUP","count"),NUPs=("NUP",lambda x:", ".join(str(v) for v in sorted(x)))).reset_index().sort_values(["Kode","Tahun"]))
        grp.columns = ["Kode Barang","Tahun","Nama Barang","Jumlah NUP","Daftar NUP"]
        grp["Status"] = grp["Tahun"].apply(lambda t: "⚠️ USANG" if int(t)<2000 else "✅ Aktif")
        st.dataframe(grp, use_container_width=True, hide_index=True)
        usang = grp[grp["Status"].str.contains("USANG")]
        if not usang.empty:
            st.warning(f"⚠️ {len(usang)} kelompok tahun<2000 ({usang['Jumlah NUP'].sum()} unit) — prioritas dihapus.")
        tc = cat_df["Tahun"].value_counts().sort_index().reset_index()
        tc.columns = ["Tahun","Jumlah"]
        st.bar_chart(tc, x="Tahun", y="Jumlah", color="#e65100")

# ============================================================
# LANGKAH 2
# ============================================================
st.markdown("<div style='height:.3rem'></div>", unsafe_allow_html=True)
s2_lock = not st.session_state.master_ready
s2c = "locked" if s2_lock else ("done" if st.session_state.dbr_processed else "active")
st.markdown(f'<div class="step-card {s2c}"><div class="step-header"><span class="step-badge {"done" if st.session_state.dbr_processed else ""}">2</span>Daftar Barang Ruangan (DBR)</div><div style="color:#4a5568;font-size:.88rem">Upload Excel DBR. Setiap <b>sheet = ruangan</b>. Hanya barang yang <b>ada di Master</b> yang akan dicetak.</div></div>', unsafe_allow_html=True)

if s2_lock:
    st.info("🔒 Selesaikan Langkah 1.")
else:
    upl_dbr = st.file_uploader("Excel DBR", type=["xlsx","xls"], key="dbr_up", label_visibility="collapsed")
    if upl_dbr:
        try:
            xb = upl_dbr.read()
            xl = pd.ExcelFile(io.BytesIO(xb))
            SKIP = ["master aset","master"]
            default_sel = [s for s in xl.sheet_names if s.lower().strip() not in SKIP]
            sel = st.multiselect("Sheet ruangan", xl.sheet_names, default=default_sel)

            if st.button("⚙️ PROSES", type="primary", disabled=not sel, use_container_width=True):
                master_idx = {}
                for l in st.session_state.catalog:
                    master_idx[(l.kode_barang, l.nup)] = l

                results, pdfs, failed = {}, {}, []
                bar = st.progress(0)
                for idx,sh in enumerate(sel):
                    bar.progress((idx+1)/len(sel)*.9, text=f"Proses: {sh}")
                    if sh.lower().strip() in SKIP: continue
                    items, ruangan = detect_dbr(io.BytesIO(xb), sh)
                    if items is None:
                        results[sh] = RuanganResult(ruangan=ruangan, sheet_name=sh)
                        failed.append(sh); continue

                    matched_labels = []
                    for it in items:
                        key = (it.kode_barang, it.nup)
                        if key in master_idx:
                            it.in_master = True
                            matched_labels.append(master_idx[key])
                        # else: in_master stays False, NOT included in print

                    rr = RuanganResult(ruangan=ruangan, sheet_name=sh,
                                       items=items, matched_labels=matched_labels)
                    results[sh] = rr
                    if matched_labels:
                        pdf_bytes = build_output_pdf(ruangan, matched_labels,
                                                     st.session_state.docs_by_id)
                        if pdf_bytes: pdfs[sh] = pdf_bytes

                st.session_state.ruangan_results = results
                st.session_state.generated_pdfs = pdfs
                st.session_state.dbr_processed = True
                st.session_state.failed_sheets = failed
                bar.progress(1.0, text=f"Selesai! {len(pdfs)} PDF.")
                st.rerun()
        except Exception as e: st.error(f"❌ {e}")

# ============================================================
# LANGKAH 3
# ============================================================
st.markdown("<div style='height:.3rem'></div>", unsafe_allow_html=True)
s3_lock = not st.session_state.dbr_processed
s3c = "locked" if s3_lock else "done"
st.markdown(f'<div class="step-card {s3c}"><div class="step-header"><span class="step-badge">3</span>Hasil Filter, Analisis &amp; Unduh</div><div style="color:#4a5568;font-size:.88rem">Label di-crop utuh dari Master PDF. Item tanpa label di master <b>tidak dicetak</b>.</div></div>', unsafe_allow_html=True)

if s3_lock:
    st.info("🔒 Selesaikan Langkah 2.")
else:
    results = st.session_state.ruangan_results
    pdfs = st.session_state.generated_pdfs
    if not results:
        st.warning("Tidak ada data.")
    else:
        failed = st.session_state.get("failed_sheets",[])
        if failed:
            st.error(f"⚠️ {len(failed)} sheet gagal deteksi: {', '.join(failed)}")

        total_items = sum(len(r.items) for r in results.values())
        total_matched = sum(len(r.matched_labels) for r in results.values())
        total_not = total_items - total_matched
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Ruangan", len(results))
        c2.metric("Total Barang (DBR)", total_items)
        c3.metric("Ada Label di Master", total_matched)
        c4.metric("Tidak Ada Label", total_not, delta=str(total_not) if total_not else None, delta_color="inverse")

        if len(pdfs)>1:
            zb = io.BytesIO()
            with zipfile.ZipFile(zb,"w",zipfile.ZIP_DEFLATED) as zf:
                for nm,pb in pdfs.items():
                    safe = re.sub(r"[^\w\-. ]","_",nm).strip()
                    zf.writestr(f"Label_{safe}.pdf", pb)
            st.download_button("📦 Unduh SEMUA (ZIP)", data=zb.getvalue(),
                               file_name="Label_Semua_Ruangan.zip", mime="application/zip",
                               type="primary", use_container_width=True)

        # Analisis kondisi
        st.markdown("---")
        st.markdown("### 📊 Analisis Kondisi Barang")
        all_data = []
        for sh,rr in results.items():
            for it in rr.items:
                k = classify_condition(it.keterangan)
                all_data.append({"ruangan":rr.ruangan,"sheet":sh,"kode_barang":it.kode_barang,
                    "nup":it.nup,"nama_barang":it.nama_barang,"merk":it.merk,
                    "tahun":it.tahun,"kondisi":k,"in_master":it.in_master,
                    "skor":round(priority_score(k,it.tahun),1),
                    "rekomendasi":priority_label(k,it.tahun)})

        if all_data:
            adf = pd.DataFrame(all_data)
            ccounts = adf["kondisi"].value_counts()
            cc = st.columns(min(len(ccounts),4))
            for ci,(kond,jml) in enumerate(ccounts.items()):
                d = COND_DISPLAY.get(kond,("?","#999","#eee"))
                with cc[ci%len(cc)]:
                    st.markdown(f'<div class="cond-card" style="background:{d[2]};border-left:4px solid {d[1]}"><div style="font-size:1.8rem;font-weight:700;color:{d[1]}">{jml}</div><div style="font-size:.78rem;color:{d[1]};font-weight:600">{d[0]}</div></div>', unsafe_allow_html=True)

            st.markdown("#### 🔴 Barang Perlu Diganti")
            ganti = adf[adf["kondisi"]=="RUSAK"]
            if not ganti.empty:
                st.dataframe(ganti[["ruangan","nama_barang","kode_barang","nup","tahun","merk","skor","rekomendasi"]].sort_values("skor",ascending=False).rename(columns={"ruangan":"Ruangan","nama_barang":"Nama","kode_barang":"Kode","tahun":"Tahun","merk":"Merk","skor":"Skor","rekomendasi":"Rekomendasi"}), use_container_width=True, hide_index=True)
            else: st.success("✅ Tidak ada barang RUSAK.")

            st.markdown("#### 🟡 Barang Perlu Perbaikan")
            perbaiki = adf[adf["kondisi"]=="RUSAK RINGAN"]
            if not perbaiki.empty:
                st.dataframe(perbaiki[["ruangan","nama_barang","kode_barang","nup","tahun","merk","skor","rekomendasi"]].sort_values("skor",ascending=False).rename(columns={"ruangan":"Ruangan","nama_barang":"Nama","kode_barang":"Kode","tahun":"Tahun","merk":"Merk","skor":"Skor","rekomendasi":"Rekomendasi"}), use_container_width=True, hide_index=True)
            else: st.success("✅ Tidak ada barang RUSAK RINGAN.")

            rusak_all = adf[adf["kondisi"].isin(["RUSAK","RUSAK RINGAN"])]
            if not rusak_all.empty:
                st.markdown("#### 🏚️ Ruangan dengan Barang Rusak Terbanyak")
                rpr = rusak_all.groupby("ruangan").size().reset_index(name="Jumlah Rusak").sort_values("Jumlah Rusak",ascending=False)
                st.bar_chart(rpr.set_index("ruangan")["Jumlah Rusak"], color="#c62828")

        # TABS SEMUA RUANGAN
        st.markdown("---")
        st.markdown("### 📑 Detail per Ruangan")
        all_sheets = list(results.keys())
        tabs = st.tabs([f"{results[s].ruangan}" for s in all_sheets])
        for tab, sh in zip(tabs, all_sheets):
            rr = results[sh]
            with tab:
                if not rr.items:
                    st.warning(f"⚠️ **{rr.ruangan}** — Tidak ada data terdeteksi.")
                    continue

                in_m = sum(1 for it in rr.items if it.in_master)
                not_m = len(rr.items) - in_m
                ca,cb,cc = st.columns(3)
                ca.metric("Total Barang (DBR)", len(rr.items))
                cb.metric("Ada Label (dicetak)", in_m)
                cc.metric("Tidak Ada Label", not_m, delta=str(not_m) if not_m else None, delta_color="inverse")

                if sh in pdfs:
                    safe = re.sub(r"[^\w\-. ]","_",sh).strip()
                    st.download_button(f"⬇️ Unduh PDF — {rr.ruangan}",
                                       data=pdfs[sh], file_name=f"Label_{safe}.pdf",
                                       mime="application/pdf", key=f"dl_{sh}", type="primary")

                # Tabel detail
                st.markdown("**Detail barang:**")
                dtbl = pd.DataFrame([{
                    "No":i+1, "Kode":it.kode_barang, "NUP":it.nup,
                    "Tahun":it.tahun, "Nama":it.nama_barang, "Merk":it.merk,
                    "Kondisi":classify_condition(it.keterangan),
                    "Label Master":"✅ Dicetak" if it.in_master else "❌ Tidak ada",
                } for i,it in enumerate(rr.items)])
                st.dataframe(dtbl, use_container_width=True, hide_index=True)

# Footer
st.markdown('<div style="text-align:center;color:#a0aec0;font-size:.78rem;margin-top:2rem;padding-top:.8rem;border-top:1px solid #edf2f7">BBPPMPV Pertanian · Cianjur · Sistem Label BMN v5.0</div>', unsafe_allow_html=True)
