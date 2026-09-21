"""
Sistem Cetak Label DBR - BBPPMPV Pertanian (v3 Final)
====================================================
Master Label PDF (SIMAK-BMN) + DBR Excel (Master_Aset_SIMAN) → Filtered Label PDF

Struktur Excel:
- Kolom D = Kode Barang
- Kolom E = NUP
- Kolom J = Lokasi Ruang

Aturan:
- Data utama: Master Label PDF (tidak dimodifikasi isi/bentuk)
- Filter: DBR Excel → group by Lokasi Ruang
- Matching key: (Kode Barang, NUP)
- Output: PDF label yang cocok, di-clip dari Master PDF
"""

import streamlit as st
import fitz  # PyMuPDF
import re
import io
import pandas as pd


# ================================================================
# 1. MASTER PDF PARSER
# ================================================================

def extract_labels_from_master(doc):
    """
    Scan setiap halaman Master PDF, cari pattern 'XXXXXXXXXX NUP: NN'.
    Tentukan posisi grid, simpan clip rectangle per label.

    Returns: dict { (kode_str, nup_int): {page_idx, clip_rect, kode, nup} }
    """
    label_map = {}

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        pw, ph = page.rect.width, page.rect.height

        blocks = page.get_text("blocks")

        # Cari pattern Kode+NUP
        found = []
        for b in blocks:
            if b[6] != 0:  # text block only
                continue
            for m in re.finditer(r'(\d{7,10})\s+NUP:\s*(\d+)', b[4]):
                found.append({
                    'kode': m.group(1),
                    'nup': int(m.group(2)),
                    'x0': b[0], 'y0': b[1],
                    'x1': b[2], 'y1': b[3],
                    'xc': (b[0] + b[2]) / 2,
                    'yc': (b[1] + b[3]) / 2,
                })

        if not found:
            continue

        # Sort by Y then X
        found.sort(key=lambda f: (f['y0'], f['xc']))

        # Cluster into rows (Y tolerance 25pt)
        rows = []
        cur = [found[0]]
        for i in range(1, len(found)):
            if abs(found[i]['y0'] - cur[-1]['y0']) < 25:
                cur.append(found[i])
            else:
                rows.append(sorted(cur, key=lambda f: f['xc']))
                cur = [found[i]]
        rows.append(sorted(cur, key=lambda f: f['xc']))

        n_rows = len(rows)
        y_centers = [sum(f['yc'] for f in r) / len(r) for r in rows]

        # Calculate row height
        if n_rows >= 2:
            spacings = [y_centers[i+1] - y_centers[i] for i in range(n_rows - 1)]
            avg_spacing = sum(spacings) / len(spacings)
            row_h = avg_spacing
        else:
            row_h = ph / 6

        # Row bounds
        row_bounds = []
        for i in range(n_rows):
            y_top = y_centers[i] - row_h * 0.55
            y_bot = y_centers[i] + row_h * 0.55
            y_top = max(0, y_top)
            y_bot = min(ph, y_bot)
            row_bounds.append((y_top, y_bot))

        x_mid = pw / 2

        # Assign clip rect to each label
        for ri, row in enumerate(rows):
            y_top, y_bot = row_bounds[ri]
            for label in row:
                if label['xc'] < x_mid:
                    clip = fitz.Rect(0, y_top, x_mid, y_bot)
                else:
                    clip = fitz.Rect(x_mid, y_top, pw, y_bot)

                key = (label['kode'], label['nup'])
                if key not in label_map:
                    label_map[key] = {
                        'page_idx': page_idx,
                        'clip_rect': clip,
                        'kode': label['kode'],
                        'nup': label['nup'],
                    }

    return label_map


# ================================================================
# 2. DBR EXCEL PARSER
# ================================================================

def parse_nup_values(val):
    """Parse NUP dari cell — support int, float, comma-separated string."""
    if val is None or pd.isna(val):
        return []
    
    s = str(val).strip().lstrip("'").rstrip(",").strip()
    if not s:
        return []
    
    # Try direct int/float first
    try:
        return [int(float(s))]
    except (ValueError, OverflowError):
        pass
    
    # Try comma/space separated
    parts = re.split(r'[,;\s]+', s)
    nups = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        try:
            nups.append(int(float(p)))
        except (ValueError, OverflowError):
            pass
    return nups


def parse_kode_barang(val):
    """Parse Kode Barang — must be 7-10 digits."""
    if val is None or pd.isna(val):
        return None
    
    s = str(val).strip()
    
    # Handle float format
    if '.' in s:
        try:
            s = str(int(float(s)))
        except (ValueError, OverflowError):
            s = s.split('.')[0]
    
    if re.match(r'^\d{7,10}$', s):
        return s
    return None


def parse_dbr_excel(excel_bytes):
    """
    Parse Excel DBR dengan struktur:
    - Kolom D (index 3) = Kode Barang
    - Kolom E (index 4) = NUP
    - Kolom J (index 9) = Lokasi Ruang
    
    Returns: dict { lokasi_ruang: [(kode, nup), ...] }
    """
    try:
        # Baca semua sheet, ambil kolom D, E, J saja
        df = pd.read_excel(
            io.BytesIO(excel_bytes),
            usecols=['Kode Barang', 'NUP', 'Lokasi Ruang'],
            dtype={'Kode Barang': str, 'NUP': str, 'Lokasi Ruang': str}
        )
    except Exception as e:
        st.error(f"Error membaca Excel: {e}")
        return {}

    groups = {}

    for _, row in df.iterrows():
        kode_raw = row.get('Kode Barang')
        nup_raw = row.get('NUP')
        lokasi_raw = row.get('Lokasi Ruang')

        kode = parse_kode_barang(kode_raw)
        if kode is None:
            continue

        nups = parse_nup_values(nup_raw)
        if not nups:
            continue

        lokasi = str(lokasi_raw).strip() if lokasi_raw else "Tidak Tercantum"

        if lokasi not in groups:
            groups[lokasi] = []

        for nup in nups:
            groups[lokasi].append((kode, nup))

    return groups


# ================================================================
# 3. PDF GENERATOR
# ================================================================

def generate_filtered_pdf(master_doc, label_map, filter_items):
    """
    Generate PDF output — clip label dari Master PDF.
    Returns: (pdf_bytes, matched_list, not_found_list)
    """
    matched = []
    not_found = []

    for kode, nup in filter_items:
        key = (kode, nup)
        if key in label_map:
            matched.append(label_map[key])
        else:
            not_found.append((kode, nup))

    if not matched:
        return None, matched, not_found

    # Calculate average label dimensions
    widths = [m['clip_rect'].width for m in matched]
    heights = [m['clip_rect'].height for m in matched]
    label_w = sum(widths) / len(widths)
    label_h = sum(heights) / len(heights)

    # Output page: 2 columns × 6 rows
    page_w = label_w * 2
    page_h = label_h * 6
    labels_per_page = 12

    output_doc = fitz.open()

    for start in range(0, len(matched), labels_per_page):
        batch = matched[start:start + labels_per_page]
        new_page = output_doc.new_page(width=page_w, height=page_h)

        for idx, label in enumerate(batch):
            row = idx // 2
            col = idx % 2

            target = fitz.Rect(
                col * label_w,
                row * label_h,
                (col + 1) * label_w,
                (row + 1) * label_h,
            )

            new_page.show_pdf_page(
                target,
                master_doc,
                label['page_idx'],
                clip=label['clip_rect'],
            )

    pdf_bytes = output_doc.tobytes()
    output_doc.close()
    return pdf_bytes, matched, not_found


# ================================================================
# 4. STREAMLIT UI
# ================================================================

def main():
    st.set_page_config(
        page_title="Sistem Cetak Label DBR",
        page_icon="🏷️",
        layout="wide",
    )

    st.title("Sistem Cetak Label DBR")
    st.caption("BBPPMPV Pertanian — Kemendikdasmen")

    st.markdown("---")

    # Upload section
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Master Label PDF")
        st.caption("File PDF Label Barang dari SIMAK-BMN")
        master_file = st.file_uploader(
            "Upload Master Label PDF",
            type=["pdf"],
            key="master",
        )
    with col_b:
        st.subheader("DBR Excel")
        st.caption("Master_Aset_SIMAN dengan kolom Lokasi Ruang")
        dbr_file = st.file_uploader(
            "Upload DBR Excel",
            type=["xlsx", "xls"],
            key="dbr",
        )

    if not master_file or not dbr_file:
        st.info("Upload kedua file (Master Label PDF + DBR Excel) untuk memulai.")
        st.markdown("---")
        st.markdown("**Format Excel yang diharapkan:**")
        st.code("Kolom D = Kode Barang\nKolom E = NUP\nKolom J = Lokasi Ruang")
        return

    # Process Master PDF
    master_bytes = master_file.read()
    master_doc = fitz.open(stream=master_bytes, filetype="pdf")

    with st.spinner("Mengekstrak label dari Master PDF..."):
        label_map = extract_labels_from_master(master_doc)

    # Process DBR Excel
    dbr_bytes = dbr_file.read()

    with st.spinner("Membaca DBR Excel (group by Lokasi Ruang)..."):
        groups = parse_dbr_excel(dbr_bytes)

    # Info section
    st.markdown("---")
    c1, c2 = st.columns(2)
    c1.metric("Label di Master PDF", len(label_map))
    c2.metric("Lokasi Ruang unik di DBR", len(groups))

    if not label_map:
        st.error(
            "❌ Tidak ada label terdeteksi di Master PDF. "
            "Pastikan PDF berisi label dengan format: Kode Barang + NUP"
        )
        return

    if not groups:
        st.error(
            "❌ Tidak ada data valid di DBR Excel. "
            "Pastikan kolom: D=Kode Barang, E=NUP, J=Lokasi Ruang"
        )
        return

    # Select location
    st.markdown("---")
    lokasi_names = sorted(groups.keys())
    selected_lokasi = st.selectbox("Pilih Lokasi Ruang:", lokasi_names)

    items = groups[selected_lokasi]

    # Preview
    st.write(f"**Total item (Kode + NUP) di lokasi ini:** {len(items)}")

    n_match = sum(1 for k, n in items if (k, n) in label_map)
    n_miss = len(items) - n_match

    col_x, col_y = st.columns(2)
    col_x.metric("✅ Ditemukan di Master", n_match)
    col_y.metric("❌ Tidak ditemukan", n_miss)

    with st.expander("Lihat detail item", expanded=False):
        for i, (kode, nup) in enumerate(items, 1):
            found = (kode, nup) in label_map
            icon = "✅" if found else "❌"
            st.text(f"{i:3d}. {icon}  Kode: {kode}  NUP: {nup}")

    # Generate button
    st.markdown("---")
    if st.button("🖨️ Cetak Label", type="primary", use_container_width=True):
        with st.spinner("Membuat PDF label..."):
            pdf_bytes, matched, not_found = generate_filtered_pdf(
                master_doc, label_map, items
            )

        if pdf_bytes:
            st.success(f"✅ {len(matched)} label berhasil digenerate")

            if not_found:
                with st.expander(
                    f"⚠️ {len(not_found)} item tidak ditemukan",
                    expanded=True,
                ):
                    for kode, nup in not_found:
                        st.text(f"  Kode: {kode}  NUP: {nup}")

            safe_name = re.sub(r'[^\w\s-]', '', selected_lokasi).strip().replace(' ', '_')
            st.download_button(
                label="📥 Download Label PDF",
                data=pdf_bytes,
                file_name=f"Label_{safe_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.error("❌ Tidak ada label cocok. Periksa data DBR dan Master PDF.")


if __name__ == "__main__":
    main()
