import io
import zipfile
from pathlib import PurePosixPath

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Gabung Excel dari ZIP", page_icon="📁", layout="wide")
st.title("📁 Gabung Excel dari ZIP")
st.write("Gabungkan file Excel atau CSV di dalam ZIP, periksa jumlah baris, lalu unduh hasilnya.")
st.info("Data diunggah untuk diproses oleh server aplikasi. Jangan unggah data rahasia ke layanan publik tanpa izin dan kontrol keamanan yang sesuai.")

MAX_ZIP_SIZE = 100 * 1024 * 1024
MAX_FILES = 200
MAX_UNCOMPRESSED = 350 * 1024 * 1024
EXTENSIONS = (".xlsx", ".xls", ".csv")


def clean_columns(columns):
    used = {}
    result = []
    for idx, name in enumerate(columns):
        base = str(name).strip() if pd.notna(name) else ""
        if not base or base.lower().startswith("unnamed:"):
            base = f"Kolom_{idx + 1}"
        used[base] = used.get(base, 0) + 1
        result.append(base if used[base] == 1 else f"{base}__{used[base]}")
    return result


def convert_csv_bytes(raw):
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, encoding=encoding, sep=None, engine="python")
        except UnicodeDecodeError:
            continue
    raise ValueError("Encoding CSV tidak didukung")


def parse_zip(uploaded, sheet_mode, requested_sheet, add_source):
    frames = []
    audit = []
    issues = []
    with zipfile.ZipFile(io.BytesIO(uploaded)) as z:
        infos = [i for i in z.infolist() if not i.is_dir() and not PurePosixPath(i.filename).name.startswith((".", "~$")) and i.filename.lower().endswith(EXTENSIONS)]
        if len(infos) > MAX_FILES:
            raise ValueError(f"Maksimal {MAX_FILES} file sumber dalam satu ZIP.")
        if sum(i.file_size for i in infos) > MAX_UNCOMPRESSED:
            raise ValueError("Jumlah ukuran file setelah ekstraksi terlalu besar (maksimal 350 MB).")
        if not infos:
            raise ValueError("ZIP tidak berisi Excel (.xlsx/.xls) atau CSV.")

        for info in infos:
            filename = info.filename
            try:
                raw = z.read(info)
                if filename.lower().endswith(".csv"):
                    sheet_frames = {"CSV": convert_csv_bytes(raw)}
                else:
                    xl = pd.ExcelFile(io.BytesIO(raw))
                    if sheet_mode == "Semua sheet":
                        names = xl.sheet_names
                    elif sheet_mode == "Nama sheet tertentu":
                        names = [requested_sheet] if requested_sheet in xl.sheet_names else []
                        if not names:
                            issues.append({"File": filename, "Sheet": requested_sheet, "Masalah": "Nama sheet tidak ditemukan"})
                    else:
                        names = xl.sheet_names[:1]
                    sheet_frames = {name: pd.read_excel(xl, sheet_name=name, dtype=str, keep_default_na=False) for name in names}

                for sheet, df in sheet_frames.items():
                    df.columns = clean_columns(df.columns)
                    # Hapus hanya baris sepenuhnya kosong; sisanya dipertahankan.
                    df = df.fillna("").astype(str)
                    nonempty = df.ne("").any(axis=1)
                    empty_count = int((~nonempty).sum())
                    df = df.loc[nonempty].copy()
                    n = len(df)
                    if add_source:
                        df.insert(0, "_sumber_sheet", str(sheet))
                        df.insert(0, "_sumber_file", filename)
                    frames.append(df)
                    audit.append({"File": filename, "Sheet": str(sheet), "Baris data": n, "Baris kosong dilewati": empty_count, "Kolom": len(df.columns), "Status": "OK"})
            except Exception as exc:
                issues.append({"File": filename, "Sheet": "-", "Masalah": f"Gagal dibaca: {type(exc).__name__}: {str(exc)[:200]}"})
                audit.append({"File": filename, "Sheet": "-", "Baris data": 0, "Baris kosong dilewati": 0, "Kolom": 0, "Status": "GAGAL"})

    if not frames:
        return pd.DataFrame(), pd.DataFrame(audit), pd.DataFrame(issues)
    combined = pd.concat(frames, ignore_index=True, sort=False).fillna("")
    return combined, pd.DataFrame(audit), pd.DataFrame(issues)


def build_excel(combined, audit, issues):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Limit Excel sheet rows (incl. header) and columns.
        if len(combined.columns) > 16384:
            raise ValueError("Kolom terlalu banyak untuk format XLSX.")
        chunk_size = 1_048_575
        for i, start in enumerate(range(0, max(len(combined), 1), chunk_size), 1):
            combined.iloc[start:start + chunk_size].to_excel(writer, index=False, sheet_name=f"Gabungan_{i}")
        audit.to_excel(writer, index=False, sheet_name="Laporan_Baris")
        issues.to_excel(writer, index=False, sheet_name="Masalah")
    output.seek(0)
    return output.getvalue()

uploaded = st.file_uploader("Pilih file ZIP", type=["zip"])
c1, c2 = st.columns(2)
with c1:
    sheet_mode = st.selectbox("Sheet yang akan digabung", ["Sheet pertama", "Semua sheet", "Nama sheet tertentu"])
with c2:
    requested_sheet = st.text_input("Nama sheet", value="Sheet1", disabled=sheet_mode != "Nama sheet tertentu")
add_source = st.checkbox("Tambahkan kolom asal file dan sheet", value=True)
st.caption("Baris yang seluruh selnya kosong dilewati. Baris duplikat tidak dihapus. Kolom disatukan berdasarkan nama header, bukan posisi.")

if uploaded and st.button("Proses ZIP", type="primary"):
    if uploaded.size > MAX_ZIP_SIZE:
        st.error("ZIP terlalu besar (maksimal 100 MB).")
    else:
        try:
            with st.spinner("Membaca dan menggabungkan file..."):
                combined, audit, issues = parse_zip(uploaded.getvalue(), sheet_mode, requested_sheet.strip(), add_source)
                expected = int(audit["Baris data"].sum()) if not audit.empty else 0
                actual = len(combined)
                st.session_state["merge_result"] = (combined, audit, issues, expected, actual)
        except (zipfile.BadZipFile, ValueError, OSError) as exc:
            st.error(f"ZIP tidak dapat diproses: {exc}")
            st.session_state.pop("merge_result", None)

if "merge_result" in st.session_state:
    combined, audit, issues, expected, actual = st.session_state["merge_result"]
    a, b, c = st.columns(3)
    a.metric("Baris dari sumber terbaca", f"{expected:,}")
    b.metric("Baris hasil gabungan", f"{actual:,}")
    c.metric("File/sheet bermasalah", len(issues))
    if expected == actual and issues.empty:
        st.success("Jumlah baris cocok dan tidak ada kesalahan baca yang terdeteksi.")
    elif not issues.empty:
        st.warning("Sebagian sumber bermasalah. Jangan gunakan hasil sebelum memeriksa laporan masalah.")
    else:
        st.error("Jumlah baris tidak cocok. Periksa proses penggabungan.")
    st.subheader("Preview gabungan")
    st.dataframe(combined.head(100), use_container_width=True)
    st.subheader("Laporan pemeriksaan")
    st.dataframe(audit, use_container_width=True)
    if not issues.empty:
        st.dataframe(issues, use_container_width=True)
    if not combined.empty:
        try:
            st.download_button("⬇️ Download Excel + laporan", data=build_excel(combined, audit, issues), file_name="hasil_gabungan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except ValueError as exc:
            st.error(str(exc))
        st.download_button("⬇️ Download CSV UTF-8", data=combined.to_csv(index=False).encode("utf-8-sig"), file_name="hasil_gabungan.csv", mime="text/csv")
    st.caption("Validasi ini memeriksa jumlah baris yang berhasil dibaca, bukan menjamin seluruh nilai asli tetap identik. Rumus, merged cells, kolom bernama berbeda, dan angka panjang memerlukan pemeriksaan khusus.")
