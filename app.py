import os
import json
import datetime
import pandas as pd
import streamlit as st

from core.text_engine import extract_date_terbilang
from core.context import build_context
from core.render import render_docs
from core.export_docx import export_zip, export_merged_docx
from core.export_pdf import export_merged_pdf

SETTINGS_FILE = "saved_settings.json"
DEFAULT_TEMPLATE = "template/BAST MASTER.docx"

st.set_page_config(page_title="BAST Generator BPS", layout="wide")
st.title("Generator BAST Dinamis")

# --- Load settings ---
if os.path.exists(SETTINGS_FILE):
    with open(SETTINGS_FILE, "r") as f:
        saved_settings = json.load(f)
else:
    saved_settings = {
        "nama_kegiatan": "Survei Komoditas Strategis",
        "alamat_bps": "Jl. Veteran No. 12, Pacitan",
        "nama_ppk": "Nama PPK, M.Si.",
        "nip_ppk": "198801012010121001",
        "jabatan_ppk": "Pejabat Pembuat Komitmen",
        "tipe_satuan": "hardcopy",
        "satuan_vol": "Dokumen",
        "nomor_bast": "093/35013/UBINAN/BAST/2026",
        "nomor_surattugas": "B-076/35010/VS.330/2026",
    }

# --- Sidebar ---
st.sidebar.header("Pengaturan Statis (Non-CSV)")
with st.sidebar.form("settings_form"):
    nama_kegiatan = st.text_input("Nama Kegiatan", saved_settings.get("nama_kegiatan", ""))
    alamat_bps = st.text_input("Alamat Kantor BPS", saved_settings.get("alamat_bps", ""))

    st.markdown("**Identitas PPK (Pihak Kedua)**")
    nama_ppk = st.text_input("Nama PPK", saved_settings.get("nama_ppk", ""))
    nip_ppk = st.text_input("NIP PPK", saved_settings.get("nip_ppk", ""))
    jabatan_ppk = st.text_input("Jabatan PPK", saved_settings.get("jabatan_ppk", ""))

    st.markdown("**Naskah BAST**")
    nomor_bast = st.text_input("Nomor BAST", saved_settings.get("nomor_bast", ""))
    tgl_bast = st.date_input("Tanggal Pelaksanaan BAST",
        datetime.date.fromisoformat(saved_settings.get("tgl_bast", str(datetime.date.today()))))

    st.markdown("**Surat Tugas**")
    nomor_surattugas = st.text_input("Nomor Surat Tugas", saved_settings.get("nomor_surattugas", ""))
    tgl_st = st.date_input("Tanggal Surat Tugas (ST)",
        datetime.date.fromisoformat(saved_settings.get("tgl_st", str(datetime.date.today()))))

    st.markdown("**Format Satuan**")
    tipe_satuan = st.text_input("Bentuk Dokumen", saved_settings.get("tipe_satuan", "hardcopy"))
    satuan_vol = st.text_input("Satuan Volume", saved_settings.get("satuan_vol", "Dokumen"))

    if st.form_submit_button("Simpan Pengaturan Default"):
        saved_settings.update({
            "nama_kegiatan": nama_kegiatan, "alamat_bps": alamat_bps,
            "nama_ppk": nama_ppk, "nip_ppk": nip_ppk, "jabatan_ppk": jabatan_ppk,
            "nomor_bast": nomor_bast, "nomor_surattugas": nomor_surattugas,
            "tipe_satuan": tipe_satuan, "satuan_vol": satuan_vol,
            "tgl_bast": str(tgl_bast), "tgl_st": str(tgl_st),
        })
        with open(SETTINGS_FILE, "w") as f:
            json.dump(saved_settings, f, indent=2)
        st.success("Konfigurasi statis tersimpan!")

# --- Date conversion ---
hari, tgl_t, bln_t, thn_t, tgl_bast_str = extract_date_terbilang(tgl_bast)
_, tgl_st_t, bln_st_t, _, _ = extract_date_terbilang(tgl_st)
dates = {
    "hari": hari, "tgl_t": tgl_t, "bln_t": bln_t, "thn_t": thn_t,
    "tgl_bast_str": tgl_bast_str, "tgl_st_t": tgl_st_t, "bln_st_t": bln_st_t,
    "tgl_st_year": tgl_st.year,
}

settings = {
    "nama_kegiatan": nama_kegiatan, "alamat_bps": alamat_bps,
    "nama_ppk": nama_ppk, "nip_ppk": nip_ppk, "jabatan_ppk": jabatan_ppk,
    "nomor_bast": nomor_bast, "nomor_surattugas": nomor_surattugas,
    "tipe_satuan": tipe_satuan, "satuan_vol": satuan_vol,
}

# --- Template & CSV upload ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. File Template Word")
    tpl_file = st.file_uploader("Upload .docx (atau kosongkan untuk default)", type=["docx"])
    if tpl_file is None and os.path.exists(DEFAULT_TEMPLATE):
        st.info(f"Template default: `{DEFAULT_TEMPLATE}`")
with col2:
    st.subheader("2. File CSV Dinamis")
    csv_file = st.file_uploader("Upload Data CSV", type=["csv"])

# --- Data editor (4 columns only) ---
initial_df = pd.DataFrame([{
    "nama_ppl": "Ahmad Fauzi", "nik_ppl": "3501012345670001",
    "alamat_ppl": "Desa Arjowinangun Pacitan", "vol_kegiatan": "4"
}])

if "df_bast" not in st.session_state:
    st.session_state["df_bast"] = initial_df
if csv_file is not None:
    st.session_state["df_bast"] = pd.read_csv(csv_file, dtype=str)

st.subheader("3. Editor Data Lapangan")
st.caption("Ubah, tambah, atau hapus baris langsung pada tabel berikut.")
edited_df = st.data_editor(
    st.session_state["df_bast"], num_rows="dynamic",
    use_container_width=True, key="bast_editor"
)
st.session_state["df_bast"] = edited_df

st.download_button(
    label="Unduh CSV Hasil Koreksi",
    data=edited_df.to_csv(index=False).encode("utf-8"),
    file_name="bast_petugas_corrected.csv", mime="text/csv"
)
st.markdown("---")

# --- Format selection ---
st.subheader("4. Eksekusi Render")
format_mode = st.radio(
    "Format Dokumen",
    ["individual", "collective"],
    format_func=lambda x: {
        "individual": "Format 2 — BAST Individual (1 set per petugas)",
        "collective": "Format 1 — Dokumen Kolektif (N halaman muka + 1 lampiran rekap)",
    }[x],
    horizontal=True,
    key="format_bast_choice",
    on_change=lambda: st.session_state.pop("result", None),
)

def context_fn(row):
    ctx = build_context(row, settings, dates)
    if format_mode == "collective":
        ctx["total_vol_kegiatan"] = str(
            edited_df["vol_kegiatan"].astype(int).sum()
        ) if "vol_kegiatan" in edited_df.columns else "0"
    return ctx

def validate():
    if tpl_file is None and not os.path.exists(DEFAULT_TEMPLATE):
        return "Silakan unggah template .docx atau pastikan template default ada."
    if edited_df.empty:
        return "Tabel data petugas tidak boleh kosong."
    return None

if st.button("Generate Dokumen", type="primary", use_container_width=True):
    err = validate()
    if err:
        st.error(err)
    else:
        progress = st.progress(0, text="Memulai render...")

        progress.progress(10, text="Rendering dokumen dari template...")
        docs = render_docs(tpl_file, DEFAULT_TEMPLATE, edited_df, context_fn, format_mode)

        progress.progress(40, text="Membuat ZIP Word satuan...")
        zip_buf = export_zip(docs)

        progress.progress(60, text="Menggabungkan Word gabungan...")
        merged_docx_buf = export_merged_docx(docs)

        progress.progress(80, text="Menggabungkan PDF gabungan...")
        merged_pdf_buf = export_merged_pdf(docs)

        progress.progress(100, text="Selesai!")
        st.session_state["result"] = {
            "count": len(docs),
            "format": format_mode,
            "zip": zip_buf,
            "merged_docx": merged_docx_buf,
            "merged_pdf": merged_pdf_buf,
        }

if "result" in st.session_state:
    r = st.session_state["result"]
    is_collective = r["format"] == "collective"
    st.success(f"Berhasil menghasilkan {r['count']} dokumen BAST.")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.download_button(
            label="ZIP Word Satuan" if not is_collective else "ZIP Word Kolektif",
            data=r["zip"],
            file_name="BAST_Generated_Batch.zip",
            mime="application/zip",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            label="Word Gabungan",
            data=r["merged_docx"],
            file_name="BAST_Gabungan.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with c3:
        st.download_button(
            label="PDF Gabungan",
            data=r["merged_pdf"],
            file_name="BAST_Gabungan.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
