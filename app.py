import os
import re
import json
import datetime
import pandas as pd
import streamlit as st

from core.text_engine import extract_date_terbilang
from core.context import build_context
from core.render import render_docs
from core.export_docx import export_zip, export_merged_docx
from core.export_pdf import export_merged_pdf
from core.perjadin import render_perjadin, PERJADIN_TEMPLATE

SETTINGS_FILE = "saved_settings.json"
PERJADIN_SETTINGS_FILE = "saved_settings_perjadin.json"
DEFAULT_TEMPLATE = "template/BAST MASTER.docx"


def slugify(text, max_len=40):
    text = text.strip().upper()
    text = re.sub(r'[^A-Z0-9]+', '_', text)
    text = re.sub(r'_+', '_', text).strip('_')
    return text[:max_len]


def load_json(path, default):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return dict(default)


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


st.set_page_config(page_title="BAST Maker Engine", layout="wide")

# ============================================================
# SIDEBAR NAVIGATION
# ============================================================
st.sidebar.title("BAST Maker Engine")
page = st.sidebar.radio(
    "Menu",
    ["BAST Generator", "Laporan Perjadin"],
    label_visibility="collapsed",
)


# ============================================================
# PAGE: BAST Generator
# ============================================================
if page == "BAST Generator":
    st.title("Generator BAST Dinamis")

    saved_settings = load_json(SETTINGS_FILE, {
        "nama_kegiatan": "Survei Komoditas Strategis",
        "alamat_bps": "Jl. Veteran No. 12, Pacitan",
        "nama_ppk": "Nama PPK, M.Si.",
        "nip_ppk": "198801012010121001",
        "jabatan_ppk": "Pejabat Pembuat Komitmen",
        "tipe_satuan": "hardcopy",
        "satuan_vol": "Dokumen",
        "nomor_bast": "093/35013/UBINAN/BAST/2026",
        "nomor_surattugas": "B-076/35010/VS.330/2026",
    })

    # --- Sidebar: Preset ---
    st.sidebar.header("Pengaturan Statis (Non-CSV)")
    with st.sidebar.expander("📁 Preset Profil (JSON)"):
        preset_upload = st.file_uploader("Unggah Preset", type=["json"], key="bast_preset_up", label_visibility="collapsed")
        if preset_upload is not None:
            try:
                st.session_state["loaded_preset"] = json.load(preset_upload)
                st.rerun()
            except json.JSONDecodeError:
                st.error("Berkas JSON tidak valid.")

        st.download_button(
            label="📥 Unduh Preset Profil",
            data=json.dumps(saved_settings, indent=2, ensure_ascii=False),
            file_name=f"preset_{slugify(saved_settings.get('nama_kegiatan', 'BAST'))}.json",
            mime="application/json", use_container_width=True, key="bast_preset_dl",
        )

    defaults = dict(saved_settings)
    if "loaded_preset" in st.session_state:
        defaults.update(st.session_state.pop("loaded_preset"))

    # --- Sidebar: Settings Form ---
    with st.sidebar.form("bast_settings"):
        nama_kegiatan = st.text_input("Nama Kegiatan", defaults.get("nama_kegiatan", ""))
        alamat_bps = st.text_input("Alamat Kantor BPS", defaults.get("alamat_bps", ""))
        st.markdown("**Identitas PPK (Pihak Kedua)**")
        nama_ppk = st.text_input("Nama PPK", defaults.get("nama_ppk", ""))
        nip_ppk = st.text_input("NIP PPK", defaults.get("nip_ppk", ""))
        jabatan_ppk = st.text_input("Jabatan PPK", defaults.get("jabatan_ppk", ""))
        st.markdown("**Naskah BAST**")
        nomor_bast = st.text_input("Nomor BAST", defaults.get("nomor_bast", ""))
        tgl_bast = st.date_input("Tanggal Pelaksanaan BAST",
            datetime.date.fromisoformat(defaults.get("tgl_bast", str(datetime.date.today()))))
        st.markdown("**Surat Tugas**")
        nomor_surattugas = st.text_input("Nomor Surat Tugas", defaults.get("nomor_surattugas", ""))
        tgl_st = st.date_input("Tanggal Surat Tugas (ST)",
            datetime.date.fromisoformat(defaults.get("tgl_st", str(datetime.date.today()))))
        st.markdown("**Format Satuan**")
        tipe_satuan = st.text_input("Bentuk Dokumen", defaults.get("tipe_satuan", "hardcopy"))
        satuan_vol = st.text_input("Satuan Volume", defaults.get("satuan_vol", "Dokumen"))

        if st.form_submit_button("Simpan Pengaturan Default"):
            saved_settings.update({
                "nama_kegiatan": nama_kegiatan, "alamat_bps": alamat_bps,
                "nama_ppk": nama_ppk, "nip_ppk": nip_ppk, "jabatan_ppk": jabatan_ppk,
                "nomor_bast": nomor_bast, "nomor_surattugas": nomor_surattugas,
                "tipe_satuan": tipe_satuan, "satuan_vol": satuan_vol,
                "tgl_bast": str(tgl_bast), "tgl_st": str(tgl_st),
            })
            save_json(SETTINGS_FILE, saved_settings)
            st.success("Konfigurasi statis tersimpan!")

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

    # --- Main: Template & CSV ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. File Template Word")
        tpl_file = st.file_uploader("Upload .docx (atau kosongkan untuk default)", type=["docx"], key="bast_tpl")
        if tpl_file is None and os.path.exists(DEFAULT_TEMPLATE):
            st.info(f"Template default: `{DEFAULT_TEMPLATE}`")
    with col2:
        st.subheader("2. File CSV Dinamis")
        csv_file = st.file_uploader("Upload Data CSV", type=["csv"], key="bast_csv")

    # --- Data editor ---
    initial_df = pd.DataFrame([{
        "nama_ppl": "Ahmad Fauzi", "nik_ppl": "3501012345670001",
        "alamat_ppl": "Desa Arjowinangun Pacitan", "vol_kegiatan": "4"
    }])
    if "df_bast" not in st.session_state:
        st.session_state["df_bast"] = initial_df
    if csv_file is not None:
        st.session_state["df_bast"] = pd.read_csv(csv_file, dtype=str)

    st.subheader("3. Editor Data Lapangan")
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

    # --- Render ---
    st.subheader("4. Eksekusi Render")
    format_mode = st.radio(
        "Format Dokumen",
        ["individual", "collective"],
        format_func=lambda x: {
            "individual": "Format 2 — BAST Individual (1 set per petugas)",
            "collective": "Format 1 — Dokumen Kolektif (N halaman muka + 1 lampiran rekap)",
        }[x],
        horizontal=True, key="format_bast_choice",
        on_change=lambda: st.session_state.pop("result", None),
    )

    def bast_context_fn(row):
        ctx = build_context(row, settings, dates)
        if format_mode == "collective":
            ctx["total_vol_kegiatan"] = str(
                edited_df["vol_kegiatan"].astype(int).sum()
            ) if "vol_kegiatan" in edited_df.columns else "0"
        return ctx

    if st.button("Generate Dokumen", type="primary", use_container_width=True, key="btn_bast"):
        if tpl_file is None and not os.path.exists(DEFAULT_TEMPLATE):
            st.error("Silakan unggah template .docx atau pastikan template default ada.")
        elif edited_df.empty:
            st.error("Tabel data petugas tidak boleh kosong.")
        else:
            progress = st.progress(0, text="Memulai render...")
            progress.progress(10, text="Rendering dokumen dari template...")
            docs = render_docs(tpl_file, DEFAULT_TEMPLATE, edited_df, bast_context_fn, format_mode)
            progress.progress(40, text="Membuat ZIP Word satuan...")
            zip_buf = export_zip(docs)
            progress.progress(60, text="Menggabungkan Word gabungan...")
            merged_docx_buf = export_merged_docx(docs)
            progress.progress(80, text="Menggabungkan PDF gabungan...")
            merged_pdf_buf = export_merged_pdf(docs)
            progress.progress(100, text="Selesai!")
            slug = slugify(nama_kegiatan)
            fmt_label = "Kolektif" if format_mode == "collective" else "Individual"
            st.session_state["result"] = {
                "count": len(docs), "format": format_mode, "slug": slug, "fmt_label": fmt_label,
                "zip": zip_buf, "merged_docx": merged_docx_buf, "merged_pdf": merged_pdf_buf,
            }

    if "result" in st.session_state:
        r = st.session_state["result"]
        is_collective = r["format"] == "collective"
        slug = r["slug"]
        fmt = r["fmt_label"]
        st.success(f"Berhasil menghasilkan {r['count']} dokumen BAST.")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                label="ZIP Word Satuan" if not is_collective else "ZIP Word Kolektif",
                data=r["zip"], file_name=f"BAST_{slug}_{fmt}.zip",
                mime="application/zip", use_container_width=True, key="dl_zip",
            )
        with c2:
            st.download_button(
                label="Word Gabungan", data=r["merged_docx"],
                file_name=f"BAST_{slug}_{fmt}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="dl_docx",
            )
        with c3:
            st.download_button(
                label="PDF Gabungan", data=r["merged_pdf"],
                file_name=f"BAST_{slug}_{fmt}.pdf",
                mime="application/pdf", use_container_width=True, key="dl_pdf",
            )


# ============================================================
# PAGE: Laporan Perjadin
# ============================================================
else:
    st.title("Laporan Perjalanan Dinas")

    pj_defaults = load_json(PERJADIN_SETTINGS_FILE, {})

    # --- Sidebar: Perjadin Preset ---
    st.sidebar.header("Pengaturan Laporan Perjadin")
    with st.sidebar.expander("📁 Preset Laporan Perjadin (JSON)"):
        pj_preset_up = st.file_uploader("Unggah Preset", type=["json"], key="pj_preset_up", label_visibility="collapsed")
        if pj_preset_up is not None:
            try:
                st.session_state["pj_loaded_preset"] = json.load(pj_preset_up)
                st.rerun()
            except json.JSONDecodeError:
                st.error("Berkas JSON tidak valid.")

        pj_export = {
            "nama_kegiatan": pj_defaults.get("nama_kegiatan", ""),
            "nama_pelapor": st.session_state.get("pj_form", {}).get("nama_pelapor", ""),
            "NIP_NIK": st.session_state.get("pj_form", {}).get("NIP_NIK", ""),
            "pangkat": st.session_state.get("pj_form", {}).get("pangkat", ""),
            "jabatan_kegiatan": st.session_state.get("pj_form", {}).get("jabatan_kegiatan", ""),
            "jabatan": st.session_state.get("pj_form", {}).get("jabatan", ""),
            "catatan_hasil": st.session_state.get("pj_form", {}).get("catatan_hasil", ""),
            "kendala": st.session_state.get("pj_form", {}).get("kendala", ""),
            "solusi": st.session_state.get("pj_form", {}).get("solusi", ""),
        }
        st.download_button(
            label="📥 Unduh Preset Perjadin",
            data=json.dumps(pj_export, indent=2, ensure_ascii=False),
            file_name=f"preset_perjadin_{slugify(pj_export.get('nama_pelapor', 'Pelapor'))}.json",
            mime="application/json", use_container_width=True, key="pj_preset_dl",
        )

    # Apply loaded preset
    if "pj_loaded_preset" in st.session_state:
        pj_defaults.update(st.session_state.pop("pj_loaded_preset"))

    # --- Sidebar: Save current form as default ---
    if st.sidebar.button("💾 Simpan Form sebagai Default", use_container_width=True, key="pj_save_default"):
        pj_defaults.update({
            "nama_pelapor": st.session_state.get("pj_form", {}).get("nama_pelapor", ""),
            "NIP_NIK": st.session_state.get("pj_form", {}).get("NIP_NIK", ""),
            "pangkat": st.session_state.get("pj_form", {}).get("pangkat", ""),
            "jabatan_kegiatan": st.session_state.get("pj_form", {}).get("jabatan_kegiatan", ""),
            "jabatan": st.session_state.get("pj_form", {}).get("jabatan", ""),
            "catatan_hasil": st.session_state.get("pj_form", {}).get("catatan_hasil", ""),
            "kendala": st.session_state.get("pj_form", {}).get("kendala", ""),
            "solusi": st.session_state.get("pj_form", {}).get("solusi", ""),
        })
        save_json(PERJADIN_SETTINGS_FILE, pj_defaults)
        st.sidebar.success("Default tersimpan!")

    # --- Main: Form ---
    nama_kegiatan_pj = st.text_input("Nama Kegiatan", pj_defaults.get("nama_kegiatan", ""), key="pj_nama_kegiatan")
    col1, col2 = st.columns(2)
    with col1:
        nama_pelapor = st.text_input("Nama Pelapor", pj_defaults.get("nama_pelapor", ""), key="pj_nama")
        NIP_NIK = st.text_input("NIP / NIK", pj_defaults.get("NIP_NIK", ""), key="pj_nip")
        pangkat = st.text_input("Pangkat / Golongan", pj_defaults.get("pangkat", ""), key="pj_pangkat")
    with col2:
        jabatan_kegiatan = st.text_input("Jabatan (di Kegiatan)", pj_defaults.get("jabatan_kegiatan", ""), key="pj_jab_keg")
        jabatan = st.text_input("Jabatan", pj_defaults.get("jabatan", ""), key="pj_jab")
        tanggal_OH = st.date_input("Tanggal Perjalanan", key="pj_tgl_oh")
        tanggal_ttd = st.date_input("Tanggal Tanda Tangan", key="pj_tgl_ttd")

    _, _, _, _, tanggal_OH_str = extract_date_terbilang(tanggal_OH)
    _, _, _, _, tanggal_ttd_str = extract_date_terbilang(tanggal_ttd)

    st.markdown("---")
    st.subheader("Catatan Lapangan")
    col1, col2 = st.columns(2)
    with col1:
        catatan_hasil = st.text_area("Catatan Hasil Pendataan", pj_defaults.get("catatan_hasil", ""), key="pj_catatan", height=120)
        kendala = st.text_area("Kendala", pj_defaults.get("kendala", ""), key="pj_kendala", height=120)
    with col2:
        solusi = st.text_area("Solusi", pj_defaults.get("solusi", ""), key="pj_solusi", height=120)

    st.markdown("---")
    st.subheader("Dokumentasi Foto")
    st.caption("Upload foto kegiatan. Foto akan tersusun otomatis dalam grid A4 (2 kolom) di halaman Dokumentasi.")
    photos = st.file_uploader(
        "Upload Foto", type=["jpg", "jpeg", "jfif", "png", "bmp", "webp"],
        accept_multiple_files=True, key="pj_photos"
    )

    if photos:
        cols = st.columns(min(len(photos), 4))
        for i, p in enumerate(photos):
            with cols[i % len(cols)]:
                st.image(p, caption=p.name, use_container_width=True)
        st.caption(f"{len(photos)} foto siap digunakan.")

    # --- Store form values for preset export ---
    st.session_state["pj_form"] = {
        "nama_kegiatan": nama_kegiatan_pj,
        "nama_pelapor": nama_pelapor, "NIP_NIK": NIP_NIK, "pangkat": pangkat,
        "jabatan_kegiatan": jabatan_kegiatan, "jabatan": jabatan,
        "catatan_hasil": catatan_hasil, "kendala": kendala, "solusi": solusi,
    }

    if st.button("Generate Laporan Perjadin", type="primary", use_container_width=True, key="btn_perjadin"):
        if not nama_pelapor:
            st.error("Nama pelapor wajib diisi.")
        else:
            context = {
                "nama_kegiatan": nama_kegiatan_pj,
                "nama_pelapor": nama_pelapor,
                "NIP_NIK": NIP_NIK, "pangkat": pangkat,
                "jabatan_kegiatan": jabatan_kegiatan, "jabatan": jabatan,
                "tanggal_OH": tanggal_OH_str, "tanggal_ttd": tanggal_ttd_str,
                "catatan_hasil": catatan_hasil, "kendala": kendala, "solusi": solusi,
            }
            photo_bytes = [p.getvalue() for p in photos] if photos else []
            docx_bytes = render_perjadin(context, photo_bytes)

            slug = slugify(nama_pelapor or "Perjadin")
            st.success("Laporan Perjadin berhasil digenerate!")
            st.download_button(
                label="Unduh Laporan Perjadin (.docx)",
                data=docx_bytes,
                file_name=f"Perjadin_{slug}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary", use_container_width=True, key="dl_perjadin",
            )
