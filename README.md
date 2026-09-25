# BAST Maker Engine

Generator dokumen BAST (Berita Acara Serah Terima) untuk kegiatan survei/sensus BPS.

## Features

- **Format 1 (Kolektif)**: N halaman muka PPL + 1 lampiran rekap di akhir
- **Format 2 (Individual)**: 1 BAST set per petugas
- Export ke ZIP Word satuan, Word gabungan, atau PDF gabungan
- Form sidebar untuk pengaturan statis (nomor surat, tanggal, identitas PPK)
- Editor tabel interaktif untuk data lapangan (nama, NIK, alamat, volume)

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## CSV Format

```csv
nama_ppl,nik_ppl,alamat_ppl,vol_kegiatan
Ahmad Fauzi,3XXX01234567XXX1,Desa Arjowinangun,4
```

## Template

Upload your own `.docx` template with Jinja2 placeholders (`{{ variable }}`), or use the default template included in `template/BAST MASTER.docx`.
