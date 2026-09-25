def build_context(row, settings, dates):
    """Build docxtpl context dict from a DataFrame row, sidebar settings, and date outputs."""
    return {
        "nama_kegiatan": settings["nama_kegiatan"],
        "nama_kegiatan_upper": settings["nama_kegiatan"].upper(),
        "alamat_bps": settings["alamat_bps"],
        "nama_ppk": settings["nama_ppk"],
        "nip_ppk": settings["nip_ppk"],
        "jabatan_ppk": settings["jabatan_ppk"],
        "tipe_satuan": settings["tipe_satuan"],
        "satuan_vol": settings["satuan_vol"],
        "hari": dates["hari"],
        "tanggal_terbilang": dates["tgl_t"],
        "bulan_terbilang": dates["bln_t"],
        "terbilang": dates["thn_t"],
        "tanggal_bast": dates["tgl_bast_str"],
        "tanggal_st_terbilang": dates["tgl_st_t"],
        "bulan_st_terbilang": dates["bln_st_t"],
        "tahun_st": str(dates["tgl_st_year"]),
        "nama_ppl": str(row.get("nama_ppl", "")),
        "nik_ppl": str(row.get("nik_ppl", "")),
        "alamat_ppl": str(row.get("alamat_ppl", "")),
        "nomor_bast": settings.get("nomor_bast", ""),
        "nomor_surattugas": settings.get("nomor_surattugas", ""),
        "vol_kegiatan": str(row.get("vol_kegiatan", ""))
    }
