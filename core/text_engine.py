import datetime

ANGKA_KATA = {
    1: "satu", 2: "dua", 3: "tiga", 4: "empat", 5: "lima",
    6: "enam", 7: "tujuh", 8: "delapan", 9: "sembilan", 10: "sepuluh",
    11: "sebelas", 12: "dua belas", 13: "tiga belas", 14: "empat belas",
    15: "lima belas", 16: "enam belas", 17: "tujuh belas", 18: "delapan belas",
    19: "sembilan belas", 20: "dua puluh", 30: "tiga puluh", 40: "empat puluh",
    50: "lima puluh", 60: "enam puluh", 70: "tujuh puluh", 80: "delapan puluh",
    90: "sembilan puluh"
}

BULAN_INDO = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

HARI_INDO = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def bilang_angka(n: int) -> str:
    if n in ANGKA_KATA:
        return ANGKA_KATA[n]
    elif n < 100:
        puluh = (n // 10) * 10
        sisa = n % 10
        return f"{ANGKA_KATA[puluh]} {ANGKA_KATA[sisa]}"
    elif n < 200:
        return f"seratus {bilang_angka(n - 100)}".strip()
    elif n < 1000:
        ratus = n // 100
        sisa = n % 100
        return f"{ANGKA_KATA[ratus]} ratus {bilang_angka(sisa)}".strip()
    elif n < 2000:
        return f"seribu {bilang_angka(n - 1000)}".strip()
    elif n < 1000000:
        ribu = n // 1000
        sisa = n % 1000
        return f"{bilang_angka(ribu)} ribu {bilang_angka(sisa)}".strip()
    return str(n)


def extract_date_terbilang(d: datetime.date):
    hari = HARI_INDO[d.weekday()]
    tgl_terbilang = bilang_angka(d.day)
    bln_terbilang = BULAN_INDO[d.month]
    thn_terbilang = bilang_angka(d.year)
    tgl_standar = f"{d.day} {BULAN_INDO[d.month]} {d.year}"
    return hari, tgl_terbilang, bln_terbilang, thn_terbilang, tgl_standar
