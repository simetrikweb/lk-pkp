"""
core/constants.py
Konfigurasi baku aplikasi PKP Rebasing 2025=100 - Kabupaten Konoha (XXXX).
Seluruh nilai dalam Juta Rupiah kecuali dinyatakan lain.
"""

# ------------------------------------------------------------------ #
# Konfigurasi waktu & wilayah
# ------------------------------------------------------------------ #
BASE_YEAR = 2025          # Tahun dasar rebasing
START_YEAR = 2019         # Awal deret historis
MAX_YEAR = 2030           # Batas atas struktur (hanya struktur, bukan aktif)
KODE_KAB = "XXXX"
NAMA_KAB = "Kabupaten Konoha"

# Label 5 baris input dasar (diisi pengguna / sumber data)
BASE_ROWS = [
    "Belanja Barang",
    "Belanja Bansos",
    "Belanja Pegawai",
    "Penyusutan",
    "Penerimaan",
]

# Label 9 baris baku (termasuk baris turunan)
STD_ROWS = [
    "Belanja Barang",     # 1
    "Belanja Bansos",     # 2
    "Belanja Pegawai",    # 3
    "Penyusutan",         # 4
    "Penerimaan",         # 5
    "Biaya Antara",       # 6  = BB + BS
    "NTB",                # 7  = BP + PY
    "Output",             # 8  = BA + NTB
    "PKP",                # 9  = Output - Penerimaan
]

# ------------------------------------------------------------------ #
# Metadata sub-tabel LK 3
# fields: code, name, sifat (Kolektif/Individu), is_rekap, rekap_of, active
# ------------------------------------------------------------------ #
def _mk(code, name, sifat, is_rekap=False, rekap_of=None, active=True):
    return {
        "code": code,
        "name": name,
        "sifat": sifat,
        "is_rekap": is_rekap,
        "rekap_of": rekap_of or [],
        "active": active,
    }

APBN_SUBTABLES = [
    _mk("1.1", "Administrasi Pemerintahan & Jamsos Wajib (KBLI O)", "Kolektif"),
    _mk("1.2", "Pengelolaan Air/Sampah/Limbah & Remediasi (KBLI E)", "Kolektif"),
    _mk("1.3", "Informasi & Komunikasi / RRI-TVRI (KBLI J)", "Kolektif"),
    _mk("1.4", "Pendidikan (KBLI P)", "Individu"),
    _mk("1.5", "Aktivitas Kesehatan & Sosial (KBLI Q)", "Individu"),
    _mk("1.6", "Kebudayaan, Hiburan & Rekreasi (KBLI R/S)", "Individu"),
    _mk("1.7", "Rekap Kolektif (1.1+1.2+1.3)", "Kolektif", is_rekap=True, rekap_of=["1.1", "1.2", "1.3"]),
    _mk("1.8", "Rekap Individu (1.4+1.5+1.6)", "Individu", is_rekap=True, rekap_of=["1.4", "1.5", "1.6"]),
    _mk("1.9", "Rekap Total (1.7+1.8)", "Total", is_rekap=True, rekap_of=["1.7", "1.8"]),
]

APBD_SUBTABLES = [
    _mk("2.1", "Administrasi Pemerintahan & Jamsos Wajib", "Kolektif"),
    _mk("2.2", "Pengelolaan Air/Sampah/Limbah & Remediasi", "Kolektif"),
    _mk("2.3", "Informasi & Komunikasi", "Kolektif"),
    _mk("2.4", "Pendidikan", "Individu"),
    _mk("2.5", "Aktivitas Kesehatan & Sosial", "Individu"),
    _mk("2.6", "Kebudayaan, Hiburan & Rekreasi", "Individu"),
    _mk("2.7", "Rekap Kolektif (2.1+2.2+2.3)", "Kolektif", is_rekap=True, rekap_of=["2.1", "2.2", "2.3"]),
    _mk("2.8", "Rekap Individu (2.4+2.5+2.6)", "Individu", is_rekap=True, rekap_of=["2.4", "2.5", "2.6"]),
    _mk("2.9", "Rekap Total (2.7+2.8)", "Total", is_rekap=True, rekap_of=["2.7", "2.8"]),
]

DESA_SUBTABLES = [
    _mk("3.1", "Administrasi Pemerintahan Desa (Kolektif)", "Kolektif"),
]

SSF_SUBTABLES = [
    _mk("4.1", "Rekap Induk SSF", "Kolektif", is_rekap=True,
        rekap_of=["4.2", "4.3", "4.4", "4.5", "4.6"]),
    _mk("4.2", "Jaminan Kesehatan (Jamkes/BPJS Kes)", "Kolektif"),
    _mk("4.3", "JKK - BPJS Ketenagakerjaan", "Kolektif"),
    _mk("4.4", "Jaminan Kematian (JK)", "Kolektif"),
    _mk("4.5", "Jaminan Kehilangan Pekerjaan (JKP)", "Kolektif"),
    _mk("4.6", "Jaminan Pensiun (JP)", "Kolektif"),
]

# Sumber yang melalui proses Denton (APBN = bypass audited)
DENTON_SOURCES = ["APBD", "Desa", "SSF"]
BYPASS_DENTON_SOURCES = ["APBN"]

# ------------------------------------------------------------------ #
# Rasio SPUN-O untuk APBDes (triwulanan, porsi dari tahunan)
# Struktur: {tahun: {komponen: [q1,q2,q3,q4]}}
# ------------------------------------------------------------------ #
SPUNO_RATIO = {
    2024: {
        "Belanja Pegawai": [0.08, 0.19, 0.30, 0.43],
        "Belanja Barang":  [0.05, 0.16, 0.29, 0.50],
        "Belanja Modal":   [0.04, 0.16, 0.30, 0.50],
        "Penerimaan":      [0.15, 0.22, 0.28, 0.35],
    },
    2025: {
        "Belanja Pegawai": [0.11, 0.22, 0.25, 0.42],
        "Belanja Barang":  [0.08, 0.18, 0.25, 0.48],
        "Belanja Modal":   [0.12, 0.22, 0.24, 0.42],
        "Penerimaan":      [0.23, 0.28, 0.22, 0.27],
    },
}

# ------------------------------------------------------------------ #
# Rasio MBG (Makan Bergizi Gratis) - pecah ke Pendidikan/Kesehatan
# {period: (rasio_pendidikan, rasio_kesehatan)}
# ------------------------------------------------------------------ #
MBG_RATIO = {
    (2025, 1): (0.99, 0.01),
    (2025, 2): (0.99, 0.01),
    (2025, 3): (0.97, 0.03),
    (2025, 4): (0.79, 0.21),
    (2026, 1): (0.88, 0.12),
    (2026, 2): (0.92, 0.08),
}

# ------------------------------------------------------------------ #
# Rasio STIK (klaim jaminan sosial) - Jamkes vs JKK
# {(tahun, triwulan): (jamkes, jkk)} ; None = default umum
# ------------------------------------------------------------------ #
STIK_RATIO = {
    "default_2025_q2_sebelum": (0.99, 0.01),   # s.d. 2025Q2
    "default_2025_q3_setelah": (0.98, 0.02),   # mulai 2025Q3
}

def stik_ratio_for(year, q):
    """Pilih rasio STIK berdasarkan periode."""
    if (year, q) <= (2025, 2):
        return STIK_RATIO["default_2025_q2_sebelum"]
    return STIK_RATIO["default_2025_q3_setelah"]

# ------------------------------------------------------------------ #
# Kalkulator komponen BPJS Kesehatan (bila file alokasi kantor tak lengkap)
# ------------------------------------------------------------------ #
BPJS_KES_COMPONENT = {
    "Belanja Barang": 0.42,
    "Belanja Pegawai": 0.55,
    "Penyusutan": 0.03,
}

# ------------------------------------------------------------------ #
# Nilai imputasi untuk Kasus 2 Denton (triwulan 0 padahal total tahunan > 0)
# ------------------------------------------------------------------ #
DENTON_ZERO_IMPUTE = 1e-6

# ------------------------------------------------------------------ #
# Deflator bawaan (basis 2025=100). Indeks, bukan persen.
# ------------------------------------------------------------------ #
DEFLATOR_NAMES = [
    "IHPB Umum tanpa Ekspor",
    "Indeks Upah adj_2",
    "Implisit PMTB",
    "IHK",
]

# Implisit PMTB 2025 sesuai pedoman (2025Q1..Q4)
IMPLISIT_PMTB_2025 = {1: 100.43, 2: 101.31, 3: 102.20, 4: 103.09}

# Pemetaan komponen -> deflator untuk deflasi ADHB ke ADHK
DEFLATOR_MAP = {
    "Belanja Pegawai": "Indeks Upah adj_2",
    "Penyusutan": "Implisit PMTB",
    "Belanja Barang": "IHPB Umum tanpa Ekspor",
    "Belanja Bansos": "IHPB Umum tanpa Ekspor",
    "Penerimaan": "IHK",
    "STIK": "IHK",
    "MBG": "IHK",
}

# ------------------------------------------------------------------ #
# Identitas Upload Siperi
# ------------------------------------------------------------------ #
SIPERI_ID = {
    "PKP_Total": "10",      # 3. Pengeluaran Konsumsi Pemerintah
    "Kolektif": "11",       # a. Konsumsi Kolektif
    "Individu": "12",       # b. Konsumsi Individu
}

# Enam lapangan usaha jasa pemerintah untuk Upload Verdana
VERDANA_LAPANGAN_USAHA = [
    "Sampah/Limbah",
    "Infokom Pemerintah",
    "O Administrasi Pemerintahan",
    "Pendidikan Pemerintah",
    "Kesehatan & Sosial Pemerintah",
    "Jasa Lainnya Pemerintah",
]

# ------------------------------------------------------------------ #
# Utilitas triwulan
# ------------------------------------------------------------------ #
def build_quarters(start_year, end_year, end_q=4):
    """Daftar tuple (tahun, triwulan) dari start_year Q1 sampai end_year end_q."""
    out = []
    for y in range(start_year, end_year + 1):
        for q in range(1, 5):
            if y == end_year and q > end_q:
                continue
            out.append((y, q))
    return out

def qlabel(y, q):
    """Label triwulan, mis. (2019,1) -> '2019Q1'."""
    return f"{y}Q{q}"

def qlabels(quarters):
    return [qlabel(y, q) for (y, q) in quarters]

def quarter_of_label(label):
    """'2019Q1' -> (2019,1)."""
    y, q = label.split("Q")
    return int(y), int(q)

def years_in_range(quarters):
    """Urutan tahun unik dalam rentang."""
    return sorted({y for (y, _) in quarters})

def complete_years(quarters):
    """Tahun yang memiliki 4 triwulan penuh dalam rentang aktif."""
    from collections import Counter
    cnt = Counter(y for (y, _) in quarters)
    return sorted(y for y, c in cnt.items() if c == 4)

def quarter_index_map(quarters):
    """Map label -> posisi indeks dalam list quarters."""
    return {qlabel(y, q): i for i, (y, q) in enumerate(quarters)}