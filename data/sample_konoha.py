"""
data/sample_konoha.py
Generator dataset simulasi Kabupaten Konoha (XXXX) yang konsisten
dari input mentah sampai siap diproses Denton & LK4.
Nilai dalam Juta Rupiah (kecuali K3 Desa yang sengaja dalam Rupiah penuh
untuk menguji auto-normalisasi).
"""

import numpy as np

from core import constants as C


def _season(q):
    """Pola musiman sederhana per triwulan."""
    return [0.20, 0.24, 0.26, 0.30][q - 1]


def _grow(year, base, g=0.05):
    return base * ((1 + g) ** (year - 2019))


def make_quarters(current_quarter=(2026, 2)):
    cy, cq = current_quarter
    return C.build_quarters(C.START_YEAR, cy, cq)


def gen_apbn(quarters):
    """APBN per sub-tabel aktif (1.1-1.6). Nilai dalam Juta Rp."""
    data = {}
    base_level = {
        "1.1": 9000, "1.2": 1500, "1.3": 1200,
        "1.4": 7000, "1.5": 5000, "1.6": 800,
    }
    for code, base in base_level.items():
        rows = {r: [] for r in C.BASE_ROWS}
        for (y, q) in quarters:
            scale = _grow(y, base) * _season(q)
            rows["Belanja Barang"].append(scale * 0.30)
            rows["Belanja Bansos"].append(scale * 0.05)
            rows["Belanja Pegawai"].append(scale * 0.45)
            rows["Penyusutan"].append(scale * 0.08)
            rows["Penerimaan"].append(scale * 0.04)
        data[code] = rows
    return data


def gen_apbd_quarterly(quarters):
    """Indikator triwulanan APBD (raw, belum Denton)."""
    data = {}
    base_level = {
        "2.1": 30000, "2.2": 4000, "2.3": 3000,
        "2.4": 25000, "2.5": 18000, "2.6": 2500,
    }
    for code, base in base_level.items():
        rows = {r: [] for r in C.BASE_ROWS}
        for (y, q) in quarters:
            scale = _grow(y, base, g=0.04) * _season(q)
            rows["Belanja Barang"].append(scale * 0.28)
            rows["Belanja Bansos"].append(scale * 0.04)
            rows["Belanja Pegawai"].append(scale * 0.50)
            rows["Penyusutan"].append(scale * 0.10)
            rows["Penerimaan"].append(scale * 0.05)
        data[code] = rows
    return data


def gen_apbd_annual(quarters, quarterly):
    """Total tahunan APBD = jumlah indikator per tahun (agar Denton matching)."""
    annual = {}
    for code, rows in quarterly.items():
        annual[code] = {}
        for (y, q) in quarters:
            pass
        years = sorted({y for (y, _) in quarters})
        for y in years:
            tot = {}
            for r in C.BASE_ROWS:
                s = 0.0
                for (yy, q) in quarters:
                    if yy == y:
                        i = quarters.index((yy, q))
                        s += rows[r][i]
                tot[r] = s
            annual[code][y] = tot
    return annual


def gen_desa_k3(quarters):
    """Data Neraca K3 Desa dalam RUPIAH PENUH (menguji auto-normalisasi)."""
    annual = {}
    years = sorted({y for (y, _) in quarters})
    for y in years:
        scale = _grow(y, 5e10, g=0.06)  # Rp 50 miliar basis
        annual[y] = {
            "Belanja Pegawai": scale * 0.55,
            "Belanja Barang": scale * 0.25,
            "Belanja Modal": scale * 0.15,
            "Penerimaan": scale * 0.02,
        }
    return annual


def gen_ssf(quarters):
    """SSF per sub-tabel aktif (4.2-4.6). Juta Rp."""
    data = {}
    base_level = {
        "4.2": 6000, "4.3": 900, "4.4": 300, "4.5": 200, "4.6": 1500,
    }
    for code, base in base_level.items():
        rows = {r: [] for r in C.BASE_ROWS}
        for (y, q) in quarters:
            scale = _grow(y, base, g=0.07) * _season(q)
            rows["Belanja Barang"].append(scale * 0.42)
            rows["Belanja Bansos"].append(0.0)
            rows["Belanja Pegawai"].append(scale * 0.55)
            rows["Penyusutan"].append(scale * 0.03)
            rows["Penerimaan"].append(0.0)
        data[code] = rows
    return data


def gen_ssf_annual(quarters, quarterly):
    annual = {}
    years = sorted({y for (y, _) in quarters})
    for code, rows in quarterly.items():
        annual[code] = {}
        for y in years:
            tot = {}
            for r in C.BASE_ROWS:
                s = 0.0
                for (yy, q) in quarters:
                    if yy == y:
                        i = quarters.index((yy, q))
                        s += rows[r][i]
                tot[r] = s
            annual[code][y] = tot
    return annual


def gen_mbg(quarters):
    """MBG triwulanan: 2025 dalam RUPIAH PENUH, 2026 sudah Juta Rp (uji normalisasi)."""
    ser = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        if y < 2025:
            ser[lab] = 0.0
        elif y == 2025:
            ser[lab] = float(_grow(y, 2.0e9, g=0.0) * _season(q))  # Rupiah penuh
        else:
            ser[lab] = float(_grow(y, 3000.0, g=0.10) * _season(q))  # Juta Rp
    return ser


def gen_stik(quarters):
    """Total STIK SSF (Juta Rp)."""
    ser = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        ser[lab] = float(_grow(y, 5500.0, g=0.06) * _season(q)) if y >= 2021 else 0.0
    return ser


def gen_output_bi(quarters):
    """Output BI: default 0 untuk Kab. Konoha (hanya tercatat di Kota Konoha)."""
    return {C.qlabel(y, q): 0.0 for (y, q) in quarters}


def gen_deflators(quarters):
    """Deflator 2020Q1-2025Q4 + 2026, basis 2025=100."""
    defl = {name: {} for name in C.DEFLATOR_NAMES}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        t = (y - 2025) + (q - 1) * 0.25
        defl["IHPB Umum tanpa Ekspor"][lab] = 100.0 + 1.2 * t * 4
        defl["Indeks Upah adj_2"][lab] = 100.0 + 1.5 * t * 4
        defl["IHK"][lab] = 100.0 + 1.0 * t * 4
        if y == 2025:
            defl["Implisit PMTB"][lab] = C.IMPLISIT_PMTB_2025[q]
        else:
            defl["Implisit PMTB"][lab] = 100.0 + 0.9 * t * 4
    return defl


def build_sample_dataset(current_quarter=(2026, 2)):
    """Kembalian dict dataset simulasi lengkap."""
    quarters = make_quarters(current_quarter)
    apbn_q = gen_apbn(quarters)
    apbd_q = gen_apbd_quarterly(quarters)
    apbd_a = gen_apbd_annual(quarters, apbd_q)
    desa_k3 = gen_desa_k3(quarters)
    ssf_q = gen_ssf(quarters)
    ssf_a = gen_ssf_annual(quarters, ssf_q)
    return {
        "quarters": quarters,
        "current_quarter": current_quarter,
        "apbn_quarterly": apbn_q,
        "apbd_quarterly": apbd_q,
        "apbd_annual": apbd_a,
        "desa_k3": desa_k3,
        "ssf_quarterly": ssf_q,
        "ssf_annual": ssf_a,
        "mbg_total": gen_mbg(quarters),
        "stik_total": gen_stik(quarters),
        "output_bi": gen_output_bi(quarters),
        "deflators": gen_deflators(quarters),
    }