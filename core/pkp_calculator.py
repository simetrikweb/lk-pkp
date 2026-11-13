"""
core/pkp_calculator.py
Perhitungan inti PKP:
  - Rumus 9 baris baku per sub-tabel
  - Growth q-to-q / y-on-y / c-to-c
  - Konstruksi APBDes (K3 x SPUN-O)
  - Pemecahan STIK / MBG
  - Konsolidasi LK4 (ADHB -> ADHK, deflasi per komponen)
  - Rekap Kolektif / Individu
"""

import numpy as np

from . import constants as C


# ------------------------------------------------------------------ #
# Util deret waktu
# ------------------------------------------------------------------ #
def series_to_dict(values, quarters):
    return {C.qlabel(y, q): values[i] for i, (y, q) in enumerate(quarters)}


def safe_div(a, b):
    if b is None or abs(b) < 1e-12:
        return None
    return a / b


def growth_qoq(ser, quarters):
    """Growth q-to-q (%) per triwulan."""
    out = {}
    labels = C.qlabels(quarters)
    for i, lab in enumerate(labels):
        if i == 0:
            out[lab] = None
            continue
        prev = ser.get(labels[i - 1])
        cur = ser.get(lab)
        g = None
        if prev not in (None,) and cur is not None and abs(prev) > 1e-12:
            g = (cur / prev - 1.0) * 100.0
        out[lab] = g
    return out


def growth_yoy(ser, quarters):
    """Growth year-on-year (%) per triwulan."""
    out = {}
    for i, (y, q) in enumerate(quarters):
        prev_lab = C.qlabel(y - 1, q)
        cur_lab = C.qlabel(y, q)
        prev = ser.get(prev_lab)
        cur = ser.get(cur_lab)
        g = None
        if prev is not None and cur is not None and abs(prev) > 1e-12:
            g = (cur / prev - 1.0) * 100.0
        out[cur_lab] = g
    return out


def growth_ctc(ser, quarters):
    """Growth cumulative-to-cumulative (%) per triwulan."""
    out = {}
    cum = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        val = ser.get(lab) or 0.0
        prev_cum = cum.get((y, q - 1), 0.0) if q > 1 else 0.0
        cum[(y, q)] = prev_cum + val
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        cur_cum = cum.get((y, q), 0.0)
        prev_year_cum = cum.get((y - 1, q), None)
        g = None
        if prev_year_cum not in (None,) and abs(prev_year_cum) > 1e-12:
            g = (cur_cum / prev_year_cum - 1.0) * 100.0
        out[lab] = g
    return out


def compute_growth_blocks(ser, quarters):
    """Kembalian dict berisi 4 blok: Nilai, qoq, yoy, ctc."""
    return {
        "Nilai": ser,
        "qoq": growth_qoq(ser, quarters),
        "yoy": growth_yoy(ser, quarters),
        "ctc": growth_ctc(ser, quarters),
    }


# ------------------------------------------------------------------ #
# Rumus 9 baris baku
# ------------------------------------------------------------------ #
def compute_nine_rows(base_series, quarters):
    """
    base_series: dict {row_name: list[float]} untuk 5 baris input, panjang=len(quarters)
    Kembalian: dict {row_name: list[float]} untuk 9 baris baku.
    """
    def get(r):
        return np.array(base_series.get(r, [0.0] * len(quarters)), dtype=float)

    BB = get("Belanja Barang")
    BS = get("Belanja Bansos")
    BP = get("Belanja Pegawai")
    PY = get("Penyusutan")
    PN = get("Penerimaan")

    BA = BB + BS
    NTB = BP + PY
    OUT = BA + NTB
    PKP = OUT - PN

    return {
        "Belanja Barang": BB.tolist(),
        "Belanja Bansos": BS.tolist(),
        "Belanja Pegawai": BP.tolist(),
        "Penyusutan": PY.tolist(),
        "Penerimaan": PN.tolist(),
        "Biaya Antara": BA.tolist(),
        "NTB": NTB.tolist(),
        "Output": OUT.tolist(),
        "PKP": PKP.tolist(),
    }


def nine_rows_as_series(nine, quarters):
    """Ubah dict 9 baris (list) menjadi dict {row: {label: value}}."""
    out = {}
    for r, arr in nine.items():
        out[r] = series_to_dict(arr, quarters)
    return out


# ------------------------------------------------------------------ #
# Konstruksi APBDes dari Data Neraca K3 (kodekab 3209)
# ------------------------------------------------------------------ #
def build_desa_quarterly_from_k3(annual_k3, quarters, spuno=None):
    """
    annual_k3: dict {tahun: dict{komponen: nilai_dalam_RUPIAH}}
        komponen minimal: 'Belanja Pegawai','Belanja Barang','Belanja Modal','Penerimaan'
    Kembalian: base_series triwulanan (Juta Rp) untuk sub-tabel 3.1
        (Belanja Barang, Bansos=0, Pegawai, Penyusutan=20%*Modal, Penerimaan)
    """
    if spuno is None:
        spuno = C.SPUNO_RATIO
    n = len(quarters)
    res = {r: [0.0] * n for r in C.BASE_ROWS}

    # Tahunan -> Juta Rp
    annual_jt = {}
    for y, comp in annual_k3.items():
        annual_jt[y] = {k: (v / 1e6) for k, v in comp.items()}

    idx = C.quarter_index_map(quarters)
    for (y, q) in quarters:
        if y not in annual_jt:
            continue
        a = annual_jt[y]
        ratio = spuno.get(y) or spuno.get(max(k for k in spuno.keys() if k <= y)) \
            if spuno else None
        if ratio is None:
            continue
        i = idx[C.qlabel(y, q)]
        pegawai = a.get("Belanja Pegawai", 0.0) * ratio["Belanja Pegawai"][q - 1]
        barang = a.get("Belanja Barang", 0.0) * ratio["Belanja Barang"][q - 1]
        modal = a.get("Belanja Modal", 0.0) * ratio["Belanja Modal"][q - 1]
        penerimaan = a.get("Penerimaan", 0.0) * ratio["Penerimaan"][q - 1]
        penyusutan = 0.20 * modal  # 20% x Belanja Modal

        res["Belanja Barang"][i] = barang
        res["Belanja Bansos"][i] = 0.0
        res["Belanja Pegawai"][i] = pegawai
        res["Penyusutan"][i] = penyusutan
        res["Penerimaan"][i] = penerimaan
    return res


# ------------------------------------------------------------------ #
# Pemecahan STIK & MBG
# ------------------------------------------------------------------ #
def split_mbg(mbg_total, quarters):
    """
    mbg_total: dict {label: nilai} (Juta Rp, sudah dinormalisasi)
    Kembalian: (pendidikan_series, kesehatan_series) sebagai dict {label: value}
    """
    pend = {}
    kes = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        val = mbg_total.get(lab) or 0.0
        r = C.MBG_RATIO.get((y, q))
        if r is None:
            # default pakai rasio terdekat sebelumnya, fallback 90/10
            r = (0.90, 0.10)
        pend[lab] = val * r[0]
        kes[lab] = val * r[1]
    return pend, kes


def split_stik(stik_total, quarters):
    """
    stik_total: dict {label: nilai} total STIK SSF (Juta Rp)
    Kembalian: (jamkes_series, jkk_series) -> masuk Kesehatan (Individu)
    """
    jam = {}
    jkk = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        val = stik_total.get(lab) or 0.0
        r = C.stik_ratio_for(y, q)
        jam[lab] = val * r[0]
        jkk[lab] = val * r[1]
    return jam, jkk


# ------------------------------------------------------------------ #
# Deflasi ADHB -> ADHK (LK 4)
# ------------------------------------------------------------------ #
def deflate_series(hb_series, deflator_series, quarters):
    """
    hb_series      : dict {label: nilai ADHB}
    deflator_series: dict {label: indeks deflator}
    Kembalian dict {label: ADHK} = ADHB / (deflator/100)
    """
    out = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        hb = hb_series.get(lab)
        d = deflator_series.get(lab)
        if hb is None:
            out[lab] = None
            continue
        if d in (None, 0):
            out[lab] = hb  # deflator kosong -> anggap 100
        else:
            out[lab] = hb / (d / 100.0)
    return out


def build_adhk_from_base(base_series_hb, deflators, quarters):
    """
    base_series_hb : dict {row: dict{label: ADHB}} untuk 5 baris input
    deflators      : dict {nama_deflator: dict{label: indeks}}
    Kembalian dict {row: dict{label: ADHK}} utk 5 baris, lalu hitung 9 baris ADHK.
    """
    hk_base = {}
    for row in C.BASE_ROWS:
        dname = C.DEFLATOR_MAP.get(row, "IHK")
        dser = deflators.get(dname, {})
        hk_base[row] = deflate_series(base_series_hb.get(row, {}), dser, quarters)

    # Susun menjadi list per baris utk compute_nine_rows
    arr = {row: [hk_base[row].get(C.qlabel(y, q), 0.0) for (y, q) in quarters]
           for row in C.BASE_ROWS}
    nine_hk = compute_nine_rows(arr, quarters)
    return nine_rows_as_series(nine_hk, quarters), hk_base


# ------------------------------------------------------------------ #
# Konsolidasi LK 4
# ------------------------------------------------------------------ #
def sum_series_list(list_of_series, quarters):
    """Jumlahkan beberapa series {label:val} menjadi satu series."""
    out = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        tot = 0.0
        for s in list_of_series:
            v = s.get(lab)
            if v is not None:
                tot += v
        out[lab] = tot
    return out


def implicit_index(adhb, adhk, quarters):
    """Indeks Implisit = ADHB/ADHK x 100."""
    out = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        a = adhb.get(lab)
        b = adhk.get(lab)
        if a is not None and b not in (None, 0) and abs(b) > 1e-12:
            out[lab] = (a / b) * 100.0
        else:
            out[lab] = None
    return out


def distribute_pct(part, total, quarters):
    """Distribusi (%) part terhadap total per triwulan."""
    out = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        p = part.get(lab)
        t = total.get(lab)
        if p is not None and t not in (None, 0) and abs(t) > 1e-12:
            out[lab] = (p / t) * 100.0
        else:
            out[lab] = None
    return out


def annual_from_quarterly(ser, quarters):
    """Agregasi triwulanan -> tahunan (jumlah). Kembalian dict {tahun: nilai}."""
    out = {}
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        v = ser.get(lab) or 0.0
        out[y] = out.get(y, 0.0) + v
    return out