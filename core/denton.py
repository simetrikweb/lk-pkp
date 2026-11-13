"""
core/denton.py
Mesin benchmarking Proportional Denton-Cholette (first-difference).

Menggantikan add-in Excel =xlpbm(). Menangani 3 kasus NOTES DENTON LK 3:
  Kasus 1: baris kosong/nol total  -> kembalikan deret nol (cegah #VALUE!)
  Kasus 2: ada triwulan 0 padahal total tahunan > 0 -> imputasi 1e-6
  Kasus 3: data baru mulai tahun tertentu -> Denton hanya pada rentang aktif

Ditambah ekstrapolasi tahun berjalan (melanjutkan rasio Benchmark/Indikator terakhir).
"""

import numpy as np

from . import constants as C


def _first_difference_matrix(n):
    """Matriks D (n-1 x n) sehingga D f = selisih pertama f."""
    D = np.zeros((n - 1, n))
    for t in range(n - 1):
        D[t, t] = -1.0
        D[t, t + 1] = 1.0
    return D


def _solve_proportional(ind, year_of, target_years, annual):
    """Selesaikan sistem KKT Denton proporsional. Kembalian deret benchmarked."""
    n = len(ind)
    D = _first_difference_matrix(n)
    DtD = D.T @ D

    m = len(target_years)
    K = np.zeros((n + m, n + m))
    K[:n, :n] = DtD

    rhs = np.zeros(n + m)
    for j, y in enumerate(target_years):
        for t in range(n):
            if year_of[t] == y:
                K[n + j, t] = ind[t]
                K[t, n + j] = ind[t]
        rhs[n + j] = annual[y]

    try:
        sol = np.linalg.solve(K, rhs)
    except np.linalg.LinAlgError:
        sol, *_ = np.linalg.lstsq(K, rhs, rcond=None)
    f = sol[:n]
    return ind * f


def detect_active_range(ind, annual_years):
    """Kasus 3: temukan indeks awal aktif (triwulan pertama berisi)."""
    n = len(ind)
    start = None
    for t in range(n):
        if ind[t] is not None and not np.isnan(ind[t]) and abs(ind[t]) > 1e-12:
            start = t
            break
    if start is None and annual_years:
        # mulai dari triwulan pertama tahun benchmark terawal
        start = 0
    return start if start is not None else 0


def denton_benchmark(ind_q, quarters, annual, use_raw_for_forecast=False,
                     series_name=""):
    """
    Jalankan Proportional Denton-Cholette.

    Parameter
    ----------
    ind_q   : list[float] nilai indikator triwulanan (panjang = len(quarters))
    quarters: list[(tahun, triwulan)] selaras dengan ind_q
    annual  : dict {tahun: total_tahunan} hanya untuk tahun lengkap yang diisi
    use_raw_for_forecast : bila True, triwulan non-benchmark memakai indikator mentah

    Kembalian
    -------
    (bench, log) : bench list[float], log list[str]
    """
    log = []
    n = len(quarters)
    ind = np.array([np.nan if v is None else float(v) for v in ind_q], dtype=float)

    # Isi NaN dengan 0 untuk kalkulasi
    ind = np.where(np.isnan(ind), 0.0, ind)

    # ---------------- Kasus 1: kosong / nol total ---------------- #
    total_annual = sum(v for v in annual.values() if v is not None)
    if np.all(np.abs(ind) < 1e-12) and abs(total_annual) < 1e-12:
        log.append(f"[Kasus 1] '{series_name}' kosong/nol total -> Denton dilewati, deret nol.")
        return [0.0] * n, log

    # ---------------- Kasus 3: deteksi rentang aktif ---------------- #
    start = detect_active_range(ind, list(annual.keys()))
    if start > 0:
        log.append(f"[Kasus 3] '{series_name}' aktif mulai triwulan "
                   f"{C.qlabel(*quarters[start])} -> Denton pada rentang aktif.")

    # ---------------- Kasus 2: imputasi triwulan 0 ---------------- #
    imputed_pos = []
    ind_work = ind.copy()
    for t in range(start, n):
        y = quarters[t][0]
        if abs(ind_work[t]) < 1e-12 and y in annual and (annual[y] or 0) > 0:
            ind_work[t] = C.DENTON_ZERO_IMPUTE
            imputed_pos.append(C.qlabel(*quarters[t]))
    if imputed_pos:
        log.append(f"[Kasus 2] '{series_name}' imputasi 1e-6 pada {', '.join(imputed_pos)} "
                   f"(total tahunan > 0).")

    # ---------------- Tentukan tahun benchmark (lengkap & terisi) ---------------- #
    year_of = [quarters[t][0] for t in range(n)]
    qcount = {}
    for t in range(n):
        qcount[year_of[t]] = qcount.get(year_of[t], 0) + 1
    # Tahun target: lengkap (4 triwulan dalam rentang) dan ada di annual
    target_years = [y for y in sorted(annual.keys())
                    if qcount.get(y, 0) == 4 and (annual[y] is not None)]

    # ---------------- Jalankan Denton pada rentang aktif ---------------- #
    seg_ind = ind_work[start:]
    seg_years = year_of[start:]
    seg_target = [y for y in target_years if y in set(seg_years)]

    bench_seg = seg_ind.copy()
    if len(seg_ind) > 0 and seg_target:
        seg_annual = {y: annual[y] for y in seg_target}
        bench_seg = _solve_proportional(seg_ind, seg_years, seg_target, seg_annual)

    bench = np.zeros(n)
    bench[start:] = bench_seg

    # ---------------- Ekstrapolasi tahun berjalan ---------------- #
    # triwulan setelah triwulan ber-benchmark terakhir
    last_bench_t = None
    for t in range(n - 1, -1, -1):
        if year_of[t] in target_years:
            last_bench_t = t
            break
    if last_bench_t is not None:
        ratio_last = (bench[last_bench_t] / ind_work[last_bench_t]) \
            if abs(ind_work[last_bench_t]) > 1e-12 else 1.0
        for t in range(last_bench_t + 1, n):
            if year_of[t] not in target_years:
                if use_raw_for_forecast:
                    bench[t] = ind_work[t]
                else:
                    bench[t] = ind_work[t] * ratio_last
        if not use_raw_for_forecast:
            log.append(f"[Forecast] '{series_name}' triwulan non-benchmark dilanjutkan "
                       f"dengan rasio B/I terakhir {ratio_last:.6f}.")
    else:
        # tidak ada benchmark sama sekali -> pakai indikator
        bench = ind_work.copy()
        log.append(f"[Info] '{series_name}' tanpa total tahunan -> pakai indikator mentah.")

    return [float(v) for v in bench], log


def verify_benchmark(bench, quarters, annual, tol=1e-6):
    """Verifikasi jumlah triwulan per tahun == total tahunan (selisih ~ 0)."""
    results = []
    sums = {}
    for t, (y, q) in enumerate(quarters):
        sums[y] = sums.get(y, 0.0) + bench[t]
    for y, target in annual.items():
        if target is None:
            continue
        diff = sums.get(y, 0.0) - target
        results.append({
            "tahun": y,
            "jumlah_triwulanan": sums.get(y, 0.0),
            "target_tahunan": target,
            "selisih": diff,
            "ok": abs(diff) <= tol,
        })
    return results