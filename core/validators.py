"""
core/validators.py
Audit validitas ramah pemula sebelum ekspor.
"""

from . import constants as C


def check_base_year_consistency(adhb, adhk, quarters, tol_pct=1.0):
    """Cek ADHB ~ ADHK pada tahun dasar 2025 (deflator ~ 100)."""
    issues = []
    for (y, q) in quarters:
        if y != C.BASE_YEAR:
            continue
        lab = C.qlabel(y, q)
        a = adhb.get(lab)
        b = adhk.get(lab)
        if a in (None,) or b in (None, 0):
            continue
        diff_pct = abs(a - b) / abs(b) * 100.0 if abs(b) > 1e-12 else 0.0
        if diff_pct > tol_pct:
            issues.append(
                f"[Tahun Dasar] {lab}: ADHB vs ADHK selisih {diff_pct:.2f}% "
                f"(>{tol_pct}%). Periksa deflator 2025 (harus ~100)."
            )
    return issues


def check_extreme_jumps(ser, quarters, threshold_pct=50.0, label=""):
    """Deteksi lonjakan ekstrem y-on-y > +/- threshold (mis. efek MBG / sel minus)."""
    issues = []
    for (y, q) in quarters:
        lab = C.qlabel(y, q)
        prev_lab = C.qlabel(y - 1, q)
        cur = ser.get(lab)
        prev = ser.get(prev_lab)
        if cur is None or prev in (None, 0):
            continue
        g = (cur / prev - 1.0) * 100.0
        if abs(g) > threshold_pct:
            issues.append(
                f"[Lonjakan {label}] {lab}: YoY {g:+.1f}% (> {threshold_pct}%). "
                f"Cek apakah karena MBG mulai berjalan atau sel minus APBD."
            )
    return issues


def check_current_quarter_completeness(base_series, quarters, current_quarter):
    """Cek kelengkapan data pada triwulan berjalan."""
    issues = []
    cy, cq = current_quarter
    lab = C.qlabel(cy, cq)
    all_zero = True
    for row, arr in base_series.items():
        # arr dapat berupa dict label->val atau list
        if isinstance(arr, dict):
            v = arr.get(lab)
        else:
            try:
                i = C.qlabels(quarters).index(lab)
                v = arr[i]
            except (ValueError, IndexError):
                v = None
        if v not in (None, 0) and abs(v) > 1e-12:
            all_zero = False
            break
    if all_zero:
        issues.append(
            f"[Kelengkapan] Triwulan berjalan {lab} seluruhnya nol/kosong. "
            f"Pastikan data triwulan berjalan sudah dimasukkan."
        )
    return issues


def check_denton_residuals(verify_results):
    """Pastikan selisih jumlah triwulan vs tahunan ~ 0 untuk tiap tahun."""
    issues = []
    for r in verify_results:
        if not r["ok"]:
            issues.append(
                f"[Denton] Tahun {r['tahun']}: selisih {r['selisih']:.6f} "
                f"(jumlah triwulan {r['jumlah_triwulanan']:.3f} vs "
                f"target {r['target_tahunan']:.3f})."
            )
    return issues


def run_full_audit(audit_ctx):
    """
    audit_ctx: dict berisi series-series yang dibutuhkan.
    Kembalian list str temuan (kosong berarti lolos).
    """
    issues = []
    quarters = audit_ctx["quarters"]

    if "adhb_total" in audit_ctx and "adhk_total" in audit_ctx:
        issues += check_base_year_consistency(
            audit_ctx["adhb_total"], audit_ctx["adhk_total"], quarters
        )

    if "adhb_total" in audit_ctx:
        issues += check_extreme_jumps(
            audit_ctx["adhb_total"], quarters,
            label="ADHB Total"
        )

    if "current_quarter" in audit_ctx and "base_series" in audit_ctx:
        issues += check_current_quarter_completeness(
            audit_ctx["base_series"], quarters, audit_ctx["current_quarter"]
        )

    if "denton_verify" in audit_ctx:
        issues += check_denton_residuals(audit_ctx["denton_verify"])

    return issues