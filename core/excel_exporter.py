"""
core/excel_exporter.py (v6)
Workbook final: LK3 (15 sheet, urutan resmi, 5. STIK, BI Q di antara 4 & 6)
+ LK4: 1.ADHB Pusat, 2ab. ADHB Daerah_Final, 3. ADHB SSF, STIK, BI, Deflator,
       1. ADHK Pusat, 2.ADHK Daerah, 3. ADHK SSF, STIK, BI, Rekap Total,
       Rekap Benchmark, Upload Siperi, Upload Verdana.
Sheet 7/8: Growth tahunan (rumus Excel). Sheet 10-13: Growth q-o-q/y-o-y/c-o-c (rumus).
"""

import io

import pandas as pd
from openpyxl.utils import get_column_letter as CL

from . import constants as C
from . import pkp_calculator as PC
from . import template_io as T

FL = T.FULL_LABELS
FQ = T.FULL_QUARTERS
FY = T.FULL_YEAR_STRS
NQ = len(FL)
NY = len(FY)

DEF_ORDER = ["Implisit PMTB", "IHPB Umum tanpa Ekspor", "IHK", "Indeks Upah adj_2"]

ROWSPEC = [
    ("Biaya Antara", "Administrasi+SSF", "ba_adm"),
    ("Biaya Antara", "Pengadaan Air", "ba_air"),
    ("Biaya Antara", "Infokom", "ba_inf"),
    ("Biaya Antara", "Pendidikan", "ba_dik"),
    ("Biaya Antara", "Kesehatan", "ba_kes"),
    ("Biaya Antara", "Hiburan", "ba_hib"),
    ("NTB", "Administrasi+SSF", "ntb_adm"),
    ("NTB", "Pengadaan Air", "ntb_air"),
    ("NTB", "Infokom", "ntb_inf"),
    ("NTB", "Pendidikan", "ntb_dik"),
    ("NTB", "Kesehatan", "ntb_kes"),
    ("NTB", "Hiburan", "ntb_hib"),
    ("Output", "Kolektif - Administrasi", "out_kadm"),
    ("Output", "Kolektif - Pengadaan Air", "out_kair"),
    ("Output", "Kolektif - Infokom", "out_kinf"),
    ("Output", "Individu - Pendidikan", "out_idik"),
    ("Output", "Individu - Kesehatan", "out_ikes"),
    ("Output", "Individu - Hiburan", "out_ihib"),
    ("Penerimaan", "Kolektif", "pen_k"),
    ("Penerimaan", "Individu", "pen_i"),
    ("Output BI", "", "bi"),
    ("STIK MBG", "", "mbg"),
    ("STIK SSF", "", "stik"),
    ("PKP", "Kolektif", "pkp_k"),
    ("PKP", "Individu", "pkp_i"),
]
GROUP_CODES = {
    "adm": ["1.1", "2.1", "4.1"], "air": ["1.2", "2.2"], "inf": ["1.3", "2.3"],
    "dik": ["1.4", "2.4"], "kes": ["1.5", "2.5"], "hib": ["1.6", "2.6"],
}
KOL_CODES = ["1.1", "1.2", "1.3", "2.1", "2.2", "2.3", "4.1"]
IND_CODES = ["1.4", "1.5", "1.6", "2.4", "2.5", "2.6"]


def _ser(arr, labels):
    if arr is None:
        return {l: None for l in labels}
    return {labels[i]: (arr[i] if i < len(arr) else None) for i in range(len(labels))}


def _add(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return [x + y for x, y in zip(a, b)]


def _sub(a, b):
    if a is None or b is None:
        return a
    return [x - y for x, y in zip(a, b)]


def _manual_nine(vals):
    z = [0.0] * len(vals)
    return {"Belanja Barang": list(z), "Belanja Bansos": list(z),
            "Belanja Pegawai": list(z), "Penyusutan": list(z),
            "Penerimaan": list(z), "Biaya Antara": list(z), "NTB": list(z),
            "Output": list(vals), "PKP": list(vals)}


def _block_nines(dataset, result, use_raw=False):
    q = dataset["quarters"]
    n = {}
    for code, rows in dataset.get("apbn_quarterly", {}).items():
        n[code] = PC.compute_nine_rows(rows, q)
    src = result.get("apbd_nine_raw" if use_raw else "apbd_nine_denton") or {}
    for code, rows in dataset.get("apbd_quarterly", {}).items():
        if code in src:
            n[code] = T._series_nine_to_lists(src[code], C.qlabels(q))
        else:
            n[code] = PC.compute_nine_rows(rows, q)
    dn = (result.get("desa_nine") or {}).get("3.1")
    if dn and not use_raw:
        n["3.1"] = T._series_nine_to_lists(dn, C.qlabels(q))
    elif dataset.get("desa_quarterly"):
        n["3.1"] = PC.compute_nine_rows(dataset["desa_quarterly"], q)
    for code, rows in dataset.get("ssf_quarterly", {}).items():
        n[code] = PC.compute_nine_rows(rows, q)
    for rk in T.REKAP_MAP:
        members = [n[m] for m in T.REKAP_MAP[rk] if m in n]
        n[rk] = T._sum_nine_lists(members) if members else None
    return n


def _nine_hk(nine_hb, defl, labels):
    base = {}
    for comp in C.BASE_ROWS:
        dname = C.DEFLATOR_MAP.get(comp, "IHK")
        dser = defl.get(dname, {})
        arr = nine_hb.get(comp) or [0.0] * len(labels)
        base[comp] = [(arr[i] or 0.0) / ((dser.get(labels[i]) or 100.0) / 100.0)
                      for i in range(len(labels))]
    return PC.compute_nine_rows(base, labels)


def _defl_series(arr, dser, labels):
    return [(arr[i] or 0.0) / ((dser.get(labels[i]) or 100.0) / 100.0)
            for i in range(len(labels))]


def _aggregates(nines, extras):
    g = {}

    def rowsum(codes, row):
        out = None
        for c in codes:
            out = _add(out, (nines.get(c) or {}).get(row))
        return out

    for tag, codes in GROUP_CODES.items():
        g["ba_" + tag] = rowsum(codes, "Biaya Antara")
        g["ntb_" + tag] = rowsum(codes, "NTB")
        if tag in ("adm", "air", "inf"):
            g["out_k" + tag] = rowsum(codes, "Output")
        else:
            g["out_i" + tag] = rowsum(codes, "Output")
    g["pen_k"] = rowsum(KOL_CODES, "Penerimaan")
    g["pen_i"] = rowsum(IND_CODES, "Penerimaan")
    g["bi"] = extras.get("bi")
    g["mbg"] = extras.get("mbg")
    g["stik"] = extras.get("stik")
    out_k = _add(_add(g["out_kadm"], g["out_kair"]), g["out_kinf"])
    out_i = _add(_add(g["out_idik"], g["out_ikes"]), g["out_ihib"])
    g["pkp_k"] = _add(_sub(out_k, g["pen_k"]), g["bi"])
    g["pkp_i"] = _add(_add(_sub(out_i, g["pen_i"]), g["mbg"]), g["stik"])
    return g


def _rekap_tables(agg_hb, agg_hk, y0, y1):
    rows = []
    hb = {k: _ser(v, FL) for k, v in agg_hb.items()}
    hk = {k: _ser(v, FL) for k, v in agg_hk.items()}
    impl = {}
    for k in hb:
        impl[k] = {l: (hb[k].get(l) / hk[k].get(l) * 100.0)
                   if hk[k].get(l) else None for l in FL}
    out_total = None
    for k in ["out_kadm", "out_kair", "out_kinf", "out_idik", "out_ikes", "out_ihib"]:
        out_total = _add(out_total, agg_hb.get(k))
    ot = _ser(out_total, FL)

    def table(title, getter):
        rows.append([title])
        rows.append(["Rincian"] + FL)
        for group, label, key in ROWSPEC:
            ser = getter(key)
            rows.append([(group + " " + label).strip()]
                        + [(ser or {}).get(l) for l in FL])
        rows.append([None] * (1 + NQ))

    table(f"Tabel 1:  ADHB Tahun {y0}-{y1}", lambda k: hb.get(k))
    table(f"Tabel 2:  ADHK Tahun {y0}-{y1}", lambda k: hk.get(k))
    table("Tabel 1:  Distribusi ADHB",
          lambda k: {l: (hb.get(k, {}).get(l) / ot.get(l) * 100.0)
                     if ot.get(l) else None for l in FL})
    table(f"Tabel 3:  Y-o-Y Tahun {y0}-{y1}",
          lambda k: PC.growth_yoy(hb.get(k, {}), FQ))
    table(f"Tabel 4:  Q-to-Q Tahun {y0}-{y1}",
          lambda k: PC.growth_qoq(hb.get(k, {}), FQ))
    table(f"Tabel 5:  C-to-C Tahun {y0}-{y1}",
          lambda k: PC.growth_ctc(hb.get(k, {}), FQ))
    table(f"Tabel 6:  Indeks Implisit Tahun {y0}-{y1}", lambda k: impl.get(k))
    table(f"Tabel 7:  Implisit Y-o-Y Tahun {y0}-{y1}",
          lambda k: PC.growth_yoy(impl.get(k, {}), FQ))
    table(f"Tabel 8:  Implisit Q-to-Q Tahun {y0}-{y1}",
          lambda k: PC.growth_qoq(impl.get(k, {}), FQ))
    return rows


def _apply_a_formulas(ws, tracker):
    g0 = 2 + NY + 2
    for (code, comp), rpos in tracker.items():
        R = rpos + 1
        for j in range(NY):
            if j > 0:
                mc = CL(2 + j)
                mp = CL(1 + j)
                cell = ws.cell(row=R, column=g0 + j)
                cell.value = f'=IF({mp}{R}=0,"",({mc}{R}-{mp}{R})/{mp}{R}*100)'
                cell.number_format = '#,##0.00'


def _apply_q_formulas(ws, tracker):
    gap = 2
    q0 = 2 + NQ + gap
    y0 = q0 + NQ + gap
    c0 = y0 + NQ + gap
    for (code, comp), rpos in tracker.items():
        R = rpos + 1
        for i in range(NQ):
            mc = CL(2 + i)
            if i > 0:
                mp = CL(1 + i)
                cell = ws.cell(row=R, column=q0 + i)
                cell.value = f'=IF({mp}{R}=0,"",({mc}{R}-{mp}{R})/{mp}{R}*100)'
                cell.number_format = '#,##0.00'
            if i >= 4:
                m4 = CL(2 + i - 4)
                cell = ws.cell(row=R, column=y0 + i)
                cell.value = f'=IF({m4}{R}=0,"",({mc}{R}-{m4}{R})/{m4}{R}*100)'
                cell.number_format = '#,##0.00'
            yy, qq = FQ[i]
            if yy > C.START_YEAR:
                s = i - (qq - 1)
                cur = f"{CL(2 + s)}{R}:{mc}{R}"
                prv = f"{CL(2 + s - 4)}{R}:{CL(2 + i - 4)}{R}"
                cell = ws.cell(row=R, column=c0 + i)
                cell.value = f'=IF(SUM({prv})=0,"",(SUM({cur})/SUM({prv})-1)*100)'
                cell.number_format = '#,##0.00'


def _write_growth_headers(ws, tracker, kind):
    code_rows = {}
    head_rows = {}
    for (code, comp), rpos in tracker.items():
        code_rows[code] = min(code_rows.get(code, 999999), rpos - 2)
        head_rows[code] = min(head_rows.get(code, 999999), rpos - 1)
    if kind == "A":
        offs = [(2 + NY + 2, "Growth (%)")]
        labs = FY
    else:
        gap = 2
        q0 = 2 + NQ + gap
        y0 = q0 + NQ + gap
        c0 = y0 + NQ + gap
        offs = [(q0, "Growth q to q (%)"), (y0, "Growth y o y (%)"),
                (c0, "Growth c to c (%)")]
        labs = FL
    for code, cr in code_rows.items():
        for off, title in offs:
            ws.cell(row=cr + 1, column=off).value = title
            for i, l in enumerate(labs):
                ws.cell(row=head_rows[code] + 1, column=off + i).value = l


def export_workbook(result, quarters, dataset):
    ds_labels = C.qlabels(dataset["quarters"])
    ds_years = sorted({y for (y, q) in dataset["quarters"]})
    ds_year_strs = [str(y) for y in ds_years]
    defl = dataset.get("deflators", {})
    y0, y1 = C.START_YEAR, 2030

    n_hb = _block_nines(dataset, result, use_raw=False)
    n_raw = _block_nines(dataset, result, use_raw=True)

    mbg = [dataset.get("mbg_total", {}).get(l, 0.0) for l in ds_labels]
    stik = [dataset.get("stik_total", {}).get(l, 0.0) for l in ds_labels]
    bi = [dataset.get("output_bi", {}).get(l, 0.0) for l in ds_labels]
    bik = [dataset.get("output_bi_konstan", {}).get(l, 0.0) for l in ds_labels]
    if all(v == 0 for v in bik):
        bik = list(bi)

    mbg_f = _ser(mbg, ds_labels)
    stik_f = _ser(stik, ds_labels)
    bi_f = _ser(bi, ds_labels)
    bik_f = _ser(bik, ds_labels)

    n_hb["3.1"] = n_hb.get("4.1")
    n_hb["3.2"] = _manual_nine(stik)
    n_hb["3.3"] = _manual_nine(bi)
    n_raw["3.1"] = n_raw.get("4.1")
    n_raw["3.2"] = _manual_nine(stik)
    n_raw["3.3"] = _manual_nine(bi)
    desa = n_hb.get("3.1")
    if desa:
        n_hb["2.1"] = T._sum_nine_lists([x for x in [n_hb.get("2.1"), desa] if x])

    n_hk = {c: _nine_hk(n, defl, ds_labels) for c, n in n_hb.items() if n}
    n_hk["3.2"] = _manual_nine(_defl_series(stik, defl.get("IHK", {}), ds_labels))
    n_hk["3.3"] = _manual_nine(bik)
    n_raw_hk = {c: _nine_hk(n, defl, ds_labels) for c, n in n_raw.items() if n}
    n_raw_hk["3.2"] = n_hk["3.2"]
    n_raw_hk["3.3"] = n_hk["3.3"]

    ex_hb = {"bi": bi, "mbg": mbg, "stik": stik}
    ex_hk = {"bi": bik,
             "mbg": _defl_series(mbg, defl.get("IHK", {}), ds_labels),
             "stik": _defl_series(stik, defl.get("IHK", {}), ds_labels)}
    agg_hb = _aggregates(n_hb, ex_hb)
    agg_hk = _aggregates(n_hk, ex_hk)
    agg_raw_hb = _aggregates(n_raw, ex_hb)
    agg_raw_hk = _aggregates(n_raw_hk, ex_hk)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        home = [["LK PKP Rebasing 2025=100 - Kabupaten Konoha (XXXX)"],
                ["Workbook final: LK 3 + LK 4 Rebasing Regional_V1"],
                ["Rentang: triwulan 2019Q1-2030Q4; tahun 2019-2030"], []]
        pd.DataFrame(home).to_excel(writer, sheet_name="Home",
                                    index=False, header=False)
        cat = list(T.CATATAN_LINES) + [None] * 10
        pd.DataFrame(cat).to_excel(writer, sheet_name="catatan",
                                   index=False, header=False)

        trackers = {}
        for sheet, kind, src in T.BUILD_ORDER:
            if kind == "DEF":
                continue
            if kind in ("Q", "A"):
                col_labels = FL if kind == "Q" else FY
                keys = ds_labels if kind == "Q" else ds_year_strs
                rows = []
                tracker = {}
                pos = 0
                for code in T.BLOCKS[src]:
                    if kind == "Q":
                        nine = T._nine_lists_for_q(dataset, src, code)
                    else:
                        nine = T._nine_lists_for_a(dataset, src, code, ds_years)
                    ns = T._nine_to_series(nine, keys)
                    blk, pos = T._block_rows_series(code, col_labels, ns, C.STD_ROWS,
                                                    tracker=tracker, pos=pos)
                    rows += blk
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)
                trackers[sheet] = (tracker, kind)

            elif kind == "S5":
                rows = []
                for code, title, order in T.S5_BLOCKS:
                    rows.append([title] + [None] * NQ)
                    rows.append(["Rincian"] + FL)
                    for comp in order:
                        if comp == "MBG":
                            src_ser = mbg_f
                        elif comp == "Klaim Jaminan kesehatan":
                            src_ser = {l: (stik_f.get(l) or 0)
                                       * C.stik_ratio_for(*C.quarter_of_label(l))[0]
                                       for l in ds_labels}
                        elif comp.startswith("Klaim jaminan kecelakaan"):
                            src_ser = {l: (stik_f.get(l) or 0)
                                       * C.stik_ratio_for(*C.quarter_of_label(l))[1]
                                       for l in ds_labels}
                        elif comp == "Output BI Berlaku":
                            src_ser = bi_f
                        else:
                            src_ser = bik_f
                        rows.append([comp] + [src_ser.get(l) for l in FL])
                    rows.append([None] * (1 + NQ))
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

            elif kind == "D":
                rows = []
                tracker = {}
                pos = 0
                src_map = {"apbn": None,
                           "apbd": result.get("apbd_nine_denton", {}),
                           "desa": result.get("desa_nine", {}),
                           "ssf": result.get("ssf_nine", {})}[src]
                for code in T.BLOCKS[src]:
                    ns = None
                    if src == "apbn":
                        ns = T._nine_to_series(
                            T._nine_lists_for_q(dataset, src, code), ds_labels)
                    elif code in T.REKAP_MAP:
                        members = [src_map.get(m) for m in T.REKAP_MAP[code]
                                   if m in src_map]
                        if members:
                            ns = {comp: {l: sum(m.get(comp, {}).get(l, 0.0) or 0.0
                                                for m in members) for l in FL}
                                  for comp in C.STD_ROWS}
                    elif code in (src_map or {}):
                        ns = src_map[code]
                    blk, pos = T._block_rows_series(code, FL, ns, C.STD_ROWS,
                                                    tracker=tracker, pos=pos)
                    rows += blk
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)
                trackers[sheet] = (tracker, "Q")

        for sheet, (tracker, kind) in trackers.items():
            ws = writer.book[sheet]
            if sheet in ("7. APBD A", "8. Desa A"):
                _write_growth_headers(ws, tracker, "A")
                _apply_a_formulas(ws, tracker)
            elif sheet in ("10. APBN D", "11. APBD D", "12. Desa D", "13. SSF D"):
                _write_growth_headers(ws, tracker, "Q")
                _apply_q_formulas(ws, tracker)

        def lk4_block(names, nines):
            mapping = [
                (names[0], ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6",
                            "1.7", "1.8", "1.9"]),
                (names[1], ["2.1", "2.2", "2.3", "2.4", "2.5", "2.6",
                            "2.7", "2.8", "2.9"]),
                (names[2], ["3.1", "3.2", "3.3"]),
            ]
            for sheet, codes in mapping:
                rows = []
                for code in codes:
                    ns = T._nine_to_series(nines.get(code), ds_labels)
                    blk, _ = T._block_rows_series(code, FL, ns, C.STD_ROWS)
                    rows += blk
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

        lk4_block(["1.ADHB Pusat", "2ab. ADHB Daerah_Final",
                   "3. ADHB SSF, STIK, BI"], n_hb)
        pd.DataFrame(T.deflator_rows(defl, FL)).to_excel(
            writer, sheet_name="Deflator", index=False, header=False)
        lk4_block(["1. ADHK Pusat", "2.ADHK Daerah",
                   "3. ADHK SSF, STIK, BI"], n_hk)

        pd.DataFrame(_rekap_tables(agg_hb, agg_hk, y0, y1)).to_excel(
            writer, sheet_name="Rekap Total", index=False, header=False)
        pd.DataFrame(_rekap_tables(agg_raw_hb, agg_raw_hk, y0, y1)).to_excel(
            writer, sheet_name="Rekap Benchmark", index=False, header=False)

        sip = []
        total = _add(agg_hb.get("pkp_k"), agg_hb.get("pkp_i"))
        for id_code, ser in [(C.SIPERI_ID["PKP_Total"], _ser(total, FL)),
                             (C.SIPERI_ID["Kolektif"], _ser(agg_hb.get("pkp_k"), FL)),
                             (C.SIPERI_ID["Individu"], _ser(agg_hb.get("pkp_i"), FL))]:
            for l in FL:
                sip.append([id_code, l, ser.get(l)])
        pd.DataFrame(sip, columns=["id", "periode", "nilai"]).to_excel(
            writer, sheet_name="Upload Siperi", index=False)

        verd = [("Sampah/Limbah", "air"), ("Infokom Pemerintah", "inf"),
                ("O Administrasi Pemerintahan", "adm"),
                ("Pendidikan Pemerintah", "dik"),
                ("Kesehatan & Sosial Pemerintah", "kes"),
                ("Jasa Lainnya Pemerintah", "hib")]
        vrows = [["Lapangan Usaha"] + FL]
        for name, tag in verd:
            vrows.append([name] + [_ser(agg_hb.get("ntb_" + tag), FL).get(l)
                                   for l in FL])
        pd.DataFrame(vrows).to_excel(writer, sheet_name="Upload Verdana",
                                     index=False, header=False)

        for ws in writer.book.worksheets:
            if ws.title in ("Home", "catatan"):
                continue
            for row_cells in ws.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

    return buf.getvalue()