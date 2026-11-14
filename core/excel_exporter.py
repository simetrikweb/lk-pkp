"""
core/excel_exporter.py (v4)
Workbook final = LK3 (15 sheet) + LK4 Rebasing Regional_V1 (11 sheet)
+ Upload Siperi + Upload Verdana. Angka berformat '#,##0.00'.
"""

import io

import pandas as pd

from . import constants as C
from . import parsers as P
from . import pkp_calculator as PC
from . import template_io as T

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
    return {labels[i]: (arr[i] if arr and i < len(arr) else None)
            for i in range(len(labels))}


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


def _block_nines(dataset, result, labels, use_raw=False):
    q = dataset["quarters"]
    n = {}
    for code, rows in dataset.get("apbn_quarterly", {}).items():
        n[code] = PC.compute_nine_rows(rows, q)
    src = result.get("apbd_nine_raw" if use_raw else "apbd_nine_denton") or {}
    for code, rows in dataset.get("apbd_quarterly", {}).items():
        if code in src:
            n[code] = T._series_nine_to_lists(src[code], labels)
        else:
            n[code] = PC.compute_nine_rows(rows, q)
    dn = (result.get("desa_nine") or {}).get("3.1")
    if dn and not use_raw:
        n["3.1"] = T._series_nine_to_lists(dn, labels)
    elif dataset.get("desa_quarterly"):
        n["3.1"] = PC.compute_nine_rows(dataset["desa_quarterly"], q)
    for code, rows in dataset.get("ssf_quarterly", {}).items():
        n[code] = PC.compute_nine_rows(rows, q)
    for rk in ["1.7", "1.8", "1.9", "2.7", "2.8", "2.9", "4.1"]:
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
            arr = (nines.get(c) or {}).get(row)
            out = _add(out, arr)
        return out

    for tag, codes in GROUP_CODES.items():
        g["ba_" + tag] = rowsum(codes, "Biaya Antara")
        g["ntb_" + tag] = rowsum(codes, "NTB")
    g["out_kadm"] = rowsum(GROUP_CODES["adm"], "Output")
    g["out_kair"] = rowsum(GROUP_CODES["air"], "Output")
    g["out_kinf"] = rowsum(GROUP_CODES["inf"], "Output")
    g["out_idik"] = rowsum(GROUP_CODES["dik"], "Output")
    g["out_ikes"] = rowsum(GROUP_CODES["kes"], "Output")
    g["out_ihib"] = rowsum(GROUP_CODES["hib"], "Output")
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


def _rekap_tables(agg_hb, agg_hk, labels, quarters, y0, y1):
    rows = []
    hb = {k: _ser(v, labels) for k, v in agg_hb.items() if v}
    hk = {k: _ser(v, labels) for k, v in agg_hk.items() if v}
    impl = {}
    for k in hb:
        if k in hk:
            impl[k] = {l: (hb[k].get(l) / hk[k].get(l) * 100.0)
                       if hk[k].get(l) else None for l in labels}
    out_total = None
    for k in ["out_kadm", "out_kair", "out_kinf", "out_idik", "out_ikes", "out_ihib"]:
        out_total = _add(out_total, agg_hb.get(k))
    out_total_ser = _ser(out_total, labels)

    def table(title, getter):
        rows.append([title])
        rows.append(["Rincian"] + labels)
        for group, label, key in ROWSPEC:
            ser = getter(key)
            name = (group + " " + label).strip()
            rows.append([name] + [(ser or {}).get(l) for l in labels])
        rows.append([None] * (1 + len(labels)))

    table(f"Tabel 1:  ADHB Tahun {y0}-{y1}", lambda k: hb.get(k))
    table(f"Tabel 2:  ADHK Tahun {y0}-{y1}", lambda k: hk.get(k))
    table("Tabel 1:  Distribusi ADHB",
          lambda k: {l: (hb.get(k, {}).get(l) / out_total_ser.get(l) * 100.0)
                     if out_total_ser.get(l) else None for l in labels})
    table(f"Tabel 3:  Y-o-Y Tahun {y0}-{y1}",
          lambda k: PC.growth_yoy(hb.get(k, {}), quarters))
    table(f"Tabel 4:  Q-to-Q Tahun {y0}-{y1}",
          lambda k: PC.growth_qoq(hb.get(k, {}), quarters))
    table(f"Tabel 5:  C-to-C Tahun {y0}-{y1}",
          lambda k: PC.growth_ctc(hb.get(k, {}), quarters))
    table(f"Tabel 6:  Indeks Implisit Tahun {y0}-{y1}", lambda k: impl.get(k))
    table(f"Tabel 7:  Implisit Y-o-Y Tahun {y0}-{y1}",
          lambda k: PC.growth_yoy(impl.get(k, {}), quarters))
    table(f"Tabel 8:  Implisit Q-to-Q Tahun {y0}-{y1}",
          lambda k: PC.growth_qoq(impl.get(k, {}), quarters))
    return rows


def export_workbook(result, quarters, dataset):
    labels = C.qlabels(quarters)
    years = sorted({y for (y, q) in quarters})
    y0, y1 = years[0], years[-1]
    defl = dataset.get("deflators", {})

    n_hb = _block_nines(dataset, result, labels, use_raw=False)
    n_raw = _block_nines(dataset, result, labels, use_raw=True)

    mbg = [dataset.get("mbg_total", {}).get(l, 0.0) for l in labels]
    stik = [dataset.get("stik_total", {}).get(l, 0.0) for l in labels]
    bi = [dataset.get("output_bi", {}).get(l, 0.0) for l in labels]
    bik = [dataset.get("output_bi_konstan", {}).get(l, 0.0) for l in labels]
    if all(v == 0 for v in bik):
        bik = list(bi)

    n_hb["3.1"] = n_hb.get("4.1")
    n_hb["3.2"] = _manual_nine(stik)
    n_hb["3.3"] = _manual_nine(bi)
    n_raw["3.1"] = n_raw.get("4.1")
    n_raw["3.2"] = _manual_nine(stik)
    n_raw["3.3"] = _manual_nine(bi)

    n_hk = {c: _nine_hk(n, defl, labels) for c, n in n_hb.items() if n}
    n_hk["3.2"] = _manual_nine(_defl_series(stik, defl.get("IHK", {}), labels))
    n_hk["3.3"] = _manual_nine(bik)

    ex_hb = {"bi": bi, "mbg": mbg, "stik": stik}
    ex_hk = {"bi": bik,
             "mbg": _defl_series(mbg, defl.get("IHK", {}), labels),
             "stik": _defl_series(stik, defl.get("IHK", {}), labels)}
    agg_hb = _aggregates(n_hb, ex_hb)
    agg_hk = _aggregates(n_hk, ex_hk)
    agg_raw_hb = _aggregates(n_raw, ex_hb)
    agg_raw_hk = {c: _nine_hk(n, defl, labels) and None for c in []}
    n_raw_hk = {c: _nine_hk(n, defl, labels) for c, n in n_raw.items() if n}
    n_raw_hk["3.2"] = n_hk["3.2"]
    n_raw_hk["3.3"] = n_hk["3.3"]
    agg_raw_hk = _aggregates(n_raw_hk, ex_hk)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        home = [["LK PKP Rebasing 2025=100 - Kabupaten Konoha (XXXX)"],
                ["Workbook final: LK 3 + LK 4 Rebasing Regional_V1"], []]
        pd.DataFrame(home).to_excel(writer, sheet_name="Home",
                                    index=False, header=False)
        cat = [["Catatan"]] + [[None] for _ in range(10)]
        pd.DataFrame(cat).to_excel(writer, sheet_name="catatan",
                                   index=False, header=False)

        # ---------- LK 3 ---------- #
        for sheet, kind, src in T.QA_SHEETS:
            col_labels = labels if kind == "Q" else [str(y) for y in years]
            rows = []
            for code in T.BLOCKS[src]:
                if kind == "Q":
                    nine = T._nine_lists_for_q(dataset, src, code)
                else:
                    nine = T._nine_lists_for_a(dataset, src, code, years)
                rows += T._block_rows(code, col_labels, nine, C.STD_ROWS)
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                        index=False, header=False)
        rows = []
        for code, order in T.S5_BLOCKS:
            n = len(labels)
            rows.append([code] + [None] * n)
            rows.append(["Rincian"] + labels)
            for comp in order:
                if comp == "MBG":
                    vals = mbg
                elif comp == "Klaim Jaminan kesehatan":
                    vals = [stik[i] * C.stik_ratio_for(*C.quarter_of_label(labels[i]))[0]
                            for i in range(n)]
                elif comp.startswith("Klaim jaminan kecelakaan"):
                    vals = [stik[i] * C.stik_ratio_for(*C.quarter_of_label(labels[i]))[1]
                            for i in range(n)]
                elif comp == "Output BI Berlaku":
                    vals = bi
                else:
                    vals = bik
                rows.append([comp] + vals)
            rows.append([None] * (1 + n))
        pd.DataFrame(rows).to_excel(writer, sheet_name="5. STIK, BI Q",
                                    index=False, header=False)
        for sheet, src in T.D_SHEETS:
            rows = []
            src_map = {"apbn": None,
                       "apbd": result.get("apbd_nine_denton", {}),
                       "desa": result.get("desa_nine", {}),
                       "ssf": result.get("ssf_nine", {})}[src]
            for code in T.BLOCKS[src]:
                if src == "apbn":
                    nine = T._nine_lists_for_q(dataset, src, code)
                elif code in T.REKAP_MAP:
                    members = [T._series_nine_to_lists(src_map[m], labels)
                               for m in T.REKAP_MAP[code] if m in src_map]
                    nine = T._sum_nine_lists(members) if members else None
                elif code in (src_map or {}):
                    nine = T._series_nine_to_lists(src_map[code], labels)
                else:
                    nine = None
                rows += T._block_rows(code, labels, nine, C.STD_ROWS)
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                        index=False, header=False)

        # ---------- LK 4 ---------- #
        def adhb_sheets(prefix, nines):
            mapping = [
                (f"{prefix}Pusat", ["1.1", "1.2", "1.3", "1.4", "1.5",
                                    "1.6", "1.7", "1.8", "1.9"]),
                (f"{prefix}Daerah", ["2.1", "2.2", "2.3", "2.4", "2.5",
                                     "2.6", "2.7", "2.8", "2.9"]),
                (f"{prefix}SSF,STIK,BI", ["3.1", "3.2", "3.3"]),
            ]
            for sheet, codes in mapping:
                rows = []
                for code in codes:
                    nine = nines.get(code)
                    if code == "2.1":
                        nine = T._sum_nine_lists([x for x in
                                                  [nine, nines.get("3.1x")] if x])
                    rows += T._block_rows(code, labels, nine, C.STD_ROWS)
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

        # Daerah 2.1 = APBD 2.1 + Desa 3.1
        for nm in ["n_hb", "n_hk", "n_raw", "n_raw_hk"]:
            pass
        n_hb["2.1"] = T._sum_nine_lists([n_hb.get("2.1"), n_hb.get("3.1d")]) \
            if n_hb.get("3.1d") else n_hb.get("2.1")
        # (desa sudah masuk lewat 4.1? tidak) -> tambahkan desa ke 2.1:
        desa_nine = n_hb.get("3.1")
        n_hb["2.1"] = T._sum_nine_lists([x for x in
                                         [(_block_nines_desa(dataset, result, labels)),
                                          n_hb.get("2.1")] if x])

        adhb_sheets("1.ADHB ", n_hb)
        adhb_sheets("2.ADHK ", n_hk)

        rows = [[d] + [defl.get(d, {}).get(l) for l in labels] for d in DEF_ORDER]
        pd.DataFrame(rows).to_excel(writer, sheet_name="Deflator",
                                    index=False, header=False)

        pd.DataFrame(_rekap_tables(agg_hb, agg_hk, labels, quarters, y0, y1)
                     ).to_excel(writer, sheet_name="Rekap Total",
                                index=False, header=False)
        pd.DataFrame(_rekap_tables(agg_raw_hb, agg_raw_hk, labels, quarters, y0, y1)
                     ).to_excel(writer, sheet_name="Rekap Benchmark",
                                index=False, header=False)

        # ---------- Siperi & Verdana ---------- #
        sip = []
        total = _add(agg_hb.get("pkp_k"), agg_hb.get("pkp_i"))
        for id_code, ser in [(C.SIPERI_ID["PKP_Total"], _ser(total, labels)),
                             (C.SIPERI_ID["Kolektif"], _ser(agg_hb.get("pkp_k"), labels)),
                             (C.SIPERI_ID["Individu"], _ser(agg_hb.get("pkp_i"), labels))]:
            for l in labels:
                sip.append([id_code, l, ser.get(l)])
        pd.DataFrame(sip, columns=["id", "periode", "nilai"]).to_excel(
            writer, sheet_name="Upload Siperi", index=False)

        verd = [
            ("Sampah/Limbah", "air"), ("Infokom Pemerintah", "inf"),
            ("O Administrasi Pemerintahan", "adm"),
            ("Pendidikan Pemerintah", "dik"),
            ("Kesehatan & Sosial Pemerintah", "kes"),
            ("Jasa Lainnya Pemerintah", "hib"),
        ]
        vrows = [["Lapangan Usaha"] + labels]
        for name, tag in verd:
            vrows.append([name] + [( _ser(agg_hb.get("ntb_" + tag), labels) or {}).get(l)
                                   for l in labels])
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


def _block_nines_desa(dataset, result, labels):
    dn = (result.get("desa_nine") or {}).get("3.1")
    if dn:
        return T._series_nine_to_lists(dn, labels)
    if dataset.get("desa_quarterly"):
        return PC.compute_nine_rows(dataset["desa_quarterly"], dataset["quarters"])
    return None