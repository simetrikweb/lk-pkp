"""
core/template_io.py (v4)
Template / File Kerja: Home, catatan, 1. APBN Q, 2. APBD Q, 3. Desa Q, 4. SSF Q,
5. STIK, BI Q, 6. APBN A, 7. APBD A, 8. Desa A, 9. SSF A, 10. APBN D, 11. APBD D,
12. Desa D, 13. SSF D, Deflator.
Header blok = "Rincian". Rentang kolom 2019Q1-2030Q4 / 2019-2030.
"""

import io
import json
import re
from datetime import datetime

import pandas as pd

from . import constants as C
from . import parsers as P
from . import pkp_calculator as PC

VERSI_FORMAT = "4.0"

FULL_QUARTERS = C.build_quarters(C.START_YEAR, 2030, 4)
FULL_LABELS = C.qlabels(FULL_QUARTERS)
FULL_YEARS = list(range(C.START_YEAR, 2031))
FULL_YEAR_STRS = [str(y) for y in FULL_YEARS]

APBN_BLOCKS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]
APBD_BLOCKS = ["2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "2.7", "2.8", "2.9"]
SSF_BLOCKS = ["4.1", "4.2", "4.3", "4.4", "4.5", "4.6"]
DESA_BLOCKS = ["3.1"]
BLOCKS = {"apbn": APBN_BLOCKS, "apbd": APBD_BLOCKS,
          "ssf": SSF_BLOCKS, "desa": DESA_BLOCKS}
INPUT_CODES = {
    "apbn": ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6"],
    "apbd": ["2.1", "2.2", "2.3", "2.4", "2.5", "2.6"],
    "ssf": ["4.2", "4.3", "4.4", "4.5", "4.6"],
}
REKAP_MAP = {
    "1.7": ["1.1", "1.2", "1.3"], "1.8": ["1.4", "1.5", "1.6"], "1.9": ["1.7", "1.8"],
    "2.7": ["2.1", "2.2", "2.3"], "2.8": ["2.4", "2.5", "2.6"], "2.9": ["2.7", "2.8"],
    "4.1": ["4.2", "4.3", "4.4", "4.5", "4.6"],
}
SHEET_NAMES = ["Home", "catatan", "1. APBN Q", "2. APBD Q", "3. Desa Q", "4. SSF Q",
               "5. STIK, BI Q", "6. APBN A", "7. APBD A", "8. Desa A", "9. SSF A",
               "10. APBN D", "11. APBD D", "12. Desa D", "13. SSF D", "Deflator"]
QA_SHEETS = [
    ("1. APBN Q", "Q", "apbn"), ("2. APBD Q", "Q", "apbd"),
    ("3. Desa Q", "Q", "desa"), ("4. SSF Q", "Q", "ssf"),
    ("6. APBN A", "A", "apbn"), ("7. APBD A", "A", "apbd"),
    ("8. Desa A", "A", "desa"), ("9. SSF A", "A", "ssf"),
]
D_SHEETS = [("10. APBN D", "apbn"), ("11. APBD D", "apbd"),
            ("12. Desa D", "desa"), ("13. SSF D", "ssf")]
BUILD_ORDER = [
    ("1. APBN Q", "Q", "apbn"), ("2. APBD Q", "Q", "apbd"),
    ("3. Desa Q", "Q", "desa"), ("4. SSF Q", "Q", "ssf"),
    ("5. STIK, BI Q", "S5", None),
    ("6. APBN A", "A", "apbn"), ("7. APBD A", "A", "apbd"),
    ("8. Desa A", "A", "desa"), ("9. SSF A", "A", "ssf"),
    ("10. APBN D", "D", "apbn"), ("11. APBD D", "D", "apbd"),
    ("12. Desa D", "D", "desa"), ("13. SSF D", "D", "ssf"),
    ("Deflator", "DEF", None),
]
S5_BLOCKS = [
    ("5.1", "5.1 STIK MBG Triwulanan", ["MBG"]),
    ("5.2", "5.2 STIK Klaim Jaminan Sosial",
     ["Klaim Jaminan kesehatan",
      "Klaim jaminan kecelakaan kerja berupa barang/jasa"]),
    ("5.3", "5.3 Output BI", ["Output BI Berlaku", "Output BI Konstan"]),
]
CATATAN_LINES = [
    "1. APBD per fungsi menggunakan rasio",
    "2. APDes Tahunan untuk tahun 2024-2025 menggunakan Data Anggaran K3 2025",
    "3. APDes menggunakan rasio SPUN-O untuk triwulannya",
    "4. Data APBN merupakan data alokasi hasil audited sehingga angka indikator = angka Denton",
]
PETUNJUK_LINES = [
    "TEMPLATE / FILE KERJA - LK PKP Rebasing 2025=100 Kab. Cirebon",
    "Struktur mengikuti LK 3 Indikator dan Denton_V1 + sheet Deflator.",
    "Rentang kolom: triwulan 2019Q1-2030Q4; tahun 2019-2030.",
    "",
    "1. Isi HANYA sel angka pada 5 baris input tiap blok: Belanja Barang,",
    "   Belanja Bansos, Belanja pegawai, Penyusutan, Penerimaan.",
    "   Baris Biaya Antara/NTB/Output/PKP serta blok rekap diisi aplikasi.",
    "2. Jangan mengubah nama sheet, nama blok, maupun nama Rincian.",
    "3. Satuan: JUTA RUPIAH. Aplikasi mendeteksi otomatis Rupiah penuh.",
    "4. Penulisan angka bebas: gaya Indonesia (5.976.606,65) atau biasa.",
    "5. Sel kosong dianggap NOL. Nilai minus dilaporkan sebagai anomali SIKD.",
    "6. Sheet 'Deflator' diisi indeks (basis 2025=100), bukan rupiah.",
    "7. Sheet 'catatan' berisi catatan metodologi; baris kosong bebas diisi.",
]
META_KEYS = ["versi_format", "kode_kab", "wilayah", "tahun_awal",
             "triwulan_berjalan", "disimpan_pada"]


def _norm_sheet(name):
    return re.sub(r"\s+", "", str(name).lower())


def _find_sheet(xls, expected):
    target = _norm_sheet(expected)
    for actual in xls.keys():
        if _norm_sheet(actual) == target:
            return actual
    return None


def _is_qlabel(x):
    return isinstance(x, str) and bool(re.fullmatch(r"\d{4}Q[1-4]", x.strip()))


def _is_year(x):
    return bool(re.fullmatch(r"\d{4}", str(x).strip()))


def _periods(row_tail, kind):
    out = []
    for x in row_tail:
        if kind == "Q" and _is_qlabel(x):
            out.append(x.strip())
        elif kind == "A" and _is_year(x):
            out.append(str(x).strip())
    return out


def _cell_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def _sum_nine_lists(list_of_nine):
    out = {}
    for comp in C.STD_ROWS:
        arrs = [n.get(comp) for n in list_of_nine if n and n.get(comp)]
        if not arrs:
            out[comp] = None
        else:
            total = [0.0] * len(arrs[0])
            for a in arrs:
                for i, v in enumerate(a):
                    total[i] += (v or 0.0)
            out[comp] = total
    return out


def _nine_to_series(nine_lists, keys):
    out = {}
    for comp, arr in (nine_lists or {}).items():
        if arr is None:
            out[comp] = {}
        else:
            out[comp] = {keys[i]: arr[i] for i in range(min(len(arr), len(keys)))}
    return out


def _nine_lists_from_quarterly(ds, src, code):
    if src == "desa":
        rows = ds.get("desa_quarterly")
    else:
        rows = ds.get(f"{src}_quarterly", {}).get(code)
    if not rows:
        return None
    return PC.compute_nine_rows(rows, ds["quarters"])


def _nine_lists_for_q(ds, src, code):
    if code in REKAP_MAP:
        members = [_nine_lists_from_quarterly(ds, src, m) for m in REKAP_MAP[code]]
        members = [m for m in members if m]
        return _sum_nine_lists(members) if members else None
    return _nine_lists_from_quarterly(ds, src, code)


def _annual_base(ds, src, code, years):
    if src == "apbn":
        q = ds.get("apbn_quarterly", {}).get(code)
        if not q:
            return None
        base = {}
        for comp in C.BASE_ROWS:
            per_year = {}
            for i, (y, qq) in enumerate(ds["quarters"]):
                per_year[y] = per_year.get(y, 0.0) + (q[comp][i] if i < len(q[comp]) else 0.0)
            base[comp] = [per_year.get(y, 0.0) for y in years]
        return base
    if src == "desa":
        k3 = ds.get("desa_k3", {})
        if not k3:
            return None
        base = {comp: [] for comp in C.BASE_ROWS}
        for y in years:
            a = k3.get(y, {})
            pegawai = a.get("Belanja Pegawai", 0.0) / 1e6
            barang = a.get("Belanja Barang", 0.0) / 1e6
            modal = a.get("Belanja Modal", 0.0) / 1e6
            penerimaan = a.get("Penerimaan", 0.0) / 1e6
            base["Belanja Barang"].append(barang)
            base["Belanja Bansos"].append(0.0)
            base["Belanja Pegawai"].append(pegawai)
            base["Penyusutan"].append(0.20 * modal)
            base["Penerimaan"].append(penerimaan)
        return base
    ann = ds.get(f"{src}_annual", {})
    if not ann.get(code):
        return None
    return {comp: [ann.get(code, {}).get(y, {}).get(comp, 0.0) for y in years]
            for comp in C.BASE_ROWS}


def _nine_lists_for_a(ds, src, code, years):
    if code in REKAP_MAP:
        members = []
        for m in REKAP_MAP[code]:
            b = _annual_base(ds, src, m, years)
            if b:
                members.append(PC.compute_nine_rows(b, years))
        return _sum_nine_lists(members) if members else None
    base = _annual_base(ds, src, code, years)
    if not base:
        return None
    return PC.compute_nine_rows(base, years)


def _series_nine_to_lists(nine_series, col_labels):
    return {comp: [nine_series.get(comp, {}).get(l) for l in col_labels]
            for comp in C.STD_ROWS}


def _block_rows_series(code, col_labels, nine_series, order, tracker=None,
                       pos=None, title=None):
    n = len(col_labels)
    rows = [[title or code] + [None] * n]
    rows.append(["Rincian"] + list(col_labels))
    p = pos if pos is not None else 0
    for comp in order:
        ser = (nine_series or {}).get(comp, {})
        rows.append([comp] + [ser.get(l) for l in col_labels])
        if tracker is not None:
            tracker[(code, comp)] = p + 2
        p += 1
    rows.append([None] * (1 + n))
    return rows, (p + 1)


def deflator_rows(defl, labels):
    rows = [["Deflator"], [None], ["Rincian"] + list(labels)]
    for name in C.DEFLATOR_NAMES:
        ser = (defl or {}).get(name, {})
        rows.append([name] + [ser.get(l) for l in labels])
    return rows


def build_workbook_bytes(dataset, setup, filled=True, result=None):
    if dataset is not None:
        ds_labels = C.qlabels(dataset["quarters"])
        ds_years = sorted({y for (y, q) in dataset["quarters"]})
    else:
        cq = (setup or {}).get("current_quarter") or (2026, 2)
        sy = (setup or {}).get("start_year", C.START_YEAR)
        qq = C.build_quarters(sy, cq[0], cq[1])
        ds_labels = C.qlabels(qq)
        ds_years = sorted({y for (y, q) in qq})
    ds_year_strs = [str(y) for y in ds_years]
    cq = (setup or {}).get("current_quarter") or (2026, 2)
    ds = dataset if (filled and dataset) else None
    result = result or {}

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        home = [["LK PKP Rebasing 2025=100 - Kabupaten Cirebon (3209)"],
                ["File Kerja / Template - struktur LK 3 + Deflator"],
                []]
        home.append(["Kunci", "Nilai"])
        home += [
            ["versi_format", VERSI_FORMAT],
            ["kode_kab", C.KODE_KAB],
            ["wilayah", C.NAMA_KAB],
            ["tahun_awal", C.START_YEAR],
            ["triwulan_berjalan", C.qlabel(*cq)],
            ["disimpan_pada", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ]
        home.append([])
        for line in PETUNJUK_LINES:
            home.append([line])
        pd.DataFrame(home).to_excel(writer, sheet_name="Home",
                                    index=False, header=False)

        cat = list(CATATAN_LINES) + [None] * 15
        pd.DataFrame(cat).to_excel(writer, sheet_name="catatan",
                                   index=False, header=False)

        for sheet, kind, src in BUILD_ORDER:
            if kind == "Q":
                col_labels = FULL_LABELS
                keys = ds_labels
            elif kind == "A":
                col_labels = FULL_YEAR_STRS
                keys = ds_year_strs
            else:
                col_labels = FULL_LABELS
                keys = ds_labels

            if kind in ("Q", "A"):
                rows = []
                for code in BLOCKS[src]:
                    nine = None
                    if ds is not None:
                        if kind == "Q":
                            nine = _nine_lists_for_q(ds, src, code)
                        else:
                            nine = _nine_lists_for_a(ds, src, code, ds_years)
                    ns = _nine_to_series(nine, keys)
                    blk, _ = _block_rows_series(code, col_labels, ns, C.STD_ROWS)
                    rows += blk
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

            elif kind == "S5":
                rows = []
                for code, title, order in S5_BLOCKS:
                    n = len(FULL_LABELS)
                    rows.append([title] + [None] * n)
                    rows.append(["Rincian"] + FULL_LABELS)
                    for comp in order:
                        vals = [None] * n
                        if ds is not None:
                            if comp == "MBG":
                                vals = [ds.get("mbg_total", {}).get(l) for l in FULL_LABELS]
                            elif comp == "Klaim Jaminan kesehatan":
                                vals = [(ds.get("stik_total", {}).get(l, 0.0) or 0.0)
                                        * C.stik_ratio_for(*C.quarter_of_label(l))[0]
                                        if l in ds_labels else None for l in FULL_LABELS]
                            elif comp.startswith("Klaim jaminan kecelakaan"):
                                vals = [(ds.get("stik_total", {}).get(l, 0.0) or 0.0)
                                        * C.stik_ratio_for(*C.quarter_of_label(l))[1]
                                        if l in ds_labels else None for l in FULL_LABELS]
                            elif comp == "Output BI Berlaku":
                                vals = [ds.get("output_bi", {}).get(l) for l in FULL_LABELS]
                            elif comp == "Output BI Konstan":
                                vals = [ds.get("output_bi_konstan", {}).get(l) for l in FULL_LABELS]
                        rows.append([comp] + vals)
                    rows.append([None] * (1 + n))
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

            elif kind == "D":
                rows = []
                for code in BLOCKS[src]:
                    nine_series = None
                    if ds is not None:
                        if src == "apbn":
                            nine_series = _nine_to_series(
                                _nine_lists_for_q(ds, src, code), ds_labels)
                        else:
                            src_map = {"apbd": result.get("apbd_nine_denton", {}),
                                       "desa": result.get("desa_nine", {}),
                                       "ssf": result.get("ssf_nine", {})}[src]
                            if code in REKAP_MAP:
                                members = [src_map.get(m) for m in REKAP_MAP[code]]
                                members = [m for m in members if m]
                                if members:
                                    nine_series = {comp: {
                                        l: sum(m.get(comp, {}).get(l, 0.0) or 0.0
                                               for m in members)
                                        for l in FULL_LABELS} for comp in C.STD_ROWS}
                            elif code in src_map:
                                nine_series = src_map[code]
                    blk, _ = _block_rows_series(code, FULL_LABELS, nine_series,
                                                C.STD_ROWS)
                    rows += blk
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

            elif kind == "DEF":
                rows = deflator_rows(ds.get("deflators") if ds else None, FULL_LABELS)
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                            index=False, header=False)

        for ws in writer.book.worksheets:
            if ws.title in ("Home", "catatan"):
                continue
            for row_cells in ws.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'
            try:
                ws.freeze_panes = "B1"
                ws.column_dimensions["A"].width = 42
            except Exception:
                pass

    return buf.getvalue()


def _empty_parsed():
    return {
        "meta": {},
        "apbn_quarterly": {}, "apbd_quarterly": {}, "ssf_quarterly": {},
        "apbd_annual": {}, "ssf_annual": {},
        "desa_k3": {},
        "mbg_total": {}, "stik_total": {},
        "output_bi": {}, "output_bi_konstan": {},
        "deflators": {},
    }


def _scan_blocks(df, kind, handler):
    current = None
    col_labels = []
    for _, row in df.iterrows():
        s0 = _cell_str(row.iloc[0])
        if s0 == "":
            continue
        m = re.match(r"^(\d+\.\d+)", s0)
        if m and "rincian" not in s0.lower():
            current = m.group(1)
            col_labels = []
            continue
        if "rincian" in s0.lower():
            col_labels = _periods(row.iloc[1:].tolist(), kind)
            continue
        if current and col_labels:
            vals = {}
            for col, v in zip(df.columns[1:], row.iloc[1:].tolist()):
                key = str(col).strip()
                if key in col_labels:
                    vals[key] = P.parse_id_number(v)
            handler(current, s0, vals)


def parse_workbook(file_bytes):
    xls = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None,
                        dtype=object, header=None)
    parsed = _empty_parsed()
    report = {"sheets_read": [], "sheets_missing": [], "units": {},
              "rows_filled": {}, "negatives": [], "unknown_codes": [],
              "meta": {}}
    found = {}
    for name in SHEET_NAMES:
        actual = _find_sheet(xls, name)
        if actual:
            found[name] = actual
            report["sheets_read"].append(name)
        else:
            report["sheets_missing"].append(name)

    if "Home" in found:
        for _, row in xls[found["Home"]].iterrows():
            k = _cell_str(row.iloc[0])
            if k in META_KEYS:
                parsed["meta"][k] = _cell_str(row.iloc[1])
        report["meta"] = parsed["meta"]

    for sheet, kind, src in QA_SHEETS:
        if sheet not in found:
            continue
        df = xls[found[sheet]]
        entries = []

        def handler(code, comp, vals):
            comp_norm = comp.lower()
            if comp_norm not in [b.lower() for b in C.BASE_ROWS]:
                return
            if src != "desa" and code not in INPUT_CODES[src]:
                if code not in BLOCKS[src]:
                    report["unknown_codes"].append(f"{sheet}:{code}")
                return
            if src == "desa" and kind == "Q":
                return
            for lab, v in vals.items():
                if v is not None:
                    entries.append((code, comp_norm, lab, v))

        _scan_blocks(df, kind, handler)
        raws = [e[3] for e in entries]
        unit = P.detect_unit(raws) if raws else "Juta"
        factor = 1e-6 if unit == "Rupiah" else 1.0
        for code, comp_norm, lab, v in entries:
            val = v * factor
            if val < 0:
                report["negatives"].append(
                    {"sheet": sheet, "subtabel": code, "komponen": comp_norm,
                     "periode": lab, "nilai": val})
            if kind == "Q":
                target = parsed[f"{src}_quarterly"].setdefault(code, {})
                target.setdefault(comp_norm, {})[lab] = val
            else:
                y = int(lab)
                if src == "desa":
                    k3 = parsed["desa_k3"].setdefault(y, {})
                    if comp_norm == "belanja pegawai":
                        k3["Belanja Pegawai"] = val * 1e6
                    elif comp_norm == "belanja barang":
                        k3["Belanja Barang"] = val * 1e6
                    elif comp_norm == "penyusutan":
                        k3["Belanja Modal"] = (val / 0.20) * 1e6
                    elif comp_norm == "penerimaan":
                        k3["Penerimaan"] = val * 1e6
                else:
                    target = parsed[f"{src}_annual"].setdefault(code, {})
                    target.setdefault(y, {})[comp_norm] = val
        report["units"][sheet] = unit
        report["rows_filled"][sheet] = len(entries)

    s5 = _find_sheet(xls, "5. STIK, BI Q")
    if s5:
        entries = []

        def handler5(code, comp, vals):
            cl = comp.lower()
            for lab, v in vals.items():
                if v is not None:
                    entries.append((cl, lab, v))

        _scan_blocks(xls[s5], "Q", handler5)
        raws = [e[2] for e in entries]
        unit = P.detect_unit(raws) if raws else "Juta"
        factor = 1e-6 if unit == "Rupiah" else 1.0
        kes = {}
        jkk = {}
        for cl, lab, v in entries:
            val = v * factor
            if val < 0:
                report["negatives"].append(
                    {"sheet": "5. STIK, BI Q", "subtabel": "-",
                     "komponen": cl, "periode": lab, "nilai": val})
            if cl == "mbg":
                parsed["mbg_total"][lab] = val
            elif cl.startswith("klaim jaminan kesehatan"):
                kes[lab] = val
            elif cl.startswith("klaim jaminan kecelakaan"):
                jkk[lab] = val
            elif cl == "output bi berlaku":
                parsed["output_bi"][lab] = val
            elif cl == "output bi konstan":
                parsed["output_bi_konstan"][lab] = val
        for lab in set(list(kes.keys()) + list(jkk.keys())):
            parsed["stik_total"][lab] = kes.get(lab, 0.0) + jkk.get(lab, 0.0)
        report["units"]["5. STIK, BI Q"] = unit
        report["rows_filled"]["5. STIK, BI Q"] = len(entries)

    defs = _find_sheet(xls, "Deflator")
    if defs:
        df = xls[defs]
        col_labels = []
        n = 0
        for _, row in df.iterrows():
            s0 = _cell_str(row.iloc[0])
            if s0 == "":
                continue
            if "rincian" in s0.lower():
                col_labels = _periods(row.iloc[1:].tolist(), "Q")
                continue
            if s0 in C.DEFLATOR_NAMES and col_labels:
                for col, v in zip(df.columns[1:], row.iloc[1:].tolist()):
                    key = str(col).strip()
                    if key in col_labels:
                        pv = P.parse_id_number(v)
                        if pv is not None:
                            parsed["deflators"].setdefault(s0, {})[key] = pv
                            n += 1
        report["units"]["Deflator"] = "Indeks"
        report["rows_filled"]["Deflator"] = n

    return parsed, report


def dataset_from_parsed(parsed, start_year, current_quarter):
    quarters = C.build_quarters(start_year, current_quarter[0], current_quarter[1])
    labels = C.qlabels(quarters)
    years = sorted({y for (y, q) in quarters})
    ds = {"quarters": quarters, "current_quarter": current_quarter}

    for src in ["apbn", "apbd", "ssf"]:
        ds[f"{src}_quarterly"] = {
            code: {comp: [parsed[f"{src}_quarterly"]
                          .get(code, {}).get(comp, {}).get(l, 0.0)
                          for l in labels]
                   for comp in C.BASE_ROWS}
            for code in INPUT_CODES[src]
        }
        ann = {}
        for code, cd in parsed[f"{src}_annual"].items():
            ann[code] = {y: {comp: yd.get(comp, 0.0) for comp in C.BASE_ROWS}
                         for y, yd in cd.items() if y in years}
        ds[f"{src}_annual"] = ann

    ds["desa_k3"] = {y: parsed["desa_k3"][y]
                     for y in years if parsed["desa_k3"].get(y)}
    for key in ["mbg_total", "stik_total", "output_bi", "output_bi_konstan"]:
        ds[key] = {l: parsed[key].get(l, 0.0) for l in labels}

    defl = parsed.get("deflators") or {}
    if defl:
        ds["deflators"] = {name: {l: defl.get(name, {}).get(l, 100.0)
                                  for l in labels}
                           for name in C.DEFLATOR_NAMES}
    else:
        ds["deflators"] = {name: {l: 100.0 for l in labels}
                           for name in C.DEFLATOR_NAMES}
    return ds