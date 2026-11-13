"""
core/template_io.py (v2.1)
File Kerja / Template = struktur resmi LK 3 Indikator dan Denton_V1 (15 sheet).
Deflator dibawa lewat meta 'deflator_json' di sheet Home (tidak menambah sheet).
"""

import io
import json
import re
from datetime import datetime

import pandas as pd

from . import constants as C
from . import parsers as P
from . import pkp_calculator as PC

VERSI_FORMAT = "2.1"

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
               "10. APBN D", "11. APBD D", "12. Desa D", "13. SSF D"]
QA_SHEETS = [
    ("1. APBN Q", "Q", "apbn"), ("2. APBD Q", "Q", "apbd"),
    ("3. Desa Q", "Q", "desa"), ("4. SSF Q", "Q", "ssf"),
    ("6. APBN A", "A", "apbn"), ("7. APBD A", "A", "apbd"),
    ("8. Desa A", "A", "desa"), ("9. SSF A", "A", "ssf"),
]
D_SHEETS = [("10. APBN D", "apbn"), ("11. APBD D", "apbd"),
            ("12. Desa D", "desa"), ("13. SSF D", "ssf")]
S5_BLOCKS = [
    ("5.1", ["MBG"]),
    ("5.2", ["Klaim Jaminan kesehatan",
             "Klaim jaminan kecelakaan kerja berupa barang/jasa"]),
    ("5.3", ["Output BI Berlaku", "Output BI Konstan"]),
]
PETUNJUK_LINES = [
    "TEMPLATE / FILE KERJA - LK PKP Rebasing 2025=100 Kab. Cirebon",
    "Struktur mengikuti LK 3 Indikator dan Denton_V1.",
    "",
    "1. Isi HANYA sel angka pada 5 baris input tiap blok: Belanja Barang,",
    "   Belanja Bansos, Belanja pegawai, Penyusutan, Penerimaan.",
    "   Baris Biaya Antara/NTB/Output/PKP serta blok rekap diisi aplikasi.",
    "2. Jangan mengubah nama sheet, nama blok, maupun nama komponen/rincian.",
    "3. Satuan: JUTA RUPIAH. Aplikasi mendeteksi otomatis bila Anda menulis",
    "   dalam Rupiah penuh dan menormalkannya.",
    "4. Penulisan angka bebas: gaya Indonesia (5.976.606,65) atau biasa.",
    "5. Sel kosong dianggap NOL. Nilai minus dilaporkan sebagai anomali SIKD",
    "   pada pratinjau unggah, sebelum diterapkan.",
    "6. Sheet 'catatan' bebas Anda isi; tidak dibaca aplikasi.",
    "7. Deflator tidak memiliki sheet sendiri (ia bagian LK 4); nilai deflator",
    "   Anda tersimpan otomatis di baris meta sheet Home saat Simpan.",
]
META_KEYS = ["versi_format", "kode_kab", "wilayah", "tahun_awal",
             "triwulan_berjalan", "disimpan_pada", "deflator_json"]


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


def _block_rows(code, col_labels, nine, order):
    n = len(col_labels)
    rows = [[code] + [None] * n]
    rows.append(["komponen/rincian"] + list(col_labels))
    for comp in order:
        vals = (nine or {}).get(comp)
        rows.append([comp] + (list(vals) if vals else [None] * n))
    rows.append([None] * (1 + n))
    return rows


def build_workbook_bytes(dataset, setup, filled=True, result=None):
    if dataset is not None:
        quarters = dataset["quarters"]
    else:
        cq = (setup or {}).get("current_quarter") or (2026, 2)
        sy = (setup or {}).get("start_year", C.START_YEAR)
        quarters = C.build_quarters(sy, cq[0], cq[1])
    labels = C.qlabels(quarters)
    years = sorted({y for (y, q) in quarters})
    year_strs = [str(y) for y in years]
    cq = (setup or {}).get("current_quarter") or quarters[-1]
    ds = dataset if (filled and dataset) else None
    result = result or {}

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        home = [["LK PKP Rebasing 2025=100 - Kabupaten Cirebon (3209)"],
                ["File Kerja / Template - struktur LK 3 Indikator dan Denton_V1"],
                []]
        home.append(["Kunci", "Nilai"])
        home += [
            ["versi_format", VERSI_FORMAT],
            ["kode_kab", C.KODE_KAB],
            ["wilayah", C.NAMA_KAB],
            ["tahun_awal", quarters[0][0]],
            ["triwulan_berjalan", C.qlabel(*cq)],
            ["disimpan_pada", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ]
        if ds is not None:
            home.append(["deflator_json", json.dumps(ds.get("deflators", {}))])
        home.append([])
        for line in PETUNJUK_LINES:
            home.append([line])
        pd.DataFrame(home).to_excel(writer, sheet_name="Home",
                                    index=False, header=False)

        cat = [["Catatan pengguna (bebas; tidak dibaca aplikasi)"]]
        for _ in range(15):
            cat.append([None])
        pd.DataFrame(cat).to_excel(writer, sheet_name="catatan",
                                   index=False, header=False)

        for sheet, kind, src in QA_SHEETS:
            col_labels = labels if kind == "Q" else year_strs
            rows = []
            for code in BLOCKS[src]:
                nine = None
                if ds is not None:
                    if kind == "Q":
                        nine = _nine_lists_for_q(ds, src, code)
                    else:
                        nine = _nine_lists_for_a(ds, src, code, years)
                rows += _block_rows(code, col_labels, nine, C.STD_ROWS)
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                        index=False, header=False)

        rows = []
        for code, order in S5_BLOCKS:
            n = len(labels)
            rows.append([code] + [None] * n)
            rows.append(["Rincian"] + labels)
            for comp in order:
                vals = [None] * n
                if ds is not None:
                    if comp == "MBG":
                        vals = [ds.get("mbg_total", {}).get(l) for l in labels]
                    elif comp == "Klaim Jaminan kesehatan":
                        vals = [(ds.get("stik_total", {}).get(l, 0.0) or 0.0)
                                * C.stik_ratio_for(*C.quarter_of_label(l))[0]
                                for l in labels]
                    elif comp.startswith("Klaim jaminan kecelakaan"):
                        vals = [(ds.get("stik_total", {}).get(l, 0.0) or 0.0)
                                * C.stik_ratio_for(*C.quarter_of_label(l))[1]
                                for l in labels]
                    elif comp == "Output BI Berlaku":
                        vals = [ds.get("output_bi", {}).get(l) for l in labels]
                    elif comp == "Output BI Konstan":
                        vals = [ds.get("output_bi_konstan", {}).get(l) for l in labels]
                rows.append([comp] + vals)
            rows.append([None] * (1 + n))
        pd.DataFrame(rows).to_excel(writer, sheet_name="5. STIK, BI Q",
                                    index=False, header=False)

        for sheet, src in D_SHEETS:
            rows = []
            for code in BLOCKS[src]:
                nine = None
                if ds is not None:
                    if src == "apbn":
                        nine = _nine_lists_for_q(ds, src, code)
                    elif src == "apbd":
                        src_map = result.get("apbd_nine_denton", {})
                        if code in REKAP_MAP:
                            members = [_series_nine_to_lists(src_map[m], labels)
                                       for m in REKAP_MAP[code] if m in src_map]
                            nine = _sum_nine_lists(members) if members else None
                        elif code in src_map:
                            nine = _series_nine_to_lists(src_map[code], labels)
                    elif src == "desa":
                        src_map = result.get("desa_nine", {})
                        if code in src_map:
                            nine = _series_nine_to_lists(src_map[code], labels)
                    elif src == "ssf":
                        src_map = result.get("ssf_nine", {})
                        if code in src_map:
                            nine = _series_nine_to_lists(src_map[code], labels)
                rows += _block_rows(code, labels, nine, C.STD_ROWS)
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
        if re.fullmatch(r"\d+\.\d+", s0):
            current = s0
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
    report["sheets_read"] = [n for n in SHEET_NAMES if n in xls]
    report["sheets_missing"] = [n for n in SHEET_NAMES if n not in xls]

    if "Home" in xls:
        for _, row in xls["Home"].iterrows():
            k = _cell_str(row.iloc[0])
            if k == "deflator_json":
                try:
                    parsed["deflators"] = json.loads(_cell_str(row.iloc[1]) or "{}")
                except Exception:
                    parsed["deflators"] = {}
            elif k in META_KEYS:
                parsed["meta"][k] = _cell_str(row.iloc[1])
        report["meta"] = parsed["meta"]

    for sheet, kind, src in QA_SHEETS:
        if sheet not in xls:
            continue
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

        _scan_blocks(xls[sheet], kind, handler)
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

    if "5. STIK, BI Q" in xls:
        entries = []

        def handler5(code, comp, vals):
            cl = comp.lower()
            for lab, v in vals.items():
                if v is not None:
                    entries.append((cl, lab, v))

        _scan_blocks(xls["5. STIK, BI Q"], "Q", handler5)
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