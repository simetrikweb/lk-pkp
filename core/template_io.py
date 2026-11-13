"""
core/template_io.py
File Kerja Excel (template kosong SEKALIGUS simpanan pekerjaan).
Layout Opsi A: satu sheet per sumber dana dengan kolom SubTabel.

Sheet: Petunjuk, Meta, APBN Q, APBD Q, APBD A, Desa K3,
       SSF Q, SSF A, MBG, STIK, Output BI, Deflator.

Kaidah angka: sel numerik diberi number_format '#,##0.00' sehingga di
Excel berbahasa Indonesia tampil sebagai 5.976.606,65. Pengguna juga
boleh mengetik teks gaya Indonesia; pembaca memakai parse_id_number.
"""

import io
import re
from datetime import datetime

import pandas as pd

from . import constants as C
from . import parsers as P

VERSI_FORMAT = "1.0"

# pemetaan sheet -> kunci sumber di dataset
SHEET_Q = {"APBN Q": "apbn", "APBD Q": "apbd", "SSF Q": "ssf"}
SHEET_A = {"APBD A": "apbd", "SSF A": "ssf"}
SINGLE_SERIES = {"MBG": "mbg_total", "STIK": "stik_total", "Output BI": "output_bi"}
ALL_SHEETS = ["Petunjuk", "Meta", "APBN Q", "APBD Q", "APBD A", "Desa K3",
              "SSF Q", "SSF A", "MBG", "STIK", "Output BI", "Deflator"]

ACTIVE_CODES = {
    "apbn": [s["code"] for s in C.APBN_SUBTABLES if not s["is_rekap"]],
    "apbd": [s["code"] for s in C.APBD_SUBTABLES if not s["is_rekap"]],
    "ssf": [s["code"] for s in C.SSF_SUBTABLES if not s["is_rekap"]],
}

DESA_COMPONENTS = ["Belanja Pegawai", "Belanja Barang",
                   "Belanja Modal", "Penerimaan"]

PETUNJUK_LINES = [
    "TEMPLATE / FILE KERJA - LK PKP Rebasing 2025=100 Kab. Cirebon",
    "",
    "1. Isi HANYA sel angka (kolom triwulan/tahun). JANGAN mengubah nama",
    "   sheet, nama kolom, kolom SubTabel, maupun kolom Komponen.",
    "2. Satuan utama: JUTA RUPIAH. Khusus sheet 'Desa K3': RUPIAH penuh.",
    "   Sheet 'MBG' boleh Rupiah penuh atau Juta - aplikasi mendeteksi sendiri.",
    "3. Penulisan angka bebas: gaya Indonesia (5.976.606,65) atau gaya",
    "   biasa (5976606.65). Di Excel Indonesia tampilan otomatis titik-koma.",
    "4. Sel kosong dianggap NOL. Nilai minus akan dilaporkan sebagai anomali",
    "   (misalnya kasus SIKD) saat file diunggah, sebelum diterapkan.",
    "5. Sheet 'Deflator' berisi INDEKS (basis 2025=100), bukan rupiah.",
    "6. Setelah diisi, unggah kembali file ini ke aplikasi (Langkah 1 atau 2),",
    "   periksa pratinjau & laporan anomali, lalu klik tombol 'Terapkan'.",
    "7. File ini juga berfungsi sebagai SIMPANAN pekerjaan: unduh lewat",
    "   tombol sidebar 'Simpan Pekerjaan', lanjutkan kapan saja lewat 'Buka'.",
]


# ------------------------------------------------------------------ #
# Util kecil
# ------------------------------------------------------------------ #
def _norm_code(x):
    if x is None:
        return None
    s = str(x).strip()
    if s in ("", "nan", "None"):
        return None
    return s.replace(",", ".")


def _labels_in(columns):
    out = []
    for c in columns:
        if isinstance(c, str) and re.fullmatch(r"\d{4}Q[1-4]", c.strip()):
            out.append(c.strip())
    return out


def _years_in(columns):
    out = []
    for c in columns:
        s = str(c).strip()
        if re.fullmatch(r"\d{4}", s):
            out.append(int(s))
    return out


# ------------------------------------------------------------------ #
# MEMBANGUN FILE KERJA (template kosong / simpanan berisi)
# ------------------------------------------------------------------ #
def build_workbook_bytes(dataset, setup, filled=True):
    """
    dataset=None + filled=False  -> template KOSONG
    dataset=isi + filled=True    -> file SIMPANAN pekerjaan
    """
    if dataset is not None:
        quarters = dataset["quarters"]
    else:
        cq = (setup or {}).get("current_quarter") or (2026, 2)
        sy = (setup or {}).get("start_year", C.START_YEAR)
        quarters = C.build_quarters(sy, cq[0], cq[1])
    labels = C.qlabels(quarters)
    years = sorted({y for (y, q) in quarters})
    cq = (setup or {}).get("current_quarter") or quarters[-1]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        # ---------- Petunjuk ---------- #
        pd.DataFrame({"Petunjuk Pengisian": PETUNJUK_LINES}).to_excel(
            writer, sheet_name="Petunjuk", index=False)

        # ---------- Meta ---------- #
        meta = [
            ["versi_format", VERSI_FORMAT],
            ["kode_kab", C.KODE_KAB],
            ["wilayah", C.NAMA_KAB],
            ["tahun_awal", quarters[0][0]],
            ["triwulan_berjalan", C.qlabel(*cq)],
            ["disimpan_pada", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ]
        pd.DataFrame(meta, columns=["Kunci", "Nilai"]).to_excel(
            writer, sheet_name="Meta", index=False)

        # ---------- Sheet triwulanan (Q) ---------- #
        for sheet, src in SHEET_Q.items():
            src_dict = (dataset or {}).get(f"{src}_quarterly", {}) if filled else {}
            rows = []
            for code in ACTIVE_CODES[src]:
                for comp in C.BASE_ROWS:
                    arr = src_dict.get(code, {}).get(comp, [])
                    rows.append([code, comp] +
                                [(arr[i] if i < len(arr) else None)
                                 for i in range(len(labels))])
            pd.DataFrame(rows, columns=["SubTabel", "Komponen"] + labels).to_excel(
                writer, sheet_name=sheet, index=False)

        # ---------- Sheet tahunan (A) ---------- #
        for sheet, src in SHEET_A.items():
            ann = (dataset or {}).get(f"{src}_annual", {}) if filled else {}
            rows = []
            for code in ACTIVE_CODES[src]:
                for comp in C.BASE_ROWS:
                    rows.append([code, comp] +
                                [ann.get(code, {}).get(y, {}).get(comp) for y in years])
            pd.DataFrame(rows,
                         columns=["SubTabel", "Komponen"] + [str(y) for y in years]
                         ).to_excel(writer, sheet_name=sheet, index=False)

        # ---------- Desa K3 (Rupiah penuh) ---------- #
        k3 = (dataset or {}).get("desa_k3", {}) if filled else {}
        rows = [[comp] + [k3.get(y, {}).get(comp) for y in years]
                for comp in DESA_COMPONENTS]
        pd.DataFrame(rows, columns=["Komponen"] + [str(y) for y in years]).to_excel(
            writer, sheet_name="Desa K3", index=False)

        # ---------- Seri tunggal ---------- #
        for sheet, key in SINGLE_SERIES.items():
            ser = (dataset or {}).get(key, {}) if filled else {}
            pd.DataFrame([["Nilai"] + [ser.get(l) for l in labels]],
                         columns=["Komponen"] + labels).to_excel(
                writer, sheet_name=sheet, index=False)

        # ---------- Deflator ---------- #
        defl = (dataset or {}).get("deflators", {}) if filled else {}
        rows = [[name] + [defl.get(name, {}).get(l) for l in labels]
                for name in C.DEFLATOR_NAMES]
        pd.DataFrame(rows, columns=["Deflator"] + labels).to_excel(
            writer, sheet_name="Deflator", index=False)

        # ---------- Format angka & kenyamanan ---------- #
        for ws in writer.book.worksheets:
            if ws.title in ("Petunjuk", "Meta"):
                continue
            for row_cells in ws.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'
            try:
                ws.freeze_panes = "C2" if ws.title.endswith(("Q", "A")) else "B2"
                ws.column_dimensions["A"].width = 12
                ws.column_dimensions["B"].width = 22
            except Exception:
                pass

    return buf.getvalue()


# ------------------------------------------------------------------ #
# MEMBACA FILE KERJA (unggah)
# ------------------------------------------------------------------ #
def _empty_parsed():
    return {
        "meta": {},
        "apbn_quarterly": {}, "apbd_quarterly": {}, "ssf_quarterly": {},
        "apbd_annual": {}, "ssf_annual": {},
        "desa_k3": {},
        "mbg_total": {}, "stik_total": {}, "output_bi": {},
        "deflators": {},
    }


def parse_workbook(file_bytes):
    """Baca File Kerja -> (parsed, report). Tidak mengubah apa pun."""
    xls = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, dtype=object)
    parsed = _empty_parsed()
    report = {"sheets_read": [], "sheets_missing": [], "units": {},
              "rows_filled": {}, "negatives": [], "unknown_codes": [],
              "meta": {}}

    for name in ALL_SHEETS:
        if name not in xls:
            report["sheets_missing"].append(name)
    report["sheets_read"] = [n for n in ALL_SHEETS if n in xls]

    # ---------- Meta ---------- #
    if "Meta" in xls:
        dfm = xls["Meta"]
        for _, row in dfm.iterrows():
            k = str(row.get("Kunci") or "").strip()
            v = row.get("Nilai")
            if k:
                parsed["meta"][k] = str(v).strip() if not isinstance(v, float) else v
        report["meta"] = parsed["meta"]

    # ---------- Sheet Q ---------- #
    for sheet, src in SHEET_Q.items():
        if sheet not in xls:
            continue
        df = xls[sheet]
        labels = _labels_in(df.columns)
        entries = []
        for _, row in df.iterrows():
            code = _norm_code(row.get("SubTabel"))
            comp = str(row.get("Komponen") or "").strip()
            if comp not in C.BASE_ROWS:
                continue
            if code not in ACTIVE_CODES[src]:
                if code:
                    report["unknown_codes"].append(f"{sheet}:{code}")
                continue
            for l in labels:
                v = P.parse_id_number(row.get(l))
                if v is not None:
                    entries.append((code, comp, l, v))
        raws = [e[3] for e in entries]
        unit = P.detect_unit(raws) if raws else "Juta"
        factor = 1e-6 if unit == "Rupiah" else 1.0
        for code, comp, l, v in entries:
            val = v * factor
            (parsed[f"{src}_quarterly"]
             .setdefault(code, {}).setdefault(comp, {})[l] = val)
            if val < 0:
                report["negatives"].append(
                    {"sheet": sheet, "subtabel": code, "komponen": comp,
                     "periode": l, "nilai": val})
        report["units"][sheet] = unit
        report["rows_filled"][sheet] = len(entries)

    # ---------- Sheet A ---------- #
    for sheet, src in SHEET_A.items():
        if sheet not in xls:
            continue
        df = xls[sheet]
        years = _years_in(df.columns)
        entries = []
        for _, row in df.iterrows():
            code = _norm_code(row.get("SubTabel"))
            comp = str(row.get("Komponen") or "").strip()
            if comp not in C.BASE_ROWS:
                continue
            if code not in ACTIVE_CODES[src]:
                if code:
                    report["unknown_codes"].append(f"{sheet}:{code}")
                continue
            for y in years:
                v = P.parse_id_number(row.get(str(y)))
                if v is None:
                    v = P.parse_id_number(row.get(y))
                if v is not None:
                    entries.append((code, comp, y, v))
        raws = [e[3] for e in entries]
        unit = P.detect_unit(raws) if raws else "Juta"
        factor = 1e-6 if unit == "Rupiah" else 1.0
        for code, comp, y, v in entries:
            (parsed[f"{src}_annual"]
             .setdefault(code, {}).setdefault(y, {})[comp] = v * factor)
        report["units"][sheet] = unit
        report["rows_filled"][sheet] = len(entries)

    # ---------- Desa K3 (target: Rupiah penuh) ---------- #
    if "Desa K3" in xls:
        df = xls["Desa K3"]
        years = _years_in(df.columns)
        entries = []
        for _, row in df.iterrows():
            comp = str(row.get("Komponen") or "").strip()
            if comp not in DESA_COMPONENTS:
                continue
            for y in years:
                v = P.parse_id_number(row.get(str(y)))
                if v is None:
                    v = P.parse_id_number(row.get(y))
                if v is not None:
                    entries.append((comp, y, v))
        raws = [e[2] for e in entries]
        unit = P.detect_unit(raws) if raws else "Rupiah"
        factor = 1e6 if unit == "Juta" else 1.0
        for comp, y, v in entries:
            parsed["desa_k3"].setdefault(y, {})[comp] = v * factor
        report["units"]["Desa K3"] = unit
        report["rows_filled"]["Desa K3"] = len(entries)

    # ---------- Seri tunggal ---------- #
    for sheet, key in SINGLE_SERIES.items():
        if sheet not in xls:
            continue
        df = xls[sheet]
        labels = _labels_in(df.columns)
        entries = []
        for _, row in df.iterrows():
            for l in labels:
                v = P.parse_id_number(row.get(l))
                if v is not None:
                    entries.append((l, v))
        raws = [e[1] for e in entries]
        unit = P.detect_unit(raws) if raws else "Juta"
        factor = 1e-6 if unit == "Rupiah" else 1.0
        for l, v in entries:
            parsed[key][l] = v * factor
        report["units"][sheet] = unit
        report["rows_filled"][sheet] = len(entries)

    # ---------- Deflator (indeks, tanpa konversi) ---------- #
    if "Deflator" in xls:
        df = xls["Deflator"]
        labels = _labels_in(df.columns)
        n = 0
        for _, row in df.iterrows():
            name = str(row.get("Deflator") or "").strip()
            if name not in C.DEFLATOR_NAMES:
                continue
            for l in labels:
                v = P.parse_id_number(row.get(l))
                if v is not None:
                    parsed["deflators"].setdefault(name, {})[l] = v
                    n += 1
        report["units"]["Deflator"] = "Indeks"
        report["rows_filled"]["Deflator"] = n

    return parsed, report


# ------------------------------------------------------------------ #
# MENJADIKAN DATASET APLIKASI (setelah tombol Terapkan)
# ------------------------------------------------------------------ #
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
            for code in ACTIVE_CODES[src]
        }
        ann = {}
        for code, cd in parsed[f"{src}_annual"].items():
            ann[code] = {y: {comp: yd.get(comp, 0.0) for comp in C.BASE_ROWS}
                         for y, yd in cd.items() if y in years}
        ds[f"{src}_annual"] = ann

    ds["desa_k3"] = {y: parsed["desa_k3"][y]
                     for y in years if parsed["desa_k3"].get(y)}
    for key in ["mbg_total", "stik_total", "output_bi"]:
        ds[key] = {l: parsed[key].get(l, 0.0) for l in labels}

    defl = {}
    for name in C.DEFLATOR_NAMES:
        src_ser = parsed["deflators"].get(name, {})
        defl[name] = {l: (src_ser.get(l) if src_ser.get(l) is not None else 100.0)
                      for l in labels}
    ds["deflators"] = defl
    return ds