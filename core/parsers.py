"""
core/parsers.py
- Normalisasi satuan (Rupiah penuh -> Juta Rp)
- Membaca file Excel/CSV (mendukung angka gaya Indonesia "5.976.606,65")
- Deteksi anomali nilai minus (SIKD)
- FORMAT TAMPILAN KAIDAH INDONESIA: titik=ribuan, koma=desimal
"""

import io
import re

import pandas as pd

# Ambang heuristik: sel >= 1 miliar dianggap masih Rupiah penuh
RUPIAH_FULL_THRESHOLD = 1e9


# ------------------------------------------------------------------ #
# FORMAT INDONESIA (tampilan)
# ------------------------------------------------------------------ #
def fmt_id(value, decimals=2):
    """Format angka gaya Indonesia: 5976606.65 -> '5.976.606,65'."""
    if value is None:
        return "-"
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    if pd.isna(v):
        return "-"
    neg = v < 0
    v = abs(v)
    s = f"{v:,.{decimals}f}"                 # dulu gaya Inggris: 5,976,606.65
    s = s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return ("-" if neg else "") + s          # jadi: 5.976.606,65


def fmt_df_id(df, decimals=2):
    """Ubah DataFrame numerik menjadi DataFrame STRING berformat Indonesia."""
    out = pd.DataFrame(index=df.index, columns=df.columns, dtype=object)
    for c in df.columns:
        for i in df.index:
            out.loc[i, c] = fmt_id(df.loc[i, c], decimals)
    return out


# ------------------------------------------------------------------ #
# BACA ANGKA GAYA INDONESIA (input / upload)
# ------------------------------------------------------------------ #
def parse_id_number(x):
    """
    Terima angka dalam berbagai gaya dan kembalkan float.
    Mendukung: 5.976.606,65 | 5976606.65 | 1.000.000 | (123,45) | -12,5 | Rp 1.000
    """
    if x is None:
        return None
    if isinstance(x, (int, float)):
        f = float(x)
        return None if pd.isna(f) else f
    s = str(x).strip()
    if s in ("", "-", "nan", "None", "--"):
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()
    if s.startswith("-"):
        neg, s = True, s[1:]
    s = re.sub(r"[^0-9.,]", "", s)           # buang Rp, spasi, dll
    if s in ("", ".", ","):
        return None
    has_koma = "," in s
    has_titik = "." in s
    if has_koma and has_titik:
        # gaya Indonesia: titik=ribuan, koma=desimal
        s = s.replace(".", "").replace(",", ".")
    elif has_koma:
        # koma saja -> desimal Indonesia
        s = s.replace(",", ".")
    elif has_titik:
        parts = s.split(".")
        if len(parts) > 2:
            s = s.replace(".", "")           # titik ribuan banyak
        elif len(parts) == 2 and len(parts[1]) == 3:
            s = s.replace(".", "")           # satu titik + 3 digit -> ribuan
        # selain itu dianggap desimal biasa
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def coerce_id_numbers(df):
    """Konversi kolom teks yang berisi angka gaya Indonesia menjadi numerik."""
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            conv = df[c].map(parse_id_number)
            if conv.notna().mean() > 0.7:
                df[c] = conv.astype(float)
    return df


# ------------------------------------------------------------------ #
# NORMALISASI SATUAN
# ------------------------------------------------------------------ #
def detect_unit(series):
    vals = pd.to_numeric(pd.Series(series).values, errors="coerce")
    vals = vals[~pd.isna(vals)]
    if len(vals) == 0:
        return "Juta"
    return "Rupiah" if float(pd.Series(vals).abs().max()) >= RUPIAH_FULL_THRESHOLD else "Juta"


def normalize_to_juta(values):
    s = pd.to_numeric(pd.Series(list(values)).reset_index(drop=True), errors="coerce")
    unit = detect_unit(s)
    if unit == "Rupiah":
        s = s / 1e6
    return s.tolist(), unit


def normalize_dataframe_to_juta(df, value_cols):
    df = df.copy()
    report = {}
    for col in value_cols:
        if col not in df.columns:
            continue
        vals, unit = normalize_to_juta(df[col].tolist())
        df[col] = vals
        report[col] = unit
    return df, report


# ------------------------------------------------------------------ #
# DETEKSI ANOMALI MINUS (SIKD)
# ------------------------------------------------------------------ #
def detect_negative_anomalies(df, value_cols, context=""):
    issues = []
    for col in value_cols:
        if col not in df.columns:
            continue
        s = pd.to_numeric(df[col], errors="coerce")
        for idx in s[s < 0].index.tolist():
            issues.append({
                "kolom": col,
                "baris": idx,
                "nilai": float(s[idx]),
                "konteks": context,
                "saran": "Rekonsiliasi dengan BPKAD / set ke 0 sebelum Denton",
            })
    return issues


# ------------------------------------------------------------------ #
# BACA FILE UNGGAHAN
# ------------------------------------------------------------------ #
def read_table_file(uploaded_file):
    name = getattr(uploaded_file, "name", "") or ""
    data = uploaded_file.read()
    buf = io.BytesIO(data)
    low = name.lower()
    try:
        if low.endswith(".csv") or low.endswith(".txt"):
            df = pd.read_csv(buf, dtype=str)
        else:
            df = pd.read_excel(buf)
    except Exception:
        buf.seek(0)
        df = pd.read_csv(buf, dtype=str)
    return coerce_id_numbers(df)


# ------------------------------------------------------------------ #
# UTIL DataFrame <-> records
# ------------------------------------------------------------------ #
def df_from_records(records, index_col="Komponen", quarter_labels=None):
    df = pd.DataFrame(records)
    if index_col in df.columns:
        df = df.set_index(index_col)
    if quarter_labels is not None:
        for ql in quarter_labels:
            if ql not in df.columns:
                df[ql] = 0.0
        df = df[[c for c in quarter_labels if c in df.columns]]
    return df.astype(float)


def records_from_df(df, index_name="Komponen"):
    out = []
    for idx, row in df.iterrows():
        rec = {index_name: idx}
        for col, val in row.items():
            try:
                rec[col] = float(val)
            except (TypeError, ValueError):
                rec[col] = 0.0
        out.append(rec)
    return out