"""
core/excel_exporter.py
Workbook .xlsx lengkap: LK3 (4 blok: Nilai, q-o-q, y-o-y, c-o-c),
LK4 (ADHB/Deflator/ADHK/Rekap), Upload Siperi, Upload Verdana.
Sel numerik diberi number_format '#,##0.00' sehingga di Excel berbahasa
Indonesia tampil otomatis sesuai kaidah (titik ribuan, koma desimal),
tetapi tetap berupa ANGKA murni agar siap unggah ke Siperi/Verdana.
"""

import io

import pandas as pd

from . import constants as C
from . import pkp_calculator as PC


def _nine_df(nine_series, quarters):
    labels = C.qlabels(quarters)
    data = {l: [nine_series.get(r, {}).get(l) for l in labels] for r in C.STD_ROWS}
    return pd.DataFrame(data, index=C.STD_ROWS)


def _write_9row_blocks(writer, sheet, nine_series, quarters):
    """Tulis 4 blok deret waktu (Nilai, qoq, yoy, ctc) untuk 9 baris baku."""
    labels = C.qlabels(quarters)
    row = 0
    blocks = [
        ("NILAI (Juta Rp)", {r: nine_series.get(r, {}) for r in C.STD_ROWS}),
        ("GROWTH q-to-q (%)", None),
        ("GROWTH y-on-y (%)", None),
        ("GROWTH c-to-c (%)", None),
    ]
    # blok nilai
    df = _nine_df(nine_series, quarters)
    df.to_excel(writer, sheet_name=sheet, startrow=row, index=True)
    row += len(df) + 3

    # blok growth per baris baku utama (PKP & Output sebagai contoh lengkap)
    for r in ["Output", "PKP"]:
        ser = nine_series.get(r, {})
        for judul, fn in [("q-to-q", PC.growth_qoq),
                          ("y-on-y", PC.growth_yoy),
                          ("c-to-c", PC.growth_ctc)]:
            g = fn(ser, quarters)
            line = pd.DataFrame([[r] + [g.get(l) for l in labels]],
                                columns=[f"Growth {judul} (%)"] + labels)
            line.to_excel(writer, sheet_name=sheet, startrow=row, index=False)
            row += 2
    return row


def export_workbook(result, quarters):
    buf = io.BytesIO()
    labels = C.qlabels(quarters)

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        # ---------- LK 3 KUNING : APBN ---------- #
        for code, nine in result.get("apbn_nine", {}).items():
            _write_9row_blocks(writer, f"1.APBN Q {code}"[:31], nine, quarters)

        # ---------- LK 3 ORANYE : APBD pra-Denton ---------- #
        for code, nine in result.get("apbd_nine_raw", {}).items():
            _write_9row_blocks(writer, f"2.APBD Q {code}"[:31], nine, quarters)

        # ---------- LK 3 HIJAU : pasca-Denton ---------- #
        for code, nine in result.get("apbd_nine_denton", {}).items():
            _write_9row_blocks(writer, f"12.APBD D {code}"[:31], nine, quarters)

        if "desa_nine" in result:
            _write_9row_blocks(writer, "3.Desa Q 3.1", result["desa_nine"]["3.1"], quarters)

        for code, nine in result.get("ssf_nine", {}).items():
            _write_9row_blocks(writer, f"4.SSF Q {code}"[:31], nine, quarters)

        # ---------- LK 4 : ADHB ---------- #
        rows = [["Seri ADHB"] + labels]
        for key, ser in result.get("lk4_adhb", {}).items():
            rows.append([key] + [ser.get(l) for l in labels])
        pd.DataFrame(rows).to_excel(writer, sheet_name="ADHB", index=False, header=False)

        # ---------- LK 4 : Deflator ---------- #
        rows = [["Deflator (2025=100)"] + labels]
        for name in C.DEFLATOR_NAMES:
            s = result.get("deflators", {}).get(name, {})
            rows.append([name] + [s.get(l) for l in labels])
        pd.DataFrame(rows).to_excel(writer, sheet_name="Deflator", index=False, header=False)

        # ---------- LK 4 : ADHK ---------- #
        rows = [["Seri ADHK"] + labels]
        for key, ser in result.get("lk4_adhk", {}).items():
            rows.append([key] + [ser.get(l) for l in labels])
        pd.DataFrame(rows).to_excel(writer, sheet_name="ADHK", index=False, header=False)

        # ---------- LK 4 : Rekap ---------- #
        rows = [["Indikator"] + labels]
        for key, ser in result.get("rekap", {}).items():
            rows.append([key] + [ser.get(l) for l in labels])
        pd.DataFrame(rows).to_excel(writer, sheet_name="Rekap", index=False, header=False)

        # ---------- Upload Siperi ---------- #
        sip_rows = []
        for id_code, ser in result.get("siperi", {}).items():
            for l in labels:
                sip_rows.append([id_code, l, ser.get(l)])
        pd.DataFrame(sip_rows, columns=["id", "periode", "nilai"]).to_excel(
            writer, sheet_name="Upload Siperi", index=False)

        # ---------- Upload Verdana ---------- #
        rows = [["Lapangan Usaha"] + labels]
        for lu, ser in result.get("verdanan", {}).items():
            rows.append([lu] + [ser.get(l) for l in labels])
        pd.DataFrame(rows).to_excel(writer, sheet_name="Upload Verdana",
                                    index=False, header=False)

        # ---------- FORMAT ANGKA SEL NUMERIK ---------- #
        for ws in writer.book.worksheets:
            for row_cells in ws.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

    return buf.getvalue()