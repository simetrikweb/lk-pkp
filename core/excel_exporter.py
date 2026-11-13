"""
core/excel_exporter.py (v3 - tahan banting)
Workbook .xlsx lengkap: LK3 (Nilai + growth q-o-q/y-o-y/c-o-c),
LK4 (ADHB/Deflator/ADHK/Rekap), Upload Siperi, Upload Verdana.
- Sheet "Info" selalu dibuat pertama -> mustahil error "0 sheet".
- Tiap bagian dibungkus try/except -> satu bagian gagal tidak membatalkan file.
- Sel numerik diberi number_format '#,##0.00' (tampil kaidah Indonesia di Excel Indonesia).
"""

import io

import pandas as pd

from . import constants as C
from . import pkp_calculator as PC


def _nine_df(nine_series, quarters):
    labels = C.qlabels(quarters)
    data = {l: [nine_series.get(r, {}).get(l) for r in C.STD_ROWS] for l in labels}
    return pd.DataFrame(data, index=C.STD_ROWS)


def _write_9row_blocks(writer, sheet, nine_series, quarters):
    """Tulis blok 9 baris baku + blok growth Output & PKP."""
    labels = C.qlabels(quarters)
    df = _nine_df(nine_series, quarters)
    df.to_excel(writer, sheet_name=sheet, startrow=0, index=True)
    row = len(df) + 3
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


def export_workbook(result, quarters):
    buf = io.BytesIO()
    labels = C.qlabels(quarters)
    catatan = []

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:

        # ---- Sheet penjamin: selalu ada, dibuat paling awal ---- #
        pd.DataFrame([["LK PKP Rebasing 2025=100 - Kabupaten Cirebon"],
                      ["File dibuat otomatis oleh aplikasi pkp.py"]]).to_excel(
            writer, sheet_name="Info", index=False, header=False)

        def aman(nama, fn):
            """Jalankan satu bagian; bila gagal, catat dan lanjut ke bagian lain."""
            try:
                fn()
            except Exception as e:
                catatan.append(f"Bagian '{nama}' dilewati: {e}")

        # ---------------- LK 3 ---------------- #
        for code, nine in (result.get("apbn_nine") or {}).items():
            aman(f"LK3 APBN {code}",
                 lambda c=code, n=nine: _write_9row_blocks(
                     writer, f"1.APBN Q {c}"[:31], n, quarters))

        for code, nine in (result.get("apbd_nine_raw") or {}).items():
            aman(f"LK3 APBD {code}",
                 lambda c=code, n=nine: _write_9row_blocks(
                     writer, f"2.APBD Q {c}"[:31], n, quarters))

        for code, nine in (result.get("apbd_nine_denton") or {}).items():
            aman(f"LK3 Denton {code}",
                 lambda c=code, n=nine: _write_9row_blocks(
                     writer, f"12.APBD D {c}"[:31], n, quarters))

        if result.get("desa_nine"):
            aman("LK3 Desa",
                 lambda: _write_9row_blocks(
                     writer, "3.Desa Q 3.1", result["desa_nine"]["3.1"], quarters))

        for code, nine in (result.get("ssf_nine") or {}).items():
            aman(f"LK3 SSF {code}",
                 lambda c=code, n=nine: _write_9row_blocks(
                     writer, f"4.SSF Q {c}"[:31], n, quarters))

        # ---------------- LK 4 ---------------- #
        def tulis_rekap(kunci, judul_kol, sheet):
            src = result.get(kunci) or {}
            rows = [[judul_kol] + labels]
            for key, ser in src.items():
                rows.append([key] + [(ser.get(l) if isinstance(ser, dict) else None)
                                     for l in labels])
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet,
                                        index=False, header=False)

        aman("ADHB", lambda: tulis_rekap("lk4_adhb", "Seri ADHB (Juta Rp)", "ADHB"))
        aman("Deflator", lambda: tulis_rekap("deflators", "Deflator (2025=100)", "Deflator"))
        aman("ADHK", lambda: tulis_rekap("lk4_adhk", "Seri ADHK (Juta Rp)", "ADHK"))
        aman("Rekap", lambda: tulis_rekap("rekap", "Indikator", "Rekap"))

        # ---------------- Upload Siperi ---------------- #
        def tulis_siperi():
            sip_rows = []
            for id_code, ser in (result.get("siperi") or {}).items():
                for l in labels:
                    sip_rows.append([id_code, l,
                                     ser.get(l) if isinstance(ser, dict) else None])
            if not sip_rows:
                sip_rows = [["10", labels[0], 0.0]]
            pd.DataFrame(sip_rows, columns=["id", "periode", "nilai"]).to_excel(
                writer, sheet_name="Upload Siperi", index=False)

        aman("Upload Siperi", tulis_siperi)

        # ---------------- Upload Verdana ---------------- #
        def tulis_verdana():
            rows = [["Lapangan Usaha"] + labels]
            for lu, ser in (result.get("verdanan") or {}).items():
                rows.append([lu] + [ser.get(l) for l in labels])
            pd.DataFrame(rows).to_excel(writer, sheet_name="Upload Verdana",
                                        index=False, header=False)

        aman("Upload Verdana", tulis_verdana)

        # ---------------- Catatan bagian yang gagal (bila ada) ---------------- #
        if catatan:
            pd.DataFrame({"Catatan Perbaikan": catatan}).to_excel(
                writer, sheet_name="Info", startrow=4, index=False)

        # ---------------- Format angka gaya Indonesia di Excel ---------------- #
        for ws in writer.book.worksheets:
            for row_cells in ws.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

    return buf.getvalue()