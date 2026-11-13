"""
pkp.py — Aplikasi Utama
Lembar Kerja PKP Rebasing 2025=100 — Kabupaten Cirebon (3209)
Guided Wizard 6 Langkah untuk pemula.
©2026 Arief Rahman (tahun berjalan otomatis)
"""

from datetime import date

import streamlit as st
import pandas as pd
import plotly.express as px

from core import constants as C
from core import parsers as P
from core import denton as D
from core import pkp_calculator as PC
from core import validators as V
from core import excel_exporter as X
from data import sample_cirebon as SC

# ------------------------------------------------------------------ #
# COPYRIGHT (tahun berjalan otomatis)
# ------------------------------------------------------------------ #
HAKI = f"©{date.today().year} Arief Rahman"

st.set_page_config(page_title="LK PKP Rebasing 2025=100 - Kab. Cirebon",
                   layout="wide")

# ------------------------------------------------------------------ #
# SESSION STATE
# ------------------------------------------------------------------ #
if "step" not in st.session_state:
    st.session_state.step = 1
if "dataset" not in st.session_state:
    st.session_state.dataset = None
if "result" not in st.session_state:
    st.session_state.result = {}
if "denton_log" not in st.session_state:
    st.session_state.denton_log = []

# ------------------------------------------------------------------ #
# TEKS PANDUAN PEMULA
# ------------------------------------------------------------------ #
PETUNJUK_MD = """
**Cara pakai singkat (ikuti urutan):**
1. **Langkah 1** — Pilih wilayah & triwulan berjalan. Pemula disarankan klik tombol
   **"Muat Dataset Simulasi Kab. Cirebon"** agar semua terisi contoh data.
2. **Langkah 2** — Isi atau unggah data mentah per tab (APBN, APBD, Desa, SSF, MBG, Deflator).
   Anda boleh mengetik angka gaya Indonesia (`5.976.606,65`) **atau** gaya biasa (`5976606.65`).
   Satuan Rupiah penuh otomatis dikonversi ke Juta Rupiah.
3. **Langkah 3** — Periksa preview LK 3 (9 baris baku) per sumber dana.
4. **Langkah 4** — Jalankan benchmarking **Denton** (APBN otomatis dilewati karena audited).
   Baca log Kasus 1/2/3 dan grafik pra-vs-pasca.
5. **Langkah 5** — Lihat hasil LK 4: ADHB, ADHK (2025=100), deflator, rekap growth, dan grafik.
6. **Langkah 6** — Jalankan audit otomatis lalu **unduh workbook .xlsx** lengkap
   (termasuk sheet siap unggah `Upload Siperi` & `Upload Verdana`).

**Tips pengisian angka:** gunakan titik sebagai pemisah ribuan dan koma sebagai desimal
sesuai kaidah Indonesia, contoh `5.976.606,65`. Aplikasi juga menerima format biasa.
"""

DATA_EKSTERNAL_MD = """
| No | Data yang dibutuhkan | Sumber / Instansi | Frekuensi | Catatan |
|----|----------------------|-------------------|-----------|---------|
| 1 | Realisasi belanja APBN per satker | KPPN / SPN / MSDKI | Triwulan & Tahun | Audited → Denton di-bypass |
| 2 | Realisasi APBD (per fungsi) | SIKD DJPK & BPKAD | Triwulan & Tahun | Waspadai sel minus (anomali SIKD) |
| 3 | Data Neraca K3 Desa (`kodekab 3209`) | Kemendesa / DJPK | Tahun | Biasanya Rupiah penuh → auto /1.000.000 |
| 4 | Alokasi SSF (Jamkes, JKK, JK, JKP, JP) | BPJS Kes / BPJS Naker | Triwulan & Tahun | Masuk fungsi Administrasi Pemerintahan |
| 5 | Total STIK & MBG | BGN / Dinkes / Dikbud | Triwulan | 2025 sering Rupiah penuh |
| 6 | Deflator: IHPB, Indeks Upah, Implisit PMTB, IHK | BPS / BPS Provinsi | Triwulan | Basis 2025=100 |
"""

KAMUS_MD = """
- **PKP**: Pengeluaran Konsumsi Pemerintah (output − penerimaan).
- **ADHB / ADHK**: Atas Dasar Harga Berlaku / Konstan.
- **Denton**: metode mencocokkan data triwulan ke total tahunan.
- **Deflator**: indeks harga untuk mengubah berlaku → konstan.
- **Kolektif vs Individu**: layanan pemerintah untuk masyarakat umum vs perorangan.
- **q-o-q / y-o-y / c-o-c**: pertumbuhan triwulan-an, tahunan, dan kumulatif.
"""


# ------------------------------------------------------------------ #
# SIDEBAR (panduan + copyright)
# ------------------------------------------------------------------ #
with st.sidebar:
    st.markdown("## 📊 PKP Rebasing 2025=100")
    st.markdown(f"**{HAKI}**")
    st.divider()
    with st.expander("📖 Petunjuk Penggunaan (baca dulu)", expanded=True):
        st.markdown(PETUNJUK_MD)
    with st.expander("🗂️ Data Eksternal yang Diperlukan"):
        st.markdown(DATA_EKSTERNAL_MD)
    with st.expander("📚 Kamus Singkat"):
        st.markdown(KAMUS_MD)
    st.divider()
    st.caption(f"{HAKI} — dibangun dengan Python Streamlit.")


# ------------------------------------------------------------------ #
# HELPER TAMPILAN & EDITOR (format Indonesia)
# ------------------------------------------------------------------ #
def _panduan(title, body):
    with st.container(border=True):
        st.markdown(f"**📖 Panduan Pemula — {title}**")
        st.markdown(body)


def tampil(df_num, **kw):
    """Tampilkan DataFrame numerik dengan format Indonesia."""
    st.dataframe(P.fmt_df_id(df_num), **kw)


def _nine_df(nine_series, quarters):
    labels = C.qlabels(quarters)
    data = {l: [nine_series.get(r, {}).get(l) for r in C.STD_ROWS] for l in labels}
    return pd.DataFrame(data, index=C.STD_ROWS)


def _rekap_nine(nine_map, codes, quarters):
    out = {r: {} for r in C.STD_ROWS}
    for r in C.STD_ROWS:
        sers = [nine_map[c][r] for c in codes if c in nine_map]
        out[r] = PC.sum_series_list(sers, quarters) if sers else {
            C.qlabel(y, q): 0.0 for (y, q) in quarters}
    return out


def edit_base_rows(source_key, code, quarters, existing=None):
    """Editor 5 baris input x triwulan; sel berupa TEKS format Indonesia."""
    labels = C.qlabels(quarters)
    base = {}
    for r in C.BASE_ROWS:
        arr = (existing or {}).get(r, [])
        base[r] = [arr[i] if i < len(arr) else 0.0 for i in range(len(labels))]
    df_num = pd.DataFrame(base, index=labels).T
    df_str = P.fmt_df_id(df_num)
    edited = st.data_editor(df_str, key=f"ed_{source_key}_{code}",
                            use_container_width=True, num_rows="fixed")
    out = {}
    for r in C.BASE_ROWS:
        out[r] = [(P.parse_id_number(edited.loc[r, l]) or 0.0) for l in labels]
    return out


def edit_series(key, ser, quarters, name="Nilai"):
    labels = C.qlabels(quarters)
    df = pd.DataFrame({name: [ser.get(l, 0.0) for l in labels]}, index=labels)
    edited = st.data_editor(P.fmt_df_id(df), key=key,
                            use_container_width=True, num_rows="fixed")
    return {l: (P.parse_id_number(edited.loc[l, name]) or 0.0) for l in labels}


def merge_uploaded(base_dict, dfu, quarters):
    """Gabungkan file unggahan ke editor (cocokkan nama baris & kolom triwulan)."""
    labels = C.qlabels(quarters)
    comp_col = next((c for c in ["Komponen", "komponen", "Baris", "Row"]
                     if c in dfu.columns), None)
    if comp_col is None:
        dfu = dfu.reset_index()
        comp_col = dfu.columns[0]
    idx = C.quarter_index_map(quarters)
    for _, row in dfu.iterrows():
        comp = str(row[comp_col]).strip()
        if comp in base_dict:
            for l in labels:
                if l in dfu.columns:
                    v = P.parse_id_number(row[l])
                    if v is not None:
                        base_dict[comp][idx[l]] = v
    return base_dict


def _need_dataset():
    if st.session_state.dataset is None:
        st.warning("⚠️ Belum ada dataset. Kembali ke **Langkah 1** dan klik "
                   "**Muat Dataset Simulasi Kab. Cirebon** atau unggah data Anda.")
        st.stop()


def _fig(fig):
    fig.update_layout(separators=".,")   # ribuan=titik, desimal=koma di grafik
    return fig


# ------------------------------------------------------------------ #
# LANGKAH 1
# ------------------------------------------------------------------ #
def step1():
    st.header("Langkah 1 — Setup & Alur Kerja")
    _panduan("Mulai di sini",
             "Pilih wilayah, tahun awal, dan triwulan berjalan. Pemula cukup klik "
             "tombol **Muat Dataset Simulasi Kab. Cirebon** untuk mencoba seluruh alur.")
    col1, col2, col3 = st.columns(3)
    wilayah = col1.selectbox("Wilayah", [f"{C.NAMA_KAB} - {C.KODE_KAB}"])
    start_year = col2.number_input("Tahun awal historis", 2015, 2025, C.START_YEAR)
    cq = col3.selectbox("Triwulan berjalan",
                        [(2025, 4), (2026, 1), (2026, 2), (2026, 3)], index=2)
    st.session_state.setup = {"wilayah": wilayah,
                              "start_year": int(start_year),
                              "current_quarter": cq}
    if st.button("🚀 Muat Dataset Simulasi Kab. Cirebon", type="primary"):
        st.session_state.dataset = SC.build_sample_dataset(cq)
        st.session_state.result = {}
        st.success("Dataset simulasi dimuat. Lanjut ke Langkah 2.")
    if st.session_state.dataset is not None:
        st.info("✅ Dataset siap. Silakan lanjut ke Langkah 2.")
    st.button("Lanjut ke Langkah 2 ➡️",
              on_click=lambda: st.session_state.update(step=2),
              disabled=st.session_state.dataset is None)


# ------------------------------------------------------------------ #
# LANGKAH 2
# ------------------------------------------------------------------ #
def step2():
    _need_dataset()
    st.header("Langkah 2 — Input & Cleaning Data Mentah")
    ds = st.session_state.dataset
    quarters = ds["quarters"]
    labels = C.qlabels(quarters)

    tabs = st.tabs(["APBN", "APBD", "APBDes K3 + Rasio", "SSF & STIK",
                    "MBG & Output BI", "Deflator"])

    with tabs[0]:
        _panduan("Data APBN", "Sumber KPPN/SPN. Bersifat audited → Denton di-bypass. "
                              "Isi 5 komponen per sub-tabel (Juta Rp).")
        up = st.file_uploader("Unggah Excel/CSV APBN (opsional)",
                              type=["xlsx", "xls", "csv"], key="up_apbn")
        for st_def in C.APBN_SUBTABLES:
            if st_def["is_rekap"]:
                continue
            with st.expander(f"{st_def['code']} — {st_def['name']} ({st_def['sifat']})"):
                if up is not None and st.button(f"Terapkan unggahan ke {st_def['code']}",
                                                key=f"btn_apbn_{st_def['code']}"):
                    ds["apbn_quarterly"][st_def["code"]] = merge_uploaded(
                        ds["apbn_quarterly"].get(st_def["code"], {}),
                        P.read_table_file(up), quarters)
                ds["apbn_quarterly"][st_def["code"]] = edit_base_rows(
                    "apbn", st_def["code"], quarters,
                    ds["apbn_quarterly"].get(st_def["code"]))

    with tabs[1]:
        _panduan("Data APBD", "Sumber SIKD DJPK (triwulan) & BPKAD (tahunan). "
                              "Aplikasi mendeteksi nilai minus anomali SIKD.")
        issues = []
        for code, rows in ds.get("apbd_quarterly", {}).items():
            issues += P.detect_negative_anomalies(pd.DataFrame(rows), labels,
                                                  context=f"APBD {code}")
        if issues:
            st.warning(f"⚠️ Ditemukan {len(issues)} nilai minus (anomali SIKD).")
            if st.button("Koreksi: set nilai minus APBD → 0"):
                for rows in ds["apbd_quarterly"].values():
                    for r, arr in rows.items():
                        rows[r] = [max(0.0, v) for v in arr]
                st.success("Koreksi diterapkan.")
        for st_def in C.APBD_SUBTABLES:
            if st_def["is_rekap"]:
                continue
            with st.expander(f"{st_def['code']} — {st_def['name']} ({st_def['sifat']})"):
                ds["apbd_quarterly"][st_def["code"]] = edit_base_rows(
                    "apbd", st_def["code"], quarters,
                    ds["apbd_quarterly"].get(st_def["code"]))
        if st.button("Hitung total tahunan APBD dari indikator"):
            ds["apbd_annual"] = SC.gen_apbd_annual(quarters, ds["apbd_quarterly"])
            st.success("Total tahunan APBD dihitung.")

    with tabs[2]:
        _panduan("APBDes (K3)", "Data Neraca K3 `kodekab 3209` biasanya Rupiah penuh → "
                                "otomatis /1.000.000. Penyusutan = 20% × Belanja Modal.")
        if st.button("Konstruksi indikator triwulanan Desa (K3 × SPUN-O)"):
            ds["desa_quarterly"] = PC.build_desa_quarterly_from_k3(
                ds.get("desa_k3", {}), quarters)
            st.success("Indikator Desa dibangun (Juta Rp).")
        if "desa_quarterly" in ds:
            tampil(pd.DataFrame(ds["desa_quarterly"], index=labels).T,
                   use_container_width=True)

    with tabs[3]:
        _panduan("SSF & STIK", "Seluruh output SSF = Kolektif. Bila alokasi kantor hanya "
                               "Kota Cirebon, pakai rasio BPJS Kes: Barang 42%, Pegawai 55%, "
                               "Penyusutan 3%.")
        for st_def in C.SSF_SUBTABLES:
            if st_def["is_rekap"]:
                continue
            with st.expander(f"{st_def['code']} — {st_def['name']}"):
                ds["ssf_quarterly"][st_def["code"]] = edit_base_rows(
                    "ssf", st_def["code"], quarters,
                    ds["ssf_quarterly"].get(st_def["code"]))
        st.subheader("Input Total STIK (Juta Rp)")
        ds["stik_total"] = edit_series("ed_stik", ds.get("stik_total", {}), quarters)
        if st.button("Hitung total tahunan SSF dari indikator"):
            ds["ssf_annual"] = SC.gen_ssf_annual(quarters, ds["ssf_quarterly"])
            st.success("Total tahunan SSF dihitung.")

    with tabs[4]:
        _panduan("MBG & Output BI", "MBG 2025 sering Rupiah penuh (auto-normalisasi). "
                                    "Output BI default 0 untuk Kab. Cirebon.")
        st.subheader("Total MBG (auto-normalisasi satuan)")
        vals = [ds.get("mbg_total", {}).get(l, 0.0) for l in labels]
        edited_mbg = edit_series("ed_mbg", ds.get("mbg_total", {}), quarters)
        norm, unit = P.normalize_to_juta([edited_mbg[l] for l in labels])
        ds["mbg_total"] = {l: norm[i] for i, l in enumerate(labels)}
        st.caption(f"Satuan terdeteksi: **{unit}** → dinormalisasi ke Juta Rp.")
        st.subheader("Output BI (Juta Rp)")
        ds["output_bi"] = edit_series("ed_bi", ds.get("output_bi", {}),
                                      quarters, name="Output BI Berlaku")

    with tabs[5]:
        _panduan("Deflator (2025=100)", "Pastikan rata-rata 4 triwulan 2025 ≈ 100,00.")
        for name in C.DEFLATOR_NAMES:
            ds["deflators"][name] = edit_series(
                f"ed_defl_{name.replace(' ', '_')}", ds["deflators"].get(name, {}),
                quarters, name=name)
            q25 = [ds["deflators"][name].get(f"2025Q{q}", 100.0) for q in range(1, 5)]
            avg = sum(q25) / 4.0
            if abs(avg - 100.0) > 0.5:
                st.warning(f"{name}: rata-rata 2025 = {P.fmt_id(avg)} (seharusnya ~100,00).")
            else:
                st.success(f"{name}: rata-rata 2025 = {P.fmt_id(avg)} ✓")

    st.button("Lanjut ke Langkah 3 ➡️",
              on_click=lambda: st.session_state.update(step=3))


# ------------------------------------------------------------------ #
# LANGKAH 3
# ------------------------------------------------------------------ #
def step3():
    _need_dataset()
    st.header("Langkah 3 — Preview LK 3 (Kuning & Oranye)")
    ds = st.session_state.dataset
    quarters = ds["quarters"]
    result = st.session_state.result

    tab_k, tab_o = st.tabs(["Kuning (APBN)", "Oranye (APBD)"])
    with tab_k:
        apbn_nine = {}
        for code, rows in ds.get("apbn_quarterly", {}).items():
            nine = PC.compute_nine_rows(rows, quarters)
            apbn_nine[code] = PC.nine_rows_as_series(nine, quarters)
        apbn_nine["1.7"] = _rekap_nine(apbn_nine, ["1.1", "1.2", "1.3"], quarters)
        apbn_nine["1.8"] = _rekap_nine(apbn_nine, ["1.4", "1.5", "1.6"], quarters)
        apbn_nine["1.9"] = _rekap_nine(apbn_nine, ["1.7", "1.8"], quarters)
        result["apbn_nine"] = apbn_nine
        for code in ["1.7", "1.8", "1.9"] + list(apbn_nine.keys()):
            if code in apbn_nine:
                with st.expander(f"APBN {code}"):
                    tampil(_nine_df(apbn_nine[code], quarters), use_container_width=True)
    with tab_o:
        apbd_raw = {}
        for code, rows in ds.get("apbd_quarterly", {}).items():
            nine = PC.compute_nine_rows(rows, quarters)
            apbd_raw[code] = PC.nine_rows_as_series(nine, quarters)
        apbd_raw["2.7"] = _rekap_nine(apbd_raw, ["2.1", "2.2", "2.3"], quarters)
        apbd_raw["2.8"] = _rekap_nine(apbd_raw, ["2.4", "2.5", "2.6"], quarters)
        apbd_raw["2.9"] = _rekap_nine(apbd_raw, ["2.7", "2.8"], quarters)
        result["apbd_nine_raw"] = apbd_raw
        for code in apbd_raw:
            with st.expander(f"APBD {code} (pra-Denton)"):
                tampil(_nine_df(apbd_raw[code], quarters), use_container_width=True)

    st.button("Lanjut ke Langkah 4 (Denton) ➡️",
              on_click=lambda: st.session_state.update(step=4))


# ------------------------------------------------------------------ #
# LANGKAH 4
# ------------------------------------------------------------------ #
def step4():
    _need_dataset()
    st.header("Langkah 4 — Benchmarking Denton (Hijau)")
    ds = st.session_state.dataset
    quarters = ds["quarters"]
    result = st.session_state.result
    logs = []
    st.info("APBN di-bypass (audited). APBD, Desa, dan SSF diproses Denton.")

    apbd_denton = {}
    for code, rows in ds.get("apbd_quarterly", {}).items():
        dent = {}
        for r in C.BASE_ROWS:
            annual_r = {y: ds.get("apbd_annual", {}).get(code, {}).get(y, {}).get(r)
                        for y in sorted({yy for (yy, _) in quarters})}
            bench, log = D.denton_benchmark(rows[r], quarters, annual_r,
                                            series_name=f"APBD {code} {r}")
            dent[r] = bench
            logs += log
        nine = PC.compute_nine_rows(dent, quarters)
        apbd_denton[code] = PC.nine_rows_as_series(nine, quarters)
    result["apbd_nine_denton"] = apbd_denton

    if "desa_quarterly" in ds:
        dent = {}
        for r in C.BASE_ROWS:
            bench, log = D.denton_benchmark(ds["desa_quarterly"][r], quarters, {},
                                            series_name=f"Desa 3.1 {r}")
            dent[r] = bench
            logs += log
        nine = PC.compute_nine_rows(dent, quarters)
        result["desa_nine"] = {"3.1": PC.nine_rows_as_series(nine, quarters)}

    ssf_nine = {}
    for code, rows in ds.get("ssf_quarterly", {}).items():
        dent = {}
        for r in C.BASE_ROWS:
            annual_r = {y: ds.get("ssf_annual", {}).get(code, {}).get(y, {}).get(r)
                        for y in sorted({yy for (yy, _) in quarters})}
            bench, log = D.denton_benchmark(rows[r], quarters, annual_r,
                                            series_name=f"SSF {code} {r}")
            dent[r] = bench
            logs += log
        nine = PC.compute_nine_rows(dent, quarters)
        ssf_nine[code] = PC.nine_rows_as_series(nine, quarters)
    ssf_nine["4.1"] = _rekap_nine(ssf_nine, ["4.2", "4.3", "4.4", "4.5", "4.6"], quarters)
    result["ssf_nine"] = ssf_nine
    st.session_state.denton_log = logs

    st.subheader("Log Penanganan Kasus Denton")
    if logs:
        for l in logs:
            st.text(l)
    else:
        st.success("Tidak ada kasus khusus — semua deret terisi penuh.")

    st.subheader("Grafik Pra vs Pasca Denton (APBD 2.1 — Belanja Pegawai)")
    try:
        labels = C.qlabels(quarters)
        pre = ds["apbd_quarterly"]["2.1"]["Belanja Pegawai"]
        post = result["apbd_nine_denton"]["2.1"]["Belanja Pegawai"]
        dfc = pd.DataFrame({"Triwulan": labels + labels,
                            "Nilai": pre + post,
                            "Seri": ["Pra-Denton"] * len(labels) +
                                    ["Pasca-Denton"] * len(labels)})
        st.plotly_chart(_fig(px.line(dfc, x="Triwulan", y="Nilai", color="Seri",
                                     markers=True)), use_container_width=True)
    except Exception:
        st.write("(Grafik tampil bila APBD 2.1 terisi.)")

    st.button("Lanjut ke Langkah 5 (LK 4) ➡️",
              on_click=lambda: st.session_state.update(step=5))


# ------------------------------------------------------------------ #
# LANGKAH 5
# ------------------------------------------------------------------ #
def step5():
    _need_dataset()
    st.header("Langkah 5 — LK 4 PKP & Dashboard Analisis")
    ds = st.session_state.dataset
    quarters = ds["quarters"]
    labels = C.qlabels(quarters)
    result = st.session_state.result
    deflators = ds["deflators"]

    def collect(nine_map, codes):
        return PC.sum_series_list([nine_map[c]["PKP"] for c in codes if c in nine_map],
                                  quarters)

    apbn = result.get("apbn_nine", {})
    apbd = result.get("apbd_nine_denton", {})
    ssf = result.get("ssf_nine", {})

    pusat_kol = collect(apbn, ["1.1", "1.2", "1.3"])
    pusat_ind = collect(apbn, ["1.4", "1.5", "1.6"])
    daerah_kol = collect(apbd, ["2.1", "2.2", "2.3"])
    daerah_ind = collect(apbd, ["2.4", "2.5", "2.6"])
    if "desa_nine" in result:
        daerah_kol = PC.sum_series_list(
            [daerah_kol, result["desa_nine"]["3.1"]["PKP"]], quarters)
    ssf_kol = collect(ssf, ["4.2", "4.3", "4.4", "4.5", "4.6"])
    mbg_pend, mbg_kes = PC.split_mbg(ds.get("mbg_total", {}), quarters)
    stik_jam, stik_jkk = PC.split_stik(ds.get("stik_total", {}), quarters)
    stik_ind = PC.sum_series_list([mbg_pend, mbg_kes, stik_jam, stik_jkk], quarters)
    bi_ser = ds.get("output_bi", {})

    adhb_kol = PC.sum_series_list([pusat_kol, daerah_kol, ssf_kol], quarters)
    adhb_ind = PC.sum_series_list([pusat_ind, daerah_ind, stik_ind], quarters)
    adhb_total = PC.sum_series_list([adhb_kol, adhb_ind], quarters)

    result["lk4_adhb"] = {
        "Pusat Total": PC.sum_series_list([pusat_kol, pusat_ind], quarters),
        "Daerah Final (APBD+Desa)": PC.sum_series_list([daerah_kol, daerah_ind], quarters),
        "SSF, STIK, BI": PC.sum_series_list([ssf_kol, stik_ind, bi_ser], quarters),
    }
    adhk_kol = PC.deflate_series(adhb_kol, deflators.get("IHPB Umum tanpa Ekspor", {}), quarters)
    adhk_ind = PC.deflate_series(adhb_ind, deflators.get("IHK", {}), quarters)
    adhk_total = PC.sum_series_list([adhk_kol, adhk_ind], quarters)
    result["lk4_adhk"] = {"Kolektif ADHK": adhk_kol, "Individu ADHK": adhk_ind,
                          "Total ADHK": adhk_total}
    result["deflators"] = deflators
    implisit = PC.implicit_index(adhb_total, adhk_total, quarters)
    result["rekap"] = {
        "ADHB Kolektif": adhb_kol, "ADHB Individu": adhb_ind, "ADHB Total": adhb_total,
        "ADHK Kolektif": adhk_kol, "ADHK Individu": adhk_ind, "ADHK Total": adhk_total,
        "Indeks Implisit": implisit,
        "Distribusi Kolektif (%)": PC.distribute_pct(adhb_kol, adhb_total, quarters),
        "Distribusi Individu (%)": PC.distribute_pct(adhb_ind, adhb_total, quarters),
        "Growth YoY ADHB (%)": PC.growth_yoy(adhb_total, quarters),
        "Growth QoQ ADHB (%)": PC.growth_qoq(adhb_total, quarters),
        "Growth CtC ADHB (%)": PC.growth_ctc(adhb_total, quarters),
    }
    result["adhb_total"] = adhb_total
    result["adhk_total"] = adhk_total

    cq_lab = C.qlabel(*ds["current_quarter"])
    m1, m2, m3 = st.columns(3)
    m1.metric(f"ADHB Total {cq_lab} (Juta Rp)", P.fmt_id(adhb_total.get(cq_lab)))
    m2.metric(f"ADHK Total {cq_lab} (Juta Rp)", P.fmt_id(adhk_total.get(cq_lab)))
    m3.metric(f"Indeks Implisit {cq_lab}", P.fmt_id(implisit.get(cq_lab)))

    tabs = st.tabs(["ADHB", "ADHK (2025=100)", "Deflator", "Rekap Growth & Implisit"])
    for tab, src in zip(tabs, [result["lk4_adhb"], result["lk4_adhk"],
                               deflators, result["rekap"]):
        with tab:
            df = pd.DataFrame({l: [s.get(l) for s in src.values()] for l in labels},
                              index=list(src.keys()))
            tampil(df, use_container_width=True)

    st.subheader("📊 Kolektif vs Individu (ADHB)")
    dfv = pd.DataFrame({"Triwulan": labels + labels,
                        "Nilai": [adhb_kol.get(l, 0) for l in labels] +
                                 [adhb_ind.get(l, 0) for l in labels],
                        "Jenis": ["Kolektif"] * len(labels) + ["Individu"] * len(labels)})
    st.plotly_chart(_fig(px.line(dfv, x="Triwulan", y="Nilai", color="Jenis",
                                 markers=True)), use_container_width=True)

    st.subheader("📈 Tren Growth ADHB Total")
    dfg = pd.DataFrame({"Triwulan": labels,
                        "YoY (%)": [result["rekap"]["Growth YoY ADHB (%)"].get(l) for l in labels],
                        "QoQ (%)": [result["rekap"]["Growth QoQ ADHB (%)"].get(l) for l in labels],
                        "CtC (%)": [result["rekap"]["Growth CtC ADHB (%)"].get(l) for l in labels]})
    st.plotly_chart(_fig(px.line(dfg, x="Triwulan", y=["YoY (%)", "QoQ (%)", "CtC (%)"],
                                 markers=True)), use_container_width=True)

    st.button("Lanjut ke Langkah 6 (Audit & Ekspor) ➡️",
              on_click=lambda: st.session_state.update(step=6))


# ------------------------------------------------------------------ #
# LANGKAH 6
# ------------------------------------------------------------------ #
def _build_siperi(result, quarters):
    return {
        C.SIPERI_ID["PKP_Total"]: result.get("adhb_total", {}),
        C.SIPERI_ID["Kolektif"]: result.get("rekap", {}).get("ADHB Kolektif", {}),
        C.SIPERI_ID["Individu"]: result.get("rekap", {}).get("ADHB Individu", {}),
    }


def _build_verdana(result, quarters):
    apbn = result.get("apbn_nine", {})
    apbd = result.get("apbd_nine_denton", {})

    def pick(codes, row="NTB"):
        sers = []
        for m in (apbn, apbd):
            for c in codes:
                if c in m:
                    sers.append(m[c].get(row, {}))
        return PC.sum_series_list(sers, quarters)

    return {
        "Sampah/Limbah": pick(["1.2", "2.2"]),
        "Infokom Pemerintah": pick(["1.3", "2.3"]),
        "O Administrasi Pemerintahan": pick(["1.1", "2.1"]),
        "Pendidikan Pemerintah": pick(["1.4", "2.4"]),
        "Kesehatan & Sosial Pemerintah": pick(["1.5", "2.5"]),
        "Jasa Lainnya Pemerintah": pick(["1.6", "2.6"]),
    }


def step6():
    _need_dataset()
    st.header("Langkah 6 — Audit Validitas Pemula & Ekspor Excel")
    ds = st.session_state.dataset
    quarters = ds["quarters"]
    result = st.session_state.result

    issues = V.run_full_audit({
        "quarters": quarters,
        "adhb_total": result.get("adhb_total", {}),
        "adhk_total": result.get("adhk_total", {}),
        "current_quarter": ds.get("current_quarter"),
        "base_series": ds.get("apbd_quarterly", {}),
    })
    st.subheader("✅ Hasil Audit Otomatis")
    if issues:
        for it in issues:
            st.warning(it)
    else:
        st.success("Seluruh cek validitas lolos. Data siap diekspor.")

    result["siperi"] = _build_siperi(result, quarters)
    result["verdanan"] = _build_verdana(result, quarters)

    st.subheader("️ Pratinjau Upload Siperi (format Indonesia)")
    rows = []
    for id_code, ser in result["siperi"].items():
        for l in C.qlabels(quarters):
            rows.append([id_code, l, P.fmt_id(ser.get(l))])
    st.dataframe(pd.DataFrame(rows, columns=["id", "periode", "nilai"]),
                 use_container_width=True)

    st.subheader("📥 Unduh Workbook Lengkap")
    st.caption("File berisi ANGKA murni siap unggah; di Excel berbahasa Indonesia "
               "tampil otomatis sesuai kaidah titik-koma.")
    try:
        xlsx = X.export_workbook(result, quarters)
        st.download_button("⬇️ Download Workbook Lengkap (.xlsx)", data=xlsx,
                           file_name="LK_PKP_Rebasing2025_KabCirebon.xlsx",
                           mime="application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Gagal membuat workbook: {e}")

    if st.button("🔁 Kembali ke Langkah 1"):
        st.session_state.step = 1


# ------------------------------------------------------------------ #
# MAIN
# ------------------------------------------------------------------ #
def main():
    st.title("📊 Lembar Kerja PKP Rebasing 2025=100 — Kabupaten Cirebon")
    st.caption(f"Guided Wizard 6 Langkah • Setara LK 3 + LK 4 Rebasing Regional • **{HAKI}**")

    step = st.session_state.step
    st.progress(min(step, 6) / 6)
    st.markdown(f"**Langkah {step} dari 6**")
    {1: step1, 2: step2, 3: step3, 4: step4, 5: step5, 6: step6}[step]()

    st.divider()
    c1, c2 = st.columns(2)
    if c1.button("⬅️ Kembali", disabled=step <= 1):
        st.session_state.step = max(1, step - 1)
        st.rerun()
    if c2.button("➡️ Lanjut", disabled=step >= 6):
        st.session_state.step = min(6, step + 1)
        st.rerun()
    st.caption(f"{HAKI}")


if __name__ == "__main__":
    main()