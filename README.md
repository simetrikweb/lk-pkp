# Lembar Kerja PKP

untuk menyusun komponen pkp

## Menjalankan lokal
pip install -r requirements.txt
streamlit run pkp.py

## Deploy Streamlit Cloud
1. Upload semua file ini ke GitHub.
2. Di Streamlit Cloud pilih repository → Main file path: `pkp.py`.
3. Klik Deploy.

## Struktur
- pkp.py                 : aplikasi utama (wizard 6 langkah)
- core/constants.py      : konfigurasi & rasio baku
- core/parsers.py        : format angka Indonesia, normalisasi satuan, baca file
- core/denton.py         : benchmarking Denton-Cholette
- core/pkp_calculator.py : rumus 9 baris, growth, ADHB/ADHK
- core/validators.py     : audit otomatis
- core/excel_exporter.py : ekspor .xlsx
- data/sample_cirebon.py : dataset simulasi
- core/template_io.py    : File Kerja Excel (template kosong, simpan & buka pekerjaan)