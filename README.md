# Dashboard K3LH 2026 – PT. Trisatria Persada Borneo

Dashboard Tujuan, Sasaran & Program K3LH 2026 (Streamlit). Tampilan identik dengan versi HTML,
data ditarik otomatis dari file **Program_Tujuan_Sasaran_K3LH_2026_Modern.xlsx** di Google Drive
(tab **Program Kerja 2026**) dan diperbarui setiap 60 detik.

## Isi repo

| File | Fungsi |
|---|---|
| `app.py` | Ambil file dari Drive, hitung semua angka, validasi silang ke rumus Excel, render halaman |
| `dashboard_template.html` | Tampilan (Tailwind + Chart.js), sama dengan dashboard HTML |
| `requirements.txt` | Library Python |
| `.streamlit/config.toml` | Tema gelap |
| `tests/test_parse.py` | Uji perhitungan terhadap isi Excel asli |

## Syarat file di Google Drive

File harus di-share **"Siapa saja yang memiliki link" → Viewer**. Kalau tidak, Streamlit Cloud tidak bisa mengunduhnya.
Isi file boleh diubah kapan saja; struktur kolom (NO, PROGRAM KERJA K3LH, JAN Plan/Act … DES Plan/Act,
TOTAL TARGET, KETERCAPAIAN, RINCIAN BIAYA, PIC) jangan diubah namanya.

Kalau file diganti dengan file lain, ubah `FILE_ID` di baris atas `app.py`.

## Deploy (GitHub + Streamlit Community Cloud)

1. Buat repository baru di GitHub (Public), misal `k3lh-dashboard-2026`.
2. Upload semua file di folder ini (termasuk folder `.streamlit`) ke root repository.
   - Lewat web: *Add file → Upload files*, drag semua file, *Commit changes*.
3. Buka https://share.streamlit.io, login dengan akun GitHub.
4. *Create app* → *Deploy a public app from GitHub*:
   - Repository: `username/k3lh-dashboard-2026`
   - Branch: `main`
   - Main file path: `app.py`
   - App URL: isi nama yang mudah, misal `k3lh-tpb-2026`
5. *Deploy*. Setelah 1–2 menit app bisa dibuka siapa saja di `https://k3lh-tpb-2026.streamlit.app`.

Update tampilan: edit file di GitHub, Streamlit Cloud otomatis redeploy.
Update data: cukup edit file Excel di Drive, dashboard ikut berubah maksimal 60 detik kemudian.

## Jalankan di komputer sendiri

```bash
pip install -r requirements.txt
streamlit run app.py
python tests/test_parse.py   # uji perhitungan
```

## Aturan perhitungan (mengikuti rumus di Excel)

| Angka | Rumus |
|---|---|
| Ketercapaian per program | Σ Actual ÷ Σ Plan (12 bulan) — sama dengan kolom TOTAL TARGET |
| Status | ≥ 100% TERPENUHI, > 0% SEBAGIAN, 0% BELUM — sama dengan kolom KETERCAPAIAN |
| Rata-rata ketercapaian | Rata-rata ketercapaian semua program — sama dengan sel TOTAL TARGET di baris TOTAL |
| Tren bulanan | Σ Actual bulan itu ÷ Σ Plan bulan itu |
| Total anggaran | Σ RINCIAN BIAYA |

Setiap kali data dimuat, hasil hitungan dicocokkan dengan nilai rumus di file Excel
(kolom TOTAL TARGET, KETERCAPAIAN, baris TOTAL, total biaya). Jika ada selisih, muncul kotak peringatan
kuning di atas KPI. Jika cocok, arahkan kursor ke chip "Sinkron Google Drive" untuk melihat status tervalidasi.
