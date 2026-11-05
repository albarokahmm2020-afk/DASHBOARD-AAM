# CHP AAM Dashboard

Streamlit entry point: `app.py`. Assets are stored beside the app for one-step GitHub web upload. Do not add database exports, credentials, `.streamlit/secrets.toml`, or local settings to this repository.

## Interface and design system

The Streamlit interface uses the light enterprise theme configured in `.streamlit/config.toml`. The project-level UI UX Pro Max skill and its searchable design guidance are installed in `.agents/skills/ui-ux-pro-max/`; the dashboard design tokens and responsive/motion decisions are recorded in `design-system/chp-aam-dashboard/MASTER.md`.

Run the dashboard locally with `streamlit run app.py` after installing `requirements.txt`. Google Drive data restoration and sync remain optional and require credentials in Streamlit secrets.

# Update database dashboard

Untuk pembaruan harian, jangan edit `database_push_branches.xlsx` karena ukurannya terlalu besar untuk Google Sheets. Dashboard secara default membaca file dari `database_monthly_split/<BRANCH>/` agar hanya mengambil data cabang aktif dan memakai cache lokal untuk file yang tidak berubah. Jika folder bulanan untuk cabang belum memiliki workbook XLSX, dashboard memakai workbook gabungan sebagai fallback.

Mode baca dapat dipilih melalui Streamlit Secrets. `database_mode` default ke `auto`, yang mengutamakan file bulanan meskipun `database_file_id` lama masih tersimpan. Gunakan `database_mode = "single"` hanya jika ingin memaksa pembacaan workbook gabungan.

1. Buka spreadsheet bulan berjalan, misalnya `database_push_BDG_2026_10.xlsx` untuk cabang BDG.
2. Tambahkan transaksi di baris paling bawah atau koreksi baris yang ada. Pertahankan header kolom dan nama file yang sama; tunggu Drive selesai menyimpan.
3. Di dashboard, klik **Periksa pembaruan dari Drive**.

File `.xlsx` akan dipakai jika tersedia untuk bulan yang sama dengan `.csv`, jadi baris tidak dihitung dua kali. File CSV lama masih dibaca untuk riwayat. Pada bulan baru, buat/unggah spreadsheet dengan pola `database_push_<BRANCH>_<YYYY>_<MM>.xlsx` ke folder cabang.

## Penyimpanan laporan dan data operasional

Aplikasi menggunakan folder Google Drive [DASHBOARD AAM](https://drive.google.com/drive/folders/1Ef0DsWdqJ-fy7ahuQ0NR3H1dO2QkHk4b) sebagai root penyimpanan. File operasional masuk ke `Dashboard_Data/<CABANG>/`, sedangkan ekspor masuk ke `Dashboard_Exports/<CABANG>/`; keduanya terpisah dari `database_monthly_split/`. Folder tujuan dapat berupa folder Drive biasa selama akun layanan bisa mengakses dan menambahkan subfolder. Jika perlu mengganti tujuan, atur ID folder di secrets Streamlit:

```toml
[google_drive]
storage_folder_id = "ID_FOLDER_SHARED_DRIVE"
```

`storage_folder_id` adalah ID folder Drive, bukan URL. Nilai ini terpisah dari `database_folder_id`, tetapi keduanya boleh menunjuk folder root yang sama. Aplikasi otomatis mencoba folder `Dashboard_Data/<CABANG>/target_breakdown/` untuk menyimpan file mentah. Jika Google Drive membalas `storageQuotaExceeded`, berarti akun layanan tidak punya kuota menulis ke Drive Saya; gunakan Shared Drive yang bisa ditulis atau autentikasi OAuth akun pemilik folder. Jangan masukkan kredensial ke repository.

## Breakdown Target bulanan

Buka **Navigasi utama → Breakdown Target**. Tahap 1: pilih periode dan Business Group, unggah file mapping assignment, masukkan target berjalan per Principal ID, lalu generate pembagian ke Sales Person Title dan Channel Outlet Group. Mapping hanya memakai assignment aktif untuk cabang yang dipilih, dengan `BISNIS` + `POSITION_CODE` + `SALES_CHANNEL`; channel yang boleh diberikan ke tiap posisi dibatasi oleh mapping, sedangkan bobot pembagian di antara channel tersebut mengikuti kontribusi historis tiga bulan. Edit `ADJUST_TARGET` untuk mengoreksi alokasi. Tahap 2: unggah file mentah; dashboard mencocokkan Seller, Principal ID, dan Channel, lalu mengisi `TARGET_CABANG` dan menyediakan ekspor Excel.

`TARGET_CABANG` dibagi ke baris file mentah menurut bobot `TARGET_REKOMENDASI_HO` di dalam kombinasi Seller, Principal ID, dan Channel. Hasil ekspor mempertahankan sheet input serta menambah sheet `Breakdown`. File diproses di sesi dashboard; unduh hasil sebelum meninggalkan dashboard. File tidak disimpan otomatis ke Google Drive.
