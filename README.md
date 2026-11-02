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

Akun layanan Google tidak memiliki kuota untuk menyimpan file di My Drive. Folder Shared Drive [yang Anda kirim](https://drive.google.com/drive/folders/1Ef0DsWdqJ-fy7ahuQ0NR3H1dO2QkHk4b) menjadi sumber database sekaligus root penyimpanan default (`1Ef0DsWdqJ-fy7ahuQ0NR3H1dO2QkHk4b`). Aplikasi menulis data operasional ke subfolder `Dashboard_Data/` dan ekspor ke `Dashboard_Exports/`, terpisah dari `database_monthly_split/`. Tambahkan akun layanan dari `gcp_service_account.client_email` sebagai anggota Shared Drive dengan peran **Kontributor** atau lebih tinggi. Jika perlu mengganti tujuan, atur ID folder di secrets Streamlit:

```toml
[google_drive]
storage_folder_id = "ID_FOLDER_SHARED_DRIVE"
```

`storage_folder_id` adalah ID folder tujuan di dalam Shared Drive, bukan URL atau ID Shared Drive. Nilai ini merupakan pengaturan terpisah dari `database_folder_id`, walau keduanya boleh memakai folder root yang sama karena aplikasi menyimpan file operasional di subfolder khusus. Jika `storage_folder_id` di Streamlit Secrets menunjuk ke ID default lama `1E88RgFi_zTHzsJZzD0RVztdcS5Wecc6e`, atau ID khusus yang tidak valid/tidak bisa diakses, aplikasi otomatis mencoba folder default di atas. Pastikan akun layanan punya izin menambah file di folder default. Jangan masukkan kredensial akun layanan ke repository.

## Breakdown Target bulanan

Buka **Navigasi utama → Breakdown Target**, lalu unggah workbook target bulanan. Kolom input mengikuti template seperti `MENTAHAN UPLOAD TARGET.xlsx`, terutama `YEAR_ID`, `MONTH_ID`, `SUBLINI_PRODUCT`, `ITEM_SUBLINI`, `SALESPERSON_POSITION`, dan `TARGET_REKOMENDASI_HO`. Sistem membagi target HO per sublini ke sales person berdasarkan rata-rata tiga bulan kalender lengkap sebelumnya. Dasar pembagian bisa berupa Sales Value Gross atau jumlah relasi aktif. Untuk baris tanpa sales person, nama sales person yang memiliki data historis untuk produk itu ditambahkan otomatis.

`TARGET_CABANG` hasil ekspor direkonsiliasi per `SUBLINI_PRODUCT` agar totalnya sama dengan satu target HO (nilai HO yang berulang pada baris mentah tidak dijumlahkan berulang). File mentah disimpan di Shared Drive pada `Dashboard_Data/<CABANG>/target_breakdown/`; hasil terbitan disimpan pada `Dashboard_Exports/<CABANG>/`. Gunakan **Muat file mentah yang tersimpan di Google Drive** untuk membuka ulang file berdasarkan cabang dan periode.
