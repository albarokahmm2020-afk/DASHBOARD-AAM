# CHP AAM Dashboard

Streamlit entry point: `app.py`. Assets are stored beside the app for one-step GitHub web upload. Do not add database exports, credentials, `.streamlit/secrets.toml`, or local settings to this repository.
# Update database dashboard

Untuk pembaruan harian, jangan edit `database_push_branches.xlsx` karena ukurannya terlalu besar untuk Google Sheets. Dashboard membaca data dari folder `database_monthly_split/<BRANCH>/` dan menggabungkan file bulanan.

1. Buka spreadsheet bulan berjalan, misalnya `database_push_BDG_2026_10.xlsx` untuk cabang BDG.
2. Tambahkan transaksi di baris paling bawah atau koreksi baris yang ada. Pertahankan header kolom dan nama file yang sama; tunggu Drive selesai menyimpan.
3. Di dashboard, klik **Periksa pembaruan dari Drive**.

File `.xlsx` akan dipakai jika tersedia untuk bulan yang sama dengan `.csv`, jadi baris tidak dihitung dua kali. File CSV lama masih dibaca untuk riwayat. Pada bulan baru, buat/unggah spreadsheet dengan pola `database_push_<BRANCH>_<YYYY>_<MM>.xlsx` ke folder cabang.
