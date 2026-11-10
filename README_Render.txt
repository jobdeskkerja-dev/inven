# InventClass — paket awal untuk Render

## Konfigurasi Render
- Service type: Web Service
- Language/runtime: Python
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
- Environment Variable: `SECRET_KEY` = isi dengan string acak yang panjang dan rahasia.

## Catatan penting tentang penyimpanan
Aplikasi ini menggunakan SQLite (`inventaris.db`) dan folder `foto/` di direktori aplikasi.
Pada filesystem sementara layanan hosting, database dan foto dapat hilang saat instance diganti,
restart, atau deploy ulang. Sebelum dipakai menyimpan data sungguhan, siapkan penyimpanan persisten
(Render Persistent Disk) atau migrasikan database ke layanan database persisten dan sesuaikan kode.

## Akun bawaan yang tertulis di kode
- admin / admin123
- guru / guru123
- siswa / siswa123

Segera ubah password bawaan setelah login. Jangan membagikan file database, SECRET_KEY,
atau password kepada orang lain.

## Struktur
- `app.py`: kode aplikasi Flask hasil ekstraksi dari notebook.
- `requirements.txt`: dependensi Python.
