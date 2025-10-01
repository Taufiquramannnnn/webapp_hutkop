# 📊 Aplikasi Pengelola Data Pinjaman Karyawan (HutKop)

Aplikasi web berbasis **Flask (Python)** untuk mengelola, menggabungkan, dan memvisualisasikan data pinjaman karyawan dari berbagai file sumber (`.dbf` atau `.xlsx`). Dirancang khusus untuk melakukan agregasi data secara cerdas, memberikan gambaran total pinjaman yang akurat untuk setiap karyawan.

Dilengkapi dengan fitur filter, pencarian, ekspor laporan profesional (**PDF, CSV, Excel**), dan halaman **dashboard interaktif** untuk analisis data yang mendalam.

---

## 🚀 Fitur Utama

-   **📤 Import Multi-File**: Unggah beberapa file `.dbf` dan `.xlsx` secara bersamaan.
-   **🧠 Agregasi Cerdas**: Secara otomatis mengakumulasi total pinjaman, sisa angsuran, dan status untuk karyawan yang sama, meskipun datanya tersebar di banyak file.
-   **🔍 Filter & Pencarian Lanjutan**: Menyaring data dengan mudah berdasarkan Nama, No. Pegawai, Divisi, atau Status Pinjaman (Lunas, Berjalan, Belum Bayar).
-   **📄 Ekspor Laporan Profesional**: Unduh data gabungan dalam format **PDF** (dengan layout rapi), **CSV**, atau **Excel** dengan satu klik.
-   **📊 Dashboard Interaktif**: Dapatkan wawasan cepat melalui visualisasi data menggunakan Chart.js, menampilkan:
    -   Ringkasan KPI (Total Pinjaman, Sisa Pinjaman, Jumlah Peminjam).
    -   Distribusi Status Pinjaman.
    -   Top 10 Peminjam Terbesar.
    -   Top 10 Divisi berdasarkan Jumlah & Total Pinjaman.
-   **🗑️ Reset Data**: Hapus semua data yang telah diunggah dengan aman untuk memulai analisis dari awal atau jika ada file terupdate dari file sebelumnya yang ingin diupload (data lama wajib dihapus / direset).

---

## 🛠️ Tumpukan Teknologi (Tech Stack)

-   **Backend**: [Flask](https://flask.palletsprojects.com/)
-   **Frontend**: Bootstrap 5 & [Chart.js](https://www.chartjs.org/)
-   **Pemrosesan Data**: `pandas` (untuk Excel) & `dbfread` (untuk DBF)
-   **Pembuatan PDF**: `reportlab`
-   **Deployment (Opsional)**: `pyinstaller` untuk packaging menjadi `.exe`

---

## 📥 Instalasi & Cara Menjalankan

Ikuti langkah-langkah berikut untuk menjalankan aplikasi di komputer lokal Anda.

**1. Clone Repository**
```bash
git clone [https://github.com/Taufiquramannnnn/webapp_hutkop.git](https://github.com/Taufiquramannnnn/webapp_hutkop.git)
cd webapp_hutkop
```

**2. Buat & Aktifkan Virtual Environment (Sangat Direkomendasikan)**
```bash
# Buat virtual environment
python -m venv venv

# Aktifkan di Windows
venv\Scripts\activate

# Aktifkan di macOS/Linux
source venv/bin/activate
```

**3. Install Semua Kebutuhan**
Pastikan Anda berada di dalam virtual environment yang aktif, lalu jalankan:
```bash
pip install -r requirements.txt
```

**4. Jalankan Aplikasi**
```bash
python app.py / klik start.bat.
```
Aplikasi akan berjalan dan otomatis membuka tab baru di browser Anda pada alamat `http://127.0.0.1:5000`.

---