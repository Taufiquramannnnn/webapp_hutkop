# 📊 Aplikasi Pengelola Data Pinjaman Karyawan

Aplikasi berbasis **Flask (Python)** untuk mengelola dan memvisualisasikan data pinjaman karyawan dari file DBF atau Excel. Aplikasi ini dirancang untuk melakukan agregasi (penjumlahan) data dari berbagai file untuk karyawan yang memiliki lebih dari satu pinjaman, memberikan gambaran total pinjaman yang akurat.

Dilengkapi dengan fitur filter, pencarian, dan ekspor ke format **PDF, CSV, dan Excel**. Terdapat juga halaman **dashboard interaktif** yang powerful menggunakan Chart.js untuk menampilkan ringkasan data secara visual.

---

## 🚀 Fitur Utama

-   🔹 **Import Multi-File**: Unggah beberapa file `.dbf` atau `.xlsx` sekaligus.
-   🔹 **Agregasi Data Cerdas**: Secara otomatis menjumlahkan total pinjaman, tenor, cicilan, dan sisa angsuran untuk karyawan yang sama yang datanya tersebar di beberapa file.
-   🔹 **Reset Data**: Fitur untuk menghapus semua data yang telah diunggah, memungkinkan perhitungan ulang dari awal dengan mudah.
-   🔹 **Filter & Pencarian Lanjutan**: Cari data berdasarkan nama, nomor pegawai, divisi, atau status pinjaman (Lunas, Berjalan, Belum Bayar).
-   🔹 **Ekspor Profesional**: Ekspor data gabungan ke format **PDF** (dengan layout rapi), **CSV**, dan **Excel**.
-   🔹 **Dashboard Interaktif**: Visualisasikan data dengan ringkasan KPI (total pinjaman, sisa pinjaman, jumlah peminjam) dan grafik interaktif yang menampilkan:
    -   Distribusi Status Pinjaman
    -   Top 10 Peminjam Terbesar
    -   Top 10 Divisi berdasarkan Jumlah Peminjam
    -   Top 10 Divisi berdasarkan Total Pinjaman

---

## 🛠️ Teknologi yang Digunakan

-   **Backend**: [Flask](https://flask.palletsprojects.com/)
-   **Frontend**: Bootstrap 5 & [Chart.js](https://www.chartjs.org/)
-   **Data Processing**:
    -   DBF: `dbfread`
    -   Excel: `pandas` & `openpyxl`
-   **PDF Generation**: `reportlab`
-   **Packaging (Optional)**: `pyinstaller`

---

## 📥 Cara Install & Jalankan

1.  **Clone repository**
    ```bash
    git clone [https://github.com/Taufiquramannnnn/webapp_hutkop.git](https://github.com/Taufiquramannnnn/webapp_hutkop.git)
    cd webapp_hutkop
    ```

2.  **Buat dan aktifkan virtual environment (sangat direkomendasikan)**
    ```bash
    # Buat venv
    python -m venv venv

    # Aktifkan di Windows
    venv\Scripts\activate

    # Aktifkan di macOS/Linux
    source venv/bin/activate
    ```

3.  **Install semua library yang dibutuhkan**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Jalankan aplikasi**
    ```bash
    python app.py
    ```

5.  **Buka di browser**
    Aplikasi akan otomatis terbuka di browser default Anda pada alamat:
    ```
    [http://127.0.0.1:5000](http://127.0.0.1:5000)
    ```