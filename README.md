# 📊 Aplikasi Pengelola Data Pinjaman Karyawan

Aplikasi berbasis **Flask (Python)** untuk mengelola data pinjaman karyawan koperasi.  
Mendukung import file **DBF/Excel**, perhitungan cicilan otomatis, filter & pencarian, serta export ke **PDF, CSV, dan Excel**.  
Juga tersedia halaman **dashboard interaktif** dengan grafik status cicilan & distribusi divisi menggunakan Chart.js.

---

## 🚀 Fitur Utama

- 🔹 **Import File** DBF atau Excel langsung dari UI  
- 🔹 **Hitung otomatis** angsuran ke, sisa angsuran, & sisa cicilan  
- 🔹 **Filter & Pencarian** berdasarkan nama, nomor pegawai, bagian, atau status  
- 🔹 **Export Data** ke PDF, CSV, dan Excel  
- 🔹 **Dashboard** dengan ringkasan total cicilan, rata-rata cicilan, jumlah karyawan, dan grafik interaktif  

---

## 🛠️ Teknologi yang Digunakan

- **Backend**: [Flask](https://flask.palletsprojects.com/)  
- **Frontend**: Bootstrap 5 + Chart.js  
- **Database File Support**: DBF (`dbfread`), Excel (`pandas`, `openpyxl`)  
- **Export PDF**: ReportLab  

---

## 📥 Cara Install & Jalankan

1. **Clone repository**
   ```bash
   git clone https://github.com/username/pinjaman-karyawan.git
   cd pinjaman-karyawan

2. **Buat virtual environment (opsional tapi direkomendasikan)**
   ```bash
   python -m venv venv
   source venv/bin/activate   # macOS/Linux
   venv\Scripts\activate      # Windows

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt

4. **Jalankan aplikasi**
   ```bash
   python app.py

5. **Akses di browser**
   ```bash
   http://127.0.0.1:5000

