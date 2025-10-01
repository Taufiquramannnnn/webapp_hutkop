"""
app.py
------
Aplikasi utama Flask.
- Bertanggung jawab untuk menjalankan server web.
- Memuat data dari file DBF atau Excel yang diunggah ke folder /uploads.
- Melakukan agregasi (menjumlahkan) data pinjaman untuk Nomor Pegawai (NOPEG) yang sama,
  bahkan jika data berasal dari file yang berbeda.
- Menampilkan data dalam format master-detail (ringkasan & rincian pinjaman per orang).
- Menyediakan fitur filter, pencarian, dan ekspor data ke format CSV, Excel, dan PDF.
- Menyediakan halaman dashboard untuk visualisasi data.
"""

# ==============================================================================
# 1. IMPORT LIBRARY
# ==============================================================================
# Bagian ini berisi semua library eksternal yang dibutuhkan aplikasi.
# Pastikan semua library ini sudah ter-install (cek di requirements.txt).
# ==============================================================================

import os  # Untuk berinteraksi dengan sistem operasi (misal: buat folder, cek path).
import logging  # Untuk mencatat log/catatan aktivitas atau error aplikasi.
import glob  # Untuk mencari file di folder berdasarkan pola (misal: *.dbf).
from flask import Flask, render_template, request, send_file, redirect, url_for, flash  # Komponen inti dari Flask untuk membuat web.
from dbfread import DBF  # Library khusus untuk membaca file .dbf.
from custom_parser import CustomFieldParser  # Mengimpor parser custom dari file custom_parser.py.
import pandas as pd  # Library powerful untuk manipulasi data, terutama untuk Excel dan CSV.
from werkzeug.utils import secure_filename  # Fungsi untuk mengamankan nama file yang di-upload.

# REPORTLAB (digunakan untuk membuat file PDF)
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import cm

# Library untuk membuka browser secara otomatis saat aplikasi dijalankan.
import webbrowser
import threading
import time


# ==============================================================================
# 2. KONFIGURASI APLIKASI
# ==============================================================================
# Bagian ini adalah tempat untuk mengatur variabel-variabel penting
# yang akan digunakan di seluruh aplikasi.
# ==============================================================================

# Inisialisasi aplikasi Flask.
app = Flask(__name__)

# Kunci rahasia (secret key) untuk mengamankan session dan flash messages.
# Ganti dengan string acak yang lebih kompleks untuk produksi.
app.secret_key = "supersecret"

# Nama folder tempat menyimpan file yang diunggah oleh user.
UPLOAD_FOLDER = "uploads"
# Membuat folder 'uploads' jika belum ada.
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# Membuat folder untuk file CSS jika belum ada (best practice).
os.makedirs("static/css", exist_ok=True)

# Setup logging dasar untuk menampilkan info atau error di terminal.
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Menentukan ekstensi file apa saja yang diizinkan untuk di-upload.
# Jika ingin menambah format lain (misal: .csv), tambahkan di sini.
ALLOWED_EXTENSIONS = {'dbf', 'xlsx'}

# PETA KOLOM (COLUMN MAPPING)
# Ini adalah bagian PENTING untuk konfigurasi tampilan.
# Gunanya untuk menerjemahkan nama kolom asli di file data (kiri)
# menjadi nama kolom yang lebih rapi dan mudah dibaca (kanan) saat ditampilkan di web atau PDF.
# Jika nama kolom di file sumber berubah, cukup ubah bagian kiri di sini.
COLUMN_MAPPING = {
    "NOPEG": "No. Pegawai",
    "NAMA": "Nama Karyawan",
    "BAGIAN": "Divisi",
    "JML": "Total Pinjaman (Rp)",
    "LAMA": "Total Tenor (Bln)",
    "ANGSURAN_KE": "Pembayaran",
    "SISA_ANGSURAN": "Sisa Tenor (Bln)",
    "SISA_CICILAN": "Sisa Pinjaman (Rp)",
    "STATUS": "Status"
}


# ==============================================================================
# 3. FUNGSI-FUNGSI BANTUAN (HELPER FUNCTIONS)
# ==============================================================================
# Berisi fungsi-fungsi kecil yang melakukan tugas spesifik,
# seperti membaca file, normalisasi data, dll.
# ==============================================================================

def allowed_file(filename):
    """Fungsi untuk memeriksa apakah file yang di-upload punya ekstensi yang diizinkan."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_dbf_file(path):
    """Membaca satu file DBF dan mengembalikannya sebagai list of dictionary."""
    try:
        # Menggunakan parser custom untuk mengatasi masalah data numerik yang aneh.
        table = DBF(path, encoding="latin1", parserclass=CustomFieldParser)
        return [dict(rec) for rec in table]
    except Exception as e:
        logger.error(f"Gagal membaca file DBF {path}: {e}")
        return []

def read_excel_file(path):
    """Membaca satu file Excel dan mengembalikannya sebagai list of dictionary."""
    try:
        df = pd.read_excel(path)
        return df.to_dict(orient="records")
    except Exception as e:
        logger.error(f"Gagal membaca file Excel {path}: {e}")
        return []

def normalize_row(row):
    """
    Fungsi krusial untuk membersihkan dan menghitung data per baris (per pinjaman).
    Fungsi ini mengambil satu record pinjaman, lalu:
    1. Membersihkan data (spasi, tipe data).
    2. Mencari nama kolom alternatif (misal JML, JML_DDL, JUMLAH).
    3. Menghitung kolom turunan seperti 'ANGSURAN_KE', 'SISA_ANGSURAN', 'SISA_CICILAN', dan 'STATUS'.
    """
    r = dict(row)
    angsuran_terbayar = 0
    # Loop untuk menghitung berapa kali angsuran sudah dibayar.
    # Logikanya: setiap kolom yang namanya diawali 'ANG' dan punya nilai dianggap sebagai pembayaran.
    for k, v in r.items():
        if str(k).upper().startswith("ANG") and v not in (None, "", b"", 0):
            try:
                # Memastikan nilai 0 tidak dihitung sebagai pembayaran.
                if isinstance(v, (int, float)) and v == 0:
                    continue
            except:
                pass
            angsuran_terbayar += 1

    # Membersihkan dan memastikan tipe data untuk kolom utama.
    r["NOPEG"] = str(r.get("NOPEG") or "").strip()
    r["NAMA"] = str(r.get("NAMA") or "").strip()
    r["BAGIAN"] = str(r.get("BAGIAN") or "").strip()
    
    # Mencoba mengambil nilai dari beberapa kemungkinan nama kolom untuk 'Jumlah Pinjaman'.
    try:
        r["JML"] = float(r.get("JML") or r.get("JML_DDL") or r.get("JUMLAH") or 0)
    except (ValueError, TypeError):
        r["JML"] = 0

    # Mencoba mengambil nilai untuk 'Lama Pinjaman' (tenor).
    try:
        r["LAMA"] = int(r.get("LAMA") or 0)
    except (ValueError, TypeError):
        r["LAMA"] = 0

    # Mencoba mengambil nilai untuk 'Cicilan per bulan'.
    try:
        r["CICIL"] = float(r.get("CICIL") or r.get("BUNGA1") or r.get("CICILAN") or 0)
    except (ValueError, TypeError):
        r["CICIL"] = 0

    # Menghitung kolom-kolom baru berdasarkan data yang sudah dibersihkan.
    r["ANGSURAN_KE"] = angsuran_terbayar
    r["SISA_ANGSURAN"] = max(r["LAMA"] - angsuran_terbayar, 0) # max(..., 0) agar tidak minus.
    r["SISA_CICILAN"] = r["SISA_ANGSURAN"] * r["CICIL"]

    # Menentukan status pinjaman berdasarkan logika pembayaran.
    if angsuran_terbayar == 0:
        r["STATUS"] = "Belum Bayar"
    elif r["SISA_ANGSURAN"] <= 0 and r["LAMA"] > 0:
        r["STATUS"] = "Lunas"
    else:
        r["STATUS"] = "Berjalan"

    return r

def load_data():
    """
    Fungsi inti untuk memuat SEMUA data dari folder /uploads, lalu menggabungkannya.
    Prosesnya:
    1. Cari semua file .dbf dan .xlsx di folder UPLOAD_FOLDER.
    2. Baca setiap file satu per satu.
    3. Normalisasi setiap baris data menggunakan `normalize_row`.
    4. Kelompokkan semua pinjaman berdasarkan NOPEG.
    5. Untuk setiap NOPEG, jumlahkan total pinjaman, sisa pinjaman, dll. (proses agregasi).
    6. Hasil akhirnya adalah sebuah list yang siap ditampilkan di web.
    """
    try:
        # Mencari semua file dbf dan xlsx.
        pattern_dbf = os.path.join(UPLOAD_FOLDER, "*.dbf")
        pattern_xlsx = os.path.join(UPLOAD_FOLDER, "*.xlsx")
        files = glob.glob(pattern_dbf) + glob.glob(pattern_xlsx)

        if not files:
            return []  # Jika tidak ada file, kembalikan list kosong.

        # Dictionary untuk menampung semua pinjaman, dikelompokkan per NOPEG.
        # Format: {'NOPEG1': [loan1, loan2], 'NOPEG2': [loan3]}
        all_loans_by_nopeg = {}
        for path in files:
            raw_data = []
            if path.lower().endswith(".dbf"):
                raw_data = read_dbf_file(path)
            elif path.lower().endswith(".xlsx"):
                raw_data = read_excel_file(path)

            for rec in raw_data:
                proc = normalize_row(rec)
                nopeg = proc.get("NOPEG")
                if not nopeg:
                    continue  # Lewati baris data jika tidak ada NOPEG.
                
                # Masukkan data pinjaman yang sudah diproses ke dictionary.
                if nopeg not in all_loans_by_nopeg:
                    all_loans_by_nopeg[nopeg] = []
                all_loans_by_nopeg[nopeg].append(proc)

        # Proses Agregasi: Menggabungkan data pinjaman untuk setiap NOPEG.
        final_data = []
        for nopeg, loans in all_loans_by_nopeg.items():
            # Menjumlahkan semua nilai dari setiap pinjaman yang dimiliki satu NOPEG.
            summary = {
                "JML": sum(l['JML'] for l in loans),
                "LAMA": sum(l['LAMA'] for l in loans),
                "ANGSURAN_KE": sum(l['ANGSURAN_KE'] for l in loans),
                "SISA_ANGSURAN": sum(l['SISA_ANGSURAN'] for l in loans),
                "SISA_CICILAN": sum(l['SISA_CICILAN'] for l in loans)
            }
            
            # Menentukan status gabungan. Prioritas: Berjalan > Belum Bayar > Lunas.
            statuses = {l['STATUS'] for l in loans}
            if "Berjalan" in statuses:
                summary['STATUS'] = "Berjalan"
            elif "Belum Bayar" in statuses:
                 summary['STATUS'] = "Belum Bayar"
            else:
                summary['STATUS'] = "Lunas"

            # Struktur data final untuk satu orang.
            # 'SUMMARY' berisi data gabungan, 'DETAILS' berisi rincian setiap pinjaman.
            person_data = {
                "NOPEG": nopeg,
                "NAMA": loans[-1]['NAMA'],  # Ambil nama & bagian dari data pinjaman terakhir.
                "BAGIAN": loans[-1]['BAGIAN'],
                "SUMMARY": summary,
                "DETAILS": loans
            }
            final_data.append(person_data)
            
        return final_data

    except Exception as e:
        logger.error(f"Error saat memuat dan memproses data: {str(e)}")
        return []


# ==============================================================================
# 4. ROUTES / ENDPOINTS
# ==============================================================================
# Bagian ini mendefinisikan URL-URL yang bisa diakses oleh user di browser.
# Setiap fungsi di bawah ini terhubung ke satu URL.
# Contoh: def index() akan dipanggil saat user membuka http://127.0.0.1:5000/
# ==============================================================================

@app.route("/", methods=["GET"])
def index():
    """
    Route untuk halaman utama (homepage).
    Fungsi ini akan:
    1. Memuat semua data dengan memanggil `load_data()`.
    2. Menerima input filter dari URL (search, bagian, status).
    3. Melakukan filter pada data berdasarkan input.
    4. Mengimplementasikan pagination (membagi data menjadi beberapa halaman).
    5. Mengirim data yang sudah difilter dan dipaginasi ke template 'index.html' untuk ditampilkan.
    """
    try:
        all_data = load_data()

        # Mengambil parameter dari URL, contoh: /?search=taufiq&page=2
        q = request.args.get("search", "").strip().lower()
        bagian_filter = request.args.get("bagian", "").strip()
        status_filter = request.args.get("status", "").strip()
        page = int(request.args.get("page", 1))
        per_page = 20  # KONFIGURASI: Ubah angka ini untuk mengatur jumlah data per halaman.

        # Proses filtering data.
        filtered = all_data
        if q:
            # Filter berdasarkan nama atau nopeg (case-insensitive).
            filtered = [r for r in filtered if q in (r.get("NAMA") or "").lower() or q in (r.get("NOPEG") or "").lower()]
        if bagian_filter:
            # Filter berdasarkan bagian/divisi.
            filtered = [r for r in filtered if (r.get("BAGIAN") or "").lower() == bagian_filter.lower()]
        if status_filter:
            # Filter berdasarkan status pinjaman.
            filtered = [r for r in filtered if (r.get("SUMMARY", {}).get("STATUS") or "").lower() == status_filter.lower()]

        # Logika Pagination.
        total_data = len(filtered)
        total_pages = (total_data + per_page - 1) // per_page
        start = (page - 1) * per_page
        end = start + per_page
        paginated_data = filtered[start:end]

        # Mengambil daftar unik semua bagian/divisi untuk ditampilkan di dropdown filter.
        bagian_list = sorted({(r.get("BAGIAN") or "").strip() for r in all_data if (r.get("BAGIAN") or "").strip()})
        
        # 'render_template' adalah jembatan antara Python dan HTML.
        # Variabel di sebelah kanan (e.g., data=paginated_data) akan dikirim
        # dan bisa diakses di dalam file 'index.html'.
        return render_template(
            "index.html", data=paginated_data, bagian_list=bagian_list,
            search=q, bagian_selected=bagian_filter, status_selected=status_filter,
            page=page, total_pages=total_pages, title="Data Koperasi Karyawan",
            column_headers=COLUMN_MAPPING
        )
    except Exception as e:
        logger.error(f"Error di halaman utama: {str(e)}")
        flash("Terjadi kesalahan fatal saat memuat data. Silakan cek log.", "danger")
        return render_template("index.html", data=[], bagian_list=[], page=1, total_pages=1)

@app.route("/reset_data", methods=["POST"])
def reset_data():
    """
    Route untuk menghapus semua file data (.dbf, .xlsx) di folder /uploads.
    Ini memberikan user cara untuk memulai dari awal.
    Hanya bisa diakses dengan metode POST (dari tombol di form).
    """
    try:
        files = glob.glob(os.path.join(UPLOAD_FOLDER, "*"))
        count = 0
        for f in files:
            if f.lower().endswith(('.dbf', '.xlsx')):
                os.remove(f)
                count += 1
        flash(f"Berhasil mereset data. Sebanyak {count} file data telah dihapus.", "success")
    except Exception as e:
        logger.error(f"Error saat mereset data: {str(e)}")
        flash("Gagal mereset data.", "danger")
    
    # Setelah selesai, kembalikan user ke halaman utama.
    return redirect(url_for("index"))

@app.route("/import", methods=["POST"])
def import_file():
    """
    Route untuk menangani proses upload file.
    Bisa menerima beberapa file sekaligus.
    """
    if "file" not in request.files:
        flash("Tidak ada file yang dipilih untuk di-upload.", "danger")
        return redirect(url_for("index"))

    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        flash("Tidak ada file yang dipilih atau nama file kosong.", "danger")
        return redirect(url_for("index"))

    saved_files = []
    errors = []

    for file in files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            if not allowed_file(filename):
                errors.append(f"{filename}: Format file tidak didukung (hanya .dbf atau .xlsx).")
                continue
            
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            
            # Mencegah menimpa file yang sudah ada dengan menambahkan timestamp.
            if os.path.exists(filepath):
                base, ext = os.path.splitext(filename)
                ts = int(time.time())
                filename = f"{base}_{ts}{ext}"
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                
            try:
                file.save(filepath)
                saved_files.append(filename)
            except Exception as e:
                logger.error(f"Error saat menyimpan file {filename}: {e}")
                errors.append(f"{filename}: Gagal menyimpan file di server.")

    # Memberikan notifikasi (flash message) kepada user tentang hasil proses upload.
    if saved_files:
        flash(f"Berhasil mengunggah: {', '.join(saved_files)}. Data akan otomatis ditambahkan dan digabungkan.", "success")
    if errors:
        flash("Beberapa file gagal diunggah: " + "; ".join(errors), "warning")

    return redirect(url_for("index"))

# --- Rute Ekspor Data ---

@app.route("/export/csv")
def export_csv():
    """Route untuk mengekspor data gabungan ke format CSV."""
    try:
        data = load_data()
        # Meratakan struktur data (dari summary dan details menjadi satu baris per orang).
        flat_data = []
        for item in data:
            row = {"NOPEG": item["NOPEG"], "NAMA": item["NAMA"], "BAGIAN": item["BAGIAN"]}
            row.update(item["SUMMARY"])
            flat_data.append(row)

        df = pd.DataFrame(flat_data)
        # Mengatur urutan kolom dan mengganti nama kolom sesuai COLUMN_MAPPING.
        df = df[list(COLUMN_MAPPING.keys())]
        df = df.rename(columns=COLUMN_MAPPING)
        
        filename = "export_data_koperasi.csv"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        # Menyimpan DataFrame ke file CSV.
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
        # Mengirim file tersebut ke browser user untuk di-download.
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"Error saat ekspor CSV: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke CSV.", "danger")
        return redirect(url_for("index"))

@app.route("/export/excel")
def export_excel():
    """Route untuk mengekspor data gabungan ke format Excel."""
    try:
        data = load_data()
        flat_data = []
        for item in data:
            row = {"NOPEG": item["NOPEG"], "NAMA": item["NAMA"], "BAGIAN": item["BAGIAN"]}
            row.update(item["SUMMARY"])
            flat_data.append(row)
            
        df = pd.DataFrame(flat_data)
        df = df[list(COLUMN_MAPPING.keys())]
        df = df.rename(columns=COLUMN_MAPPING)
        
        filename = "export_data_koperasi.xlsx"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        df.to_excel(filepath, index=False)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"Error saat ekspor Excel: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke Excel.", "danger")
        return redirect(url_for("index"))

@app.route("/export/pdf")
def export_pdf():
    """Route untuk mengekspor data gabungan ke format PDF menggunakan ReportLab."""
    try:
        data = load_data()
        filename = "export_data_koperasi.pdf"
        filepath = os.path.join(UPLOAD_FOLDER, filename)

        # Inisialisasi dokumen PDF dengan ukuran A4 potrait.
        doc = SimpleDocTemplate(
            filepath, pagesize=portrait(A4),
            rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm
        )
        
        # Pengaturan style untuk teks (judul, isi tabel, header).
        styles = getSampleStyleSheet()
        style_title = ParagraphStyle(name='Title', parent=styles['h1'], alignment=TA_CENTER, spaceAfter=12)
        style_body_left = ParagraphStyle(name='BodyLeft', parent=styles['Normal'], alignment=TA_LEFT, fontSize=7, leading=9)
        style_body_center = ParagraphStyle(name='BodyCenter', parent=styles['Normal'], alignment=TA_CENTER, fontSize=7, leading=9)
        style_header = ParagraphStyle(name='Header', parent=styles['Normal'], alignment=TA_CENTER, fontName='Helvetica-Bold', fontSize=8, textColor=colors.whitesmoke)

        elements = [Paragraph("Data Koperasi Karyawan (Total Gabungan)", style_title)]
        
        # Teks untuk header tabel (mendukung line break dengan <br/>).
        header_text = {
            "NOPEG": "No.<br/>Pegawai", "NAMA": "Nama Karyawan", "BAGIAN": "Divisi",
            "JML": "Total<br/>Pinjaman<br/>(Rp)", "LAMA": "Total<br/>Tenor<br/>(Bln)",
            "ANGSURAN_KE": "Pembayaran", "SISA_ANGSURAN": "Sisa<br/>Tenor<br/>(Bln)",
            "SISA_CICILAN": "Sisa<br/>Pinjaman<br/>(Rp)", "STATUS": "Status"
        }
        
        # Membuat list header dari teks di atas.
        header_keys = list(COLUMN_MAPPING.keys())
        header = [Paragraph(header_text[key], style_header) for key in header_keys]
        table_data = [header]

        # Mengisi data tabel baris per baris.
        for item in data:
            row = item["SUMMARY"]
            # Menggunakan Paragraph untuk setiap sel agar style bisa diterapkan.
            # Format angka (e.g., ,0f) untuk menambahkan pemisah ribuan.
            table_data.append([
                Paragraph(str(item.get("NOPEG", "")), style_body_center),
                Paragraph(str(item.get("NAMA", "")), style_body_left),
                Paragraph(str(item.get("BAGIAN", "")), style_body_left),
                Paragraph(f"{row.get('JML', 0):,.0f}", style_body_center),
                Paragraph(f"{row.get('LAMA', 0)}", style_body_center),
                Paragraph(f"{row.get('ANGSURAN_KE', 0)}", style_body_center),
                Paragraph(f"{row.get('SISA_ANGSURAN', 0)}", style_body_center),
                Paragraph(f"{row.get('SISA_CICILAN', 0):,.0f}", style_body_center),
                Paragraph(str(row.get("STATUS", "")), style_body_center)
            ])
        
        # KONFIGURASI: Lebar setiap kolom dalam PDF. Sesuaikan di sini jika tabel terpotong.
        col_widths = [2.2*cm, 4*cm, 2.3*cm, 2.3*cm, 1.5*cm, 2.2*cm, 1.8*cm, 2.3*cm, 1.5*cm]
        table = Table(table_data, colWidths=col_widths, repeatRows=1) # repeatRows=1 agar header muncul di setiap halaman.
        
        # Pengaturan style tabel (warna background header, garis, padding).
        tbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#343a40")), # Warna header
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
        ])
        table.setStyle(tbl_style)

        # Menambahkan tabel ke dalam dokumen PDF dan membangun filenya.
        elements.append(table)
        doc.build(elements)

        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"Error saat ekspor PDF: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke PDF.", "danger")
        return redirect(url_for("index"))

@app.route("/dashboard")
def dashboard():
    """
    Route untuk halaman dashboard statistik.
    Fungsi ini akan:
    1. Memuat semua data.
    2. Menghitung berbagai metrik: total pinjaman, sisa pinjaman, jumlah karyawan.
    3. Mengagregasi data berdasarkan status dan divisi.
    4. Mencari 10 peminjam terbesar.
    5. Mengirim semua data statistik ini ke template 'dashboard.html' untuk divisualisasikan dengan Chart.js.
    """
    try:
        all_data = load_data()

        # Jika tidak ada data, tampilkan dashboard kosong.
        if not all_data:
            return render_template(
                "dashboard.html", title="Dashboard Ringkasan",
                total_pinjaman=0, sisa_pinjaman=0, total_karyawan=0,
                status_count={}, bagian_count={}, top_borrowers=[], bagian_pinjaman={},
                status_details={}
            )

        # Kalkulasi statistik untuk chart status pinjaman.
        status_count = {"Lunas": 0, "Berjalan": 0, "Belum Bayar": 0}
        status_amount = {"Lunas": 0, "Berjalan": 0, "Belum Bayar": 0}

        for r in all_data:
            status = r["SUMMARY"]["STATUS"]
            if status in status_count:
                status_count[status] += 1
            
            # Hitung total sisa pinjaman untuk setiap status.
            if status == "Lunas":
                status_amount["Lunas"] += 0 # Sisa pinjaman lunas adalah 0.
            else:
                status_amount[status] += r["SUMMARY"]["SISA_CICILAN"]

        total_karyawan = len(all_data)

        # Menyiapkan data detail untuk tooltip di chart, agar lebih informatif.
        status_details = {
            "labels": list(status_count.keys()),
            "counts": list(status_count.values()),
            "amounts": list(status_amount.values()),
            "percentages": [round((count / total_karyawan) * 100, 1) if total_karyawan > 0 else 0 for count in status_count.values()]
        }
 
        # Kalkulasi statistik untuk chart divisi.
        bagian_count_raw = {}
        bagian_pinjaman_raw = {}
        for r in all_data:
            bagian = r.get("BAGIAN") or "Tidak Ada Divisi"
            bagian_count_raw[bagian] = bagian_count_raw.get(bagian, 0) + 1
            bagian_pinjaman_raw[bagian] = bagian_pinjaman_raw.get(bagian, 0) + (r["SUMMARY"].get("JML") or 0)

        # Mengurutkan dan mengambil 10 divisi teratas.
        sorted_bagian_count = sorted(bagian_count_raw.items(), key=lambda item: item[1], reverse=True)
        top_10_bagian_count = dict(sorted_bagian_count[:10])
        
        sorted_bagian_pinjaman = sorted(bagian_pinjaman_raw.items(), key=lambda item: item[1], reverse=True)
        top_10_bagian_pinjaman = dict(sorted_bagian_pinjaman[:10])
        
        # Kalkulasi KPI utama.
        total_pinjaman = sum(r["SUMMARY"].get('JML', 0) for r in all_data)
        sisa_pinjaman = sum(r["SUMMARY"].get('SISA_CICILAN', 0) for r in all_data)
        
        # Mencari 10 peminjam teratas berdasarkan total pinjaman.
        sorted_by_pinjaman = sorted(all_data, key=lambda x: x["SUMMARY"].get("JML", 0), reverse=True)
        top_borrowers = [
            {"nama": r["NAMA"], "jumlah": r["SUMMARY"]["JML"]}
            for r in sorted_by_pinjaman[:10]
        ]
        
        # Mengirim semua hasil kalkulasi ke template 'dashboard.html'.
        return render_template(
            "dashboard.html",
            title="Dashboard Ringkasan",
            status_details=status_details,
            bagian_count=top_10_bagian_count,
            bagian_pinjaman=top_10_bagian_pinjaman,
            total_pinjaman=total_pinjaman,
            sisa_pinjaman=sisa_pinjaman,
            total_karyawan=total_karyawan,
            top_borrowers=top_borrowers
        )
    except Exception as e:
        logger.error(f"Error di halaman dashboard: {str(e)}")
        flash("Terjadi kesalahan saat memuat data dashboard.", "danger")
        return redirect(url_for("index"))

# ==============================================================================
# 5. TITIK MASUK APLIKASI (ENTRY POINT)
# ==============================================================================
# Bagian ini akan dieksekusi ketika kamu menjalankan `python app.py` di terminal.
# ==============================================================================
if __name__ == "__main__":
    # Fungsi kecil untuk membuka browser secara otomatis setelah server siap.
    def open_browser():
        webbrowser.open("http://127.0.0.1:5000/")
    
    # Menjalankan fungsi open_browser setelah 1 detik.
    threading.Timer(1.0, open_browser).start()
    
    # Menjalankan server Flask.
    # debug=False lebih cocok untuk "produksi" atau demo. Ganti ke True jika sedang development.
    app.run(debug=False)