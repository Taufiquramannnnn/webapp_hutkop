"""
app.py
------
Aplikasi utama Flask.
- Load data dari DBF / Excel (sekarang multiple files di folder uploads/)
- Agregasi (jumlahkan) data untuk NOPEG yang sama dari file berbeda.
- Tampilkan di halaman web, export, dan import.
- Tambahan fitur reset data.
"""

import os
import logging
import glob
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from dbfread import DBF
from custom_parser import CustomFieldParser
import pandas as pd
from werkzeug.utils import secure_filename

# REPORTLAB (digunakan untuk export PDF)
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import cm

import webbrowser
import threading
import time

# =============================
# Config
# =============================
app = Flask(__name__)
app.secret_key = "supersecret"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("static/css", exist_ok=True) # Pastikan folder static/css ada

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ekstensi file yang diizinkan
ALLOWED_EXTENSIONS = {'dbf', 'xlsx'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Nama Kolom untuk Tampilan Web dan PDF
COLUMN_MAPPING = {
    "NOPEG": "No. Pegawai",
    "NAMA": "Nama Karyawan",
    "BAGIAN": "Divisi",
    "JML": "Total Pinjaman (Rp)",
    "LAMA": "Total Tenor (Bln)",
    "CICIL": "Cicilan/Bulan (Rp)",
    "ANGSURAN_KE": "Pembayaran", 
    "SISA_ANGSURAN": "Sisa Tenor (Bln)",
    "SISA_CICILAN": "Sisa Pinjaman (Rp)",
    "STATUS": "Status"
}


# =============================
# Load Data Function
# =============================
def read_dbf_file(path):
    """Baca DBF file dan return list of dict"""
    try:
        table = DBF(path, encoding="latin1", parserclass=CustomFieldParser)
        return [dict(rec) for rec in table]
    except Exception as e:
        logger.error(f"Error reading DBF {path}: {e}")
        return []

def read_excel_file(path):
    """Baca Excel file dan return list of dict"""
    try:
        df = pd.read_excel(path)
        return df.to_dict(orient="records")
    except Exception as e:
        logger.error(f"Error reading Excel {path}: {e}")
        return []

def normalize_row(row):
    """
    Normalize row fields and compute derived fields for a SINGLE loan record.
    """
    r = dict(row)
    angsuran_terbayar = 0
    for k, v in r.items():
        if str(k).upper().startswith("ANG") and v not in (None, "", b"", 0):
            try:
                if isinstance(v, (int, float)) and v == 0:
                    continue
            except:
                pass
            angsuran_terbayar += 1

    r["NOPEG"] = str(r.get("NOPEG") or "").strip()
    r["NAMA"] = str(r.get("NAMA") or "").strip()
    r["BAGIAN"] = str(r.get("BAGIAN") or "").strip()

    try:
        r["JML"] = float(r.get("JML") or r.get("JML_DDL") or r.get("JUMLAH") or 0)
    except (ValueError, TypeError):
        r["JML"] = 0

    try:
        r["LAMA"] = int(r.get("LAMA") or 0)
    except (ValueError, TypeError):
        r["LAMA"] = 0

    try:
        r["CICIL"] = float(r.get("CICIL") or r.get("BUNGA1") or r.get("CICILAN") or 0)
    except (ValueError, TypeError):
        r["CICIL"] = 0

    r["ANGSURAN_KE"] = angsuran_terbayar
    r["SISA_ANGSURAN"] = max(r["LAMA"] - angsuran_terbayar, 0)
    r["SISA_CICILAN"] = r["SISA_ANGSURAN"] * r["CICIL"]

    if angsuran_terbayar == 0:
        r["STATUS"] = "Belum Bayar"
    elif r["SISA_ANGSURAN"] <= 0 and r["LAMA"] > 0:
        r["STATUS"] = "Lunas"
    else:
        r["STATUS"] = "Berjalan"

    return r

def load_data():
    """
    Load data dari semua file di folder uploads/, lalu AGREGASI (jumlahkan)
    data untuk NOPEG yang sama.
    """
    try:
        pattern_dbf = os.path.join(UPLOAD_FOLDER, "*.dbf")
        pattern_xlsx = os.path.join(UPLOAD_FOLDER, "*.xlsx")
        files = glob.glob(pattern_dbf) + glob.glob(pattern_xlsx)

        if not files:
            logger.warning("No data files found in uploads/")
            return []

        data_aggregator = {}

        for path in files:
            logger.info(f"Loading and aggregating file: {path}")
            raw = []
            if path.lower().endswith(".dbf"):
                raw = read_dbf_file(path)
            elif path.lower().endswith(".xlsx"):
                raw = read_excel_file(path)

            for rec in raw:
                proc = normalize_row(rec)
                nopeg = proc.get("NOPEG")
                if not nopeg:
                    continue

                if nopeg not in data_aggregator:
                    data_aggregator[nopeg] = {
                        "NOPEG": proc["NOPEG"], "NAMA": proc["NAMA"], "BAGIAN": proc["BAGIAN"],
                        "JML": proc["JML"], "LAMA": proc["LAMA"], "CICIL": proc["CICIL"],
                        "ANGSURAN_KE": proc["ANGSURAN_KE"], "SISA_ANGSURAN": proc["SISA_ANGSURAN"],
                        "SISA_CICILAN": proc["SISA_CICILAN"],
                        "STATUS_LIST": [proc["STATUS"]]
                    }
                else:
                    agg = data_aggregator[nopeg]
                    agg["JML"] += proc["JML"]
                    agg["LAMA"] += proc["LAMA"]
                    agg["CICIL"] += proc["CICIL"]
                    agg["ANGSURAN_KE"] += proc["ANGSURAN_KE"]
                    agg["SISA_ANGSURAN"] += proc["SISA_ANGSURAN"]
                    agg["SISA_CICILAN"] += proc["SISA_CICILAN"]
                    agg["STATUS_LIST"].append(proc["STATUS"])
                    if proc["NAMA"]: agg["NAMA"] = proc["NAMA"]
                    if proc["BAGIAN"]: agg["BAGIAN"] = proc["BAGIAN"]

        final_rows = []
        for nopeg, agg_data in data_aggregator.items():
            statuses = set(agg_data["STATUS_LIST"])
            if "Berjalan" in statuses:
                agg_data["STATUS"] = "Berjalan"
            elif "Belum Bayar" in statuses:
                 agg_data["STATUS"] = "Belum Bayar"
            else:
                agg_data["STATUS"] = "Lunas"
            
            del agg_data["STATUS_LIST"]
            final_rows.append(agg_data)

        return final_rows

    except Exception as e:
        logger.error(f"Error loading data: {str(e)}")
        return []

# =============================
# Routes
# =============================

@app.route("/", methods=["GET"])
def index():
    """Halaman utama + filter & pagination"""
    try:
        all_data = load_data()

        q = request.args.get("search", "").strip().lower()
        bagian_filter = request.args.get("bagian", "").strip()
        status_filter = request.args.get("status", "").strip()
        page = int(request.args.get("page", 1))
        per_page = 20

        filtered = all_data
        if q:
            filtered = [r for r in filtered if q in (r.get("NAMA") or "").lower() or q in (r.get("NOPEG") or "").lower()]
        if bagian_filter and bagian_filter.lower() != "":
            filtered = [r for r in filtered if (r.get("BAGIAN") or "").lower() == bagian_filter.lower()]
        if status_filter and status_filter.lower() != "":
            filtered = [r for r in filtered if (r.get("STATUS") or "").lower() == status_filter.lower()]

        total_data = len(filtered)
        total_pages = (total_data + per_page - 1) // per_page
        start = (page - 1) * per_page
        end = start + per_page
        paginated_data = filtered[start:end]

        bagian_list = sorted({(r.get("BAGIAN") or "").strip() for r in all_data if (r.get("BAGIAN") or "").strip()})
        
        return render_template(
            "index.html", data=paginated_data, bagian_list=bagian_list,
            search=q, bagian_selected=bagian_filter, status_selected=status_filter,
            page=page, total_pages=total_pages, title="Data Koperasi Karyawan",
            column_headers=COLUMN_MAPPING
        )
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        flash("Terjadi kesalahan saat memuat data", "danger")
        return render_template("index.html", data=[], bagian_list=[], page=1, total_pages=1)

@app.route("/reset_data", methods=["POST"])
def reset_data():
    """Menghapus semua file di folder uploads."""
    try:
        files = glob.glob(os.path.join(UPLOAD_FOLDER, "*"))
        count = 0
        for f in files:
            if f.lower().endswith(('.dbf', '.xlsx')):
                os.remove(f)
                count += 1
        flash(f"Berhasil mereset data. Sebanyak {count} file data telah dihapus.", "success")
    except Exception as e:
        logger.error(f"Error resetting data: {str(e)}")
        flash("Gagal mereset data.", "danger")
    return redirect(url_for("index"))

@app.route("/import", methods=["POST"])
def import_file():
    """Import multiple file DBF/Excel dari UI"""
    if "file" not in request.files:
        flash("Tidak ada file diupload", "danger")
        return redirect(url_for("index"))

    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        flash("Nama file kosong", "danger")
        return redirect(url_for("index"))

    saved_files = []
    errors = []

    for file in files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            if not allowed_file(filename):
                errors.append(f"{filename}: Format tidak didukung")
                continue
            
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            
            if os.path.exists(filepath):
                base, ext = os.path.splitext(filename)
                ts = int(time.time())
                filename = f"{base}_{ts}{ext}"
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                
            try:
                file.save(filepath)
                saved_files.append(filename)
            except Exception as e:
                logger.error(f"Error saving file {filename}: {e}")
                errors.append(f"{filename}: Gagal menyimpan file")

    if saved_files:
        flash(f"Berhasil mengunggah: {', '.join(saved_files)}. Data akan ditambahkan dan dijumlahkan.", "success")
    if errors:
        flash("Beberapa file bermasalah: " + "; ".join(errors), "warning")

    return redirect(url_for("index"))

@app.route("/export/csv")
def export_csv():
    try:
        data = load_data()
        df = pd.DataFrame(data)
        df = df[list(COLUMN_MAPPING.keys())]
        df = df.rename(columns=COLUMN_MAPPING)
        filename = "export_data_koperasi.csv"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"Error exporting CSV: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke CSV", "danger")
        return redirect(url_for("index"))

@app.route("/export/excel")
def export_excel():
    try:
        data = load_data()
        df = pd.DataFrame(data)
        df = df[list(COLUMN_MAPPING.keys())]
        df = df.rename(columns=COLUMN_MAPPING)
        filename = "export_data_koperasi.xlsx"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        df.to_excel(filepath, index=False)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"Error exporting Excel: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke Excel", "danger")
        return redirect(url_for("index"))

@app.route("/export/pdf")
def export_pdf():
    try:
        data = load_data()
        filename = "export_data_koperasi.pdf"
        filepath = os.path.join(UPLOAD_FOLDER, filename)

        doc = SimpleDocTemplate(
            filepath, pagesize=portrait(A4),
            rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm
        )
        
        styles = getSampleStyleSheet()
        style_title = ParagraphStyle(name='Title', parent=styles['h1'], alignment=TA_CENTER, spaceAfter=12)
        style_body_left = ParagraphStyle(name='BodyLeft', parent=styles['Normal'], alignment=TA_LEFT, fontSize=7, leading=9)
        style_body_center = ParagraphStyle(name='BodyCenter', parent=styles['Normal'], alignment=TA_CENTER, fontSize=7, leading=9)
        style_header = ParagraphStyle(name='Header', parent=styles['Normal'], alignment=TA_CENTER, fontName='Helvetica-Bold', fontSize=8, textColor=colors.whitesmoke)

        elements = [Paragraph("Data Koperasi Karyawan (Total Gabungan)", style_title)]
        
        # Wrapping header text for better fit in portrait
        header_text = {
            "NOPEG": "No.<br/>Pegawai",
            "NAMA": "Nama Karyawan",
            "BAGIAN": "Divisi",
            "JML": "Total<br/>Pinjaman<br/>(Rp)",
            "LAMA": "Total<br/>Tenor<br/>(Bln)",
            "CICIL": "Cicilan/<br/>Bulan<br/>(Rp)",
            "ANGSURAN_KE": "Pembayaran", 
            "SISA_ANGSURAN": "Sisa<br/>Tenor<br/>(Bln)",
            "SISA_CICILAN": "Sisa<br/>Pinjaman<br/>(Rp)",
            "STATUS": "Status"
        }

        header = [Paragraph(header_text[key], style_header) for key in COLUMN_MAPPING.keys()]
        table_data = [header]

        for row in data:
            table_data.append([
                Paragraph(str(row.get("NOPEG", "")), style_body_center),
                Paragraph(str(row.get("NAMA", "")), style_body_left),
                Paragraph(str(row.get("BAGIAN", "")), style_body_left),
                Paragraph(f"{row.get('JML', 0):,.0f}", style_body_center),
                Paragraph(f"{row.get('LAMA', 0)}", style_body_center),
                Paragraph(f"{row.get('CICIL', 0):,.0f}", style_body_center),
                Paragraph(f"{row.get('ANGSURAN_KE', 0)}", style_body_center),
                Paragraph(f"{row.get('SISA_ANGSURAN', 0)}", style_body_center),
                Paragraph(f"{row.get('SISA_CICILAN', 0):,.0f}", style_body_center),
                Paragraph(str(row.get("STATUS", "")), style_body_center)
            ])
        
        col_widths = [2.1*cm, 4*cm, 2.2*cm, 2*cm, 1.7*cm, 2*cm, 1.4*cm, 1.4*cm, 1.8*cm, 1.7*cm]
        
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        tbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#343a40")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
        ])
        table.setStyle(tbl_style)

        elements.append(table)
        doc.build(elements)

        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        logger.error(f"Error exporting PDF: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke PDF", "danger")
        return redirect(url_for("index"))

@app.route("/dashboard")
def dashboard():
    """Halaman ringkasan statistik"""
    try:
        all_data = load_data()

        if not all_data:
            return render_template(
                "dashboard.html", title="Dashboard Ringkasan",
                total_pinjaman=0, sisa_pinjaman=0, total_karyawan=0,
                status_count={}, bagian_count={}, top_borrowers=[], bagian_pinjaman={}
            )

        status_count = {
            "Lunas": sum(1 for r in all_data if r["STATUS"] == "Lunas"),
            "Berjalan": sum(1 for r in all_data if r["STATUS"] == "Berjalan"),
            "Belum Bayar": sum(1 for r in all_data if r["STATUS"] == "Belum Bayar")
        }

        bagian_count_raw = {}
        bagian_pinjaman_raw = {}
        for r in all_data:
            bagian = r.get("BAGIAN") or "Tidak Ada"
            bagian_count_raw[bagian] = bagian_count_raw.get(bagian, 0) + 1
            bagian_pinjaman_raw[bagian] = bagian_pinjaman_raw.get(bagian, 0) + (r.get("JML") or 0)

        sorted_bagian_count = sorted(bagian_count_raw.items(), key=lambda item: item[1], reverse=True)
        top_10_bagian_count = dict(sorted_bagian_count[:10])
        
        sorted_bagian_pinjaman = sorted(bagian_pinjaman_raw.items(), key=lambda item: item[1], reverse=True)
        top_10_bagian_pinjaman = dict(sorted_bagian_pinjaman[:10])
        
        total_pinjaman = sum(r.get('JML', 0) for r in all_data)
        sisa_pinjaman = sum(r.get('SISA_CICILAN', 0) for r in all_data)
        total_karyawan = len(all_data)

        sorted_by_pinjaman = sorted(all_data, key=lambda x: x.get("JML", 0), reverse=True)
        top_borrowers = [
            {"nama": r["NAMA"], "jumlah": r["JML"]}
            for r in sorted_by_pinjaman[:10]
        ]
        
        return render_template(
            "dashboard.html",
            title="Dashboard Ringkasan",
            status_count=status_count,
            bagian_count=top_10_bagian_count,
            bagian_pinjaman=top_10_bagian_pinjaman,
            total_pinjaman=total_pinjaman,
            sisa_pinjaman=sisa_pinjaman,
            total_karyawan=total_karyawan,
            top_borrowers=top_borrowers
        )
    except Exception as e:
        logger.error(f"Error in dashboard route: {str(e)}")
        flash("Terjadi kesalahan saat memuat dashboard", "danger")
        return redirect(url_for("index"))

if __name__ == "__main__":
    def open_browser():
        webbrowser.open("http://127.0.0.1:5000/")
    threading.Timer(1.0, open_browser).start()
    app.run(debug=False)