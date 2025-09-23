"""
app.py
------
Aplikasi utama Flask.
- Load data dari DBF / Excel
- Hitung angsuran & status
- Tampilkan di halaman web
- Export ke Excel, CSV, PDF
- Import file (replace data lama)
"""

import os
import logging
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from dbfread import DBF
from custom_parser import CustomFieldParser
import pandas as pd
from werkzeug.utils import secure_filename

# REPORTLAB (digunakan untuk export PDF)
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

# =============================
# Config
# =============================
app = Flask(__name__)
app.secret_key = "supersecret"  # buat flash message
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ekstensi file yang diizinkan
ALLOWED_EXTENSIONS = {'dbf', 'xlsx'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Default file DBF
current_file = os.path.join(UPLOAD_FOLDER, "HUTKOP1.DBF")
if not os.path.exists(current_file):
    # kalau belum ada, copy dari root
    if os.path.exists("HUTKOP1.DBF"):
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        os.system(f"copy HUTKOP1.DBF {UPLOAD_FOLDER}")

# =============================
# Load Data Function
# =============================
def load_data():
    """
    Load data dari file DBF/Excel sesuai current_file,
    lalu bersihkan field, hitung angsuran & status.
    """
    global current_file
    rows = []

    try:
        if not os.path.exists(current_file):
            logger.error(f"File not found: {current_file}")
            return rows

        if current_file.lower().endswith(".dbf"):
            table = DBF(current_file, encoding="latin1", parserclass=CustomFieldParser)
            raw_data = [dict(rec) for rec in table]
        elif current_file.lower().endswith(".xlsx"):
            raw_data = pd.read_excel(current_file).to_dict(orient="records")
        else:
            logger.error(f"Unsupported file format: {current_file}")
            return rows

        for rec in raw_data:
            try:
                row = dict(rec)

                # Hitung angsuran terbayar
                angsuran_terbayar = 0
                for i in range(1, 101):
                    key = f"ANG{i}"
                    v = row.get(key)
                    if v not in (None, "", b"", 0):
                        angsuran_terbayar += 1
                row["ANGSURAN_KE"] = angsuran_terbayar

                # Field utama dengan default values
                row["NOPEG"] = str(row.get("NOPEG") or "").strip()
                row["NAMA"] = str(row.get("NAMA") or "").strip()
                row["BAGIAN"] = str(row.get("BAGIAN") or "").strip()
                
                # Pastikan nilai numerik valid
                try:
                    row["JML"] = float(row.get("JML") or 0)
                except (ValueError, TypeError):
                    row["JML"] = 0
                    
                try:
                    row["LAMA"] = int(row.get("LAMA") or 0)
                except (ValueError, TypeError):
                    row["LAMA"] = 0
                    
                try:
                    row["CICIL"] = float(row.get("CICIL") or 0)
                except (ValueError, TypeError):
                    row["CICIL"] = 0

                # Hitungan tambahan
                try:
                    row["SISA_ANGSURAN"] = max(int(row.get("LAMA") or 0) - int(angsuran_terbayar), 0)
                except Exception:
                    row["SISA_ANGSURAN"] = 0

                try:
                    row["SISA_CICILAN"] = int(row["SISA_ANGSURAN"]) * int(row["CICIL"])
                except Exception:
                    row["SISA_CICILAN"] = 0

                # Status
                if angsuran_terbayar == 0:
                    row["STATUS"] = "Belum Bayar"
                elif angsuran_terbayar >= row["LAMA"]:
                    row["STATUS"] = "Lunas"
                else:
                    row["STATUS"] = "Berjalan"

                rows.append(row)
            except Exception as e:
                logger.error(f"Error processing row: {rec}. Error: {str(e)}")
                continue

    except Exception as e:
        logger.error(f"Error loading data from {current_file}: {str(e)}")
        flash(f"Error loading data: {str(e)}", "danger")

    return rows


# =============================
# Routes
# =============================

@app.route("/", methods=["GET"])
def index():
    """Halaman utama + filter & pagination"""
    try:
        all_data = load_data()

        # Filter
        q = request.args.get("search", "").strip().lower()
        bagian_filter = request.args.get("bagian", "").strip()
        status_filter = request.args.get("status", "").strip()
        page = int(request.args.get("page", 1))
        per_page = 20

        filtered = all_data
        if q:
            filtered = [
                r for r in filtered
                if q in (r.get("NAMA") or "").lower() or q in (r.get("NOPEG") or "").lower()
            ]
        if bagian_filter and bagian_filter.lower() != "all":
            filtered = [r for r in filtered if (r.get("BAGIAN") or "").lower() == bagian_filter.lower()]
        if status_filter and status_filter.lower() != "all":
            filtered = [r for r in filtered if (r.get("STATUS") or "").lower() == status_filter.lower()]

        # Pagination
        total_data = len(filtered)
        total_pages = (total_data + per_page - 1) // per_page
        start = (page - 1) * per_page
        end = start + per_page
        paginated_data = filtered[start:end]

        # Bagian unik
        bagian_list = sorted({
            (r.get("BAGIAN") or "").strip()
            for r in all_data if (r.get("BAGIAN") or "").strip()
        })

        # ✅ Tambahan: hitung data untuk grafik
        status_count = {
            "Lunas": sum(1 for r in filtered if r["STATUS"] == "Lunas"),
            "Berjalan": sum(1 for r in filtered if r["STATUS"] == "Berjalan"),
            "Belum Bayar": sum(1 for r in filtered if r["STATUS"] == "Belum Bayar"),
        }

        bagian_count = {}
        for r in filtered:
            bagian = r.get("BAGIAN") or "Tidak Ada"
            bagian_count[bagian] = bagian_count.get(bagian, 0) + 1

        return render_template(
            "index.html",
            data=paginated_data,
            bagian_list=bagian_list,
            search=q,
            bagian_selected=bagian_filter,
            status_selected=status_filter,
            page=page,
            total_pages=total_pages,
            title="Data Koperasi Karyawan",
            status_count=status_count,
            bagian_count=bagian_count
        )
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        flash("Terjadi kesalahan saat memuat data", "danger")
        return render_template("index.html", data=[], bagian_list=[], page=1, total_pages=1)


@app.route("/import", methods=["POST"])
def import_file():
    """Import file DBF/Excel dari UI, replace data lama"""
    global current_file
    
    if "file" not in request.files:
        flash("Tidak ada file diupload", "danger")
        return redirect(url_for("index"))

    file = request.files["file"]
    if file.filename == "":
        flash("Nama file kosong", "danger")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Format file tidak didukung. Hanya file DBF dan Excel yang diperbolehkan.", "danger")
        return redirect(url_for("index"))

    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        # Validasi file sebelum menggunakannya
        if filename.lower().endswith(".dbf"):
            # Coba baca file DBF untuk memvalidasi
            test_table = DBF(filepath, encoding="latin1", parserclass=CustomFieldParser)
            test_data = [dict(rec) for rec in test_table]
            if not test_data:
                flash("File DBF tidak berisi data atau format tidak valid", "danger")
                os.remove(filepath)
                return redirect(url_for("index"))
                
        elif filename.lower().endswith(".xlsx"):
            # Coba baca file Excel untuk memvalidasi
            test_data = pd.read_excel(filepath)
            if test_data.empty:
                flash("File Excel tidak berisi data atau format tidak valid", "danger")
                os.remove(filepath)
                return redirect(url_for("index"))

        # Update global file jika validasi berhasil
        current_file = filepath
        flash(f"Berhasil import {filename}. Data telah dimuat.", "success")
        
    except Exception as e:
        logger.error(f"Error importing file: {str(e)}")
        flash(f"Terjadi kesalahan saat mengimpor file: {str(e)}", "danger")
        # Hapus file jika terjadi error
        if 'filepath' in locals() and os.path.exists(filepath):
            os.remove(filepath)

    return redirect(url_for("index"))


# =============================
# Export Routes
# =============================

@app.route("/export/csv")
def export_csv():
    try:
        data = load_data()
        # Pilih hanya kolom hasil
        df = pd.DataFrame(data)[[
            "NOPEG", "NAMA", "BAGIAN", "JML", "LAMA", "CICIL",
            "ANGSURAN_KE", "SISA_ANGSURAN", "SISA_CICILAN", "STATUS"
        ]]
        # Rename biar sama persis kaya PDF
        df = df.rename(columns={
            "JML": "JUMLAH",
            "CICIL": "CICILAN",
            "SISA_ANGSURAN": "SISA",
            "SISA_CICILAN": "SISA CICILAN"
        })
        filename = os.path.join(UPLOAD_FOLDER, "export.csv")
        df.to_csv(filename, index=False, encoding="utf-8-sig")
        return send_file(filename, as_attachment=True, download_name="export_data.csv")
    except Exception as e:
        logger.error(f"Error exporting CSV: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke CSV", "danger")
        return redirect(url_for("index"))

@app.route("/export/excel")
def export_excel():
    try:
        data = load_data()
        # Pilih hanya kolom hasil
        df = pd.DataFrame(data)[[
            "NOPEG", "NAMA", "BAGIAN", "JML", "LAMA", "CICIL",
            "ANGSURAN_KE", "SISA_ANGSURAN", "SISA_CICILAN", "STATUS"
        ]]
        # Rename biar sama persis kaya PDF
        df = df.rename(columns={
            "JML": "JUMLAH",
            "CICIL": "CICILAN",
            "SISA_ANGSURAN": "SISA",
            "SISA_CICILAN": "SISA CICILAN"
        })
        filename = os.path.join(UPLOAD_FOLDER, "export.xlsx")
        df.to_excel(filename, index=False)
        return send_file(filename, as_attachment=True, download_name="export_data.xlsx")
    except Exception as e:
        logger.error(f"Error exporting Excel: {str(e)}")
        flash("Terjadi kesalahan saat mengekspor data ke Excel", "danger")
        return redirect(url_for("index"))

@app.route("/export/pdf")
def export_pdf():
    try:
        data = load_data()
        filename = "export_data.pdf"

        # 🔥 pake landscape + margin biar tabel muat di kertas
        doc = SimpleDocTemplate(
            filename,
            pagesize=landscape(A4),
            rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
        )
        styles = getSampleStyleSheet()
        elements = []

        header = ["NOPEG", "NAMA", "BAGIAN", "JUMLAH", "LAMA", "CICILAN",
                "ANGSURAN KE", "SISA", "SISA CICILAN", "STATUS"]
        table_data = [header]

        for row in data:
            table_data.append([
                row["NOPEG"],
                row["NAMA"],
                row["BAGIAN"],
                f"{row['JML']:,}",
                row["LAMA"],
                f"{row['CICIL']:,}",
                row["ANGSURAN_KE"],
                row["SISA_ANGSURAN"],
                f"{row['SISA_CICILAN']:,}",
                row["STATUS"],
            ])

        col_widths = [
            2.8*cm, 5.0*cm, 3.0*cm, 2.0*cm, 1.5*cm,
            2.0*cm, 2.8*cm, 1.4*cm, 2.5*cm, 1.8*cm
        ]

        table = Table(table_data, repeatRows=1, colWidths=col_widths, hAlign="CENTER")

        tbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#343a40")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),

            ("ALIGN", (0, 0), (2, -1), "LEFT"),
            ("ALIGN", (3, 0), (3, -1), "CENTER"),
            ("ALIGN", (5, 0), (5, -1), "CENTER"),
            ("ALIGN", (8, 0), (8, -1), "CENTER"),
            ("ALIGN", (4, 0), (4, -1), "CENTER"),
            ("ALIGN", (6, 0), (7, -1), "CENTER"),
            ("ALIGN", (9, 0), (9, -1), "CENTER"),

            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 0), (-1, 0), 6),
        ])
        table.setStyle(tbl_style)

        elements.append(Paragraph("Data Koperasi Karyawan", styles["Title"]))
        elements.append(table)
        doc.build(elements)

        return send_file(filename, as_attachment=True)
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
            flash("Data masih kosong, silakan import file dulu.", "warning")
            return render_template(
                "dashboard.html",
                title="Dashboard Ringkasan",
                status_count={},
                bagian_count={},
                total_cicilan=0,
                avg_cicilan=0,
                total_karyawan=0
            )

        # Ringkasan status
        status_count = {
            "Lunas": sum(1 for r in all_data if r["STATUS"] == "Lunas"),
            "Berjalan": sum(1 for r in all_data if r["STATUS"] == "Berjalan"),
            "Belum Bayar": sum(1 for r in all_data if r["STATUS"] == "Belum Bayar"),
        }

        # Ringkasan bagian/divisi
        bagian_count = {}
        for r in all_data:
            bagian = r.get("BAGIAN") or "Tidak Ada"
            bagian_count[bagian] = bagian_count.get(bagian, 0) + 1

        # ✅ Ringkasan angka besar
        total_cicilan = sum(r.get("CICIL") or 0 for r in all_data)
        avg_cicilan = round(total_cicilan / len(all_data), 2) if all_data else 0
        total_karyawan = len(all_data)

        return render_template(
            "dashboard.html",
            title="Dashboard Ringkasan",
            status_count=status_count,
            bagian_count=bagian_count,
            total_cicilan=total_cicilan,
            avg_cicilan=avg_cicilan,
            total_karyawan=total_karyawan
        )
    except Exception as e:
        logger.error(f"Error in dashboard route: {str(e)}")
        flash("Terjadi kesalahan saat memuat dashboard", "danger")
        return redirect(url_for("index"))



if __name__ == "__main__":
    app.run(debug=True)