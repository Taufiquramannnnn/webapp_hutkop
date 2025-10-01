"""
custom_parser.py
----------------
File ini berisi class 'CustomFieldParser' yang di-wariskan (inherit)
dari FieldParser milik library dbfread. Tujuannya adalah untuk
membuat aturan parsing sendiri yang lebih tangguh.

KENAPA INI DIPERLUKAN?
File DBF (dBase) adalah format file yang sudah tua. Kadang, data numerik (angka)
di dalamnya tidak bersih. Bisa jadi mengandung:
- Karakter null (b'\x00') yang tidak terlihat.
- Spasi ekstra.
- Format desimal yang menggunakan koma (,) khas Indonesia, bukan titik (.).

Parser bawaan dbfread bisa gagal membaca data seperti ini.
Dengan parser custom ini, kita bisa membersihkan data tersebut sebelum
diubah menjadi tipe data integer atau float di Python.
"""

from dbfread import FieldParser
import logging

# Setup logger khusus untuk file ini.
logger = logging.getLogger(__name__)

class CustomFieldParser(FieldParser):
    def parseN(self, field, data):
        """
        Method ini adalah "override" dari method asli di dbfread.
        Method `parseN` secara spesifik dipanggil oleh dbfread setiap kali
        ia menemukan kolom dengan tipe 'N' (Numeric).

        Alur logika parsing custom:
        1. Terima data mentah dalam bentuk bytes (e.g., b'  123,45\x00').
        2. Bersihkan karakter null (b'\x00') dan spasi di awal/akhir.
        3. Jika data kosong setelah dibersihkan, kembalikan None (nilai kosong).
        4. Coba ubah data menjadi integer. Jika berhasil, kembalikan hasilnya.
        5. Jika gagal jadi integer (misal karena ada koma), coba ubah jadi float:
           a. Decode byte string menjadi regular string.
           b. Ganti karakter koma (',') menjadi titik ('.').
           c. Ubah string yang sudah diperbaiki menjadi float.
        6. Jika semua usaha di atas gagal, catat pesan error di log
           dan kembalikan nilai 0 sebagai nilai default agar aplikasi tidak crash.
        """
        try:
            # 1 & 2: Bersihkan data mentah.
            data = data.replace(b'\x00', b'').strip()
            
            # 3: Cek jika data kosong.
            if data == b'':
                return None

            # 4: Coba parsing sebagai integer.
            try:
                return int(data)
            except ValueError:
                # 5: Jika gagal, coba parsing sebagai float.
                try:
                    # 5a & 5b: Decode dan ganti koma.
                    # 'latin1' adalah encoding umum untuk file DBF lama.
                    data_str = data.decode('latin1').replace(',', '.')
                    # 5c: Ubah ke float.
                    return float(data_str)
                except (ValueError, UnicodeDecodeError):
                    # 6: Jika masih gagal, ini adalah langkah pengaman terakhir.
                    logger.warning(f"Gagal mem-parsing nilai numerik: {data}. Dikembalikan sebagai 0.")
                    return 0
        except Exception as e:
            # Menangkap error tak terduga lainnya.
            logger.error(f"Terjadi error pada fungsi parseN: {str(e)}")
            return 0