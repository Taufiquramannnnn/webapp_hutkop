"""
custom_parser.py
----------------
Parser custom untuk file DBF agar kolom numeric yang "aneh"
bisa dibaca dengan benar ke int/float.
"""

from dbfread import FieldParser, InvalidValue
import logging

logger = logging.getLogger(__name__)

class CustomFieldParser(FieldParser):
    def parseN(self, field, data):
        """
        Parsing field numeric custom.
        - Bersihin null char (b'\x00') dan spasi
        - Coba ubah ke int
        - Kalau gagal, coba ke float (ganti koma jadi titik)
        """
        try:
            data = data.replace(b'\x00', b'').strip()
            if data == b'':
                return None

            # Coba parsing sebagai integer
            try:
                return int(data)
            except ValueError:
                # Coba parsing sebagai float
                try:
                    # Ganti koma dengan titik untuk format desimal Indonesia
                    data_str = data.decode('latin1').replace(',', '.')
                    return float(data_str)
                except (ValueError, UnicodeDecodeError):
                    # Jika masih gagal, log warning dan return 0
                    logger.warning(f"Failed to parse numeric value: {data}")
                    return 0
        except Exception as e:
            logger.error(f"Error in parseN: {str(e)}")
            return 0