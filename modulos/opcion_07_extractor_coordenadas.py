import os
import piexif
import pyperclip
from pyproj import Transformer

def ejecutar():
    def _rat_to_float(r):
        if isinstance(r, tuple):
            num, den = r
            return float(num) / float(den) if den else float(num)
        return float(r)

    def _dms_to_dd(dms, ref):
        deg = _rat_to_float(dms[0])
        minute = _rat_to_float(dms[1]) if len(dms) > 1 else 0.0
        sec = _rat_to_float(dms[2]) if len(dms) > 2 else 0.0
        dd = deg + minute / 60.0 + sec / 3600.0
        if ref in (b'S', b'W', 'S', 'W'):
            dd *= -1.0
        return dd

    def get_latlon_from_exif(img_path: str):
        exif_dict = piexif.load(img_path)
        gps = exif_dict.get("GPS", {})
        if not gps:
            raise ValueError("La imagen no contiene bloque GPS en el EXIF.")
        lat = gps.get(piexif.GPSIFD.GPSLatitude)
        lat_ref = gps.get(piexif.GPSIFD.GPSLatitudeRef)
        lon = gps.get(piexif.GPSIFD.GPSLongitude)
        lon_ref = gps.get(piexif.GPSIFD.GPSLongitudeRef)
        if not all([lat, lat_ref, lon, lon_ref]):
            raise ValueError("Faltan GPSLatitude/GPSLongitude en el EXIF.")
        return _dms_to_dd(lat, lat_ref), _dms_to_dd(lon, lon_ref)

    def wgs84_to_utm17s(lat: float, lon: float):
        transformer = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
        e, n = transformer.transform(lon, lat)
        return e, n

    def copiar_coordenadas(img_path: str, formato: str) -> str:
        if not os.path.isfile(img_path):
            raise FileNotFoundError(f"No se encontró el archivo: {img_path}")
        lat, lon = get_latlon_from_exif(img_path)
        formato_norm = (formato or "").strip().upper()
        if formato_norm == "WGS84":
            text = f"{lat:.6f} {lon:.6f}"
        elif formato_norm == "UTM17S":
            e, n = wgs84_to_utm17s(lat, lon)
            text = f"{e:.0f} {n:.0f}"
        else:
            raise ValueError("Formato inválido. Use WGS84 o UTM17S.")
        pyperclip.copy(text)
        return text

    op2 = "0"
    while op2 == "0":
        ruta_foto = input("Ruta de la foto: (o 's' para salir)")
        if ruta_foto.lower() == "s":
            op2 = 1
        else:
            ruta_foto = ruta_foto.strip().strip("'").strip('"')
            try:
                print("¿En qué sistema de coordenadas quieres copiar el resultado?")
                print("1. WGS84 (Latitud Longitud)")
                print("2. UTM 17S (X Y)")
                while True:
                    op_formato = input("> ").strip()
                    if op_formato == "1":
                        formato = "WGS84"
                        break
                    if op_formato == "2":
                        formato = "UTM17S"
                        break
                    print("Opción no válida. Elige 1 (WGS84) o 2 (UTM 17S).")

                resultado = copiar_coordenadas(ruta_foto, formato)
                print("Copiado al portapapeles ✅")
                print(resultado)
            except Exception as e:
                print("⚠️", e)