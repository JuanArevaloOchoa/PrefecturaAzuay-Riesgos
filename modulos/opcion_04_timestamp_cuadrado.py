import os
from datetime import datetime
from pathlib import Path
import piexif
from PIL import Image, ImageDraw, ImageFont, ExifTags
from pyproj import Transformer
import geopandas as gpd
from shapely.geometry import Point
from modulos.comun import RUTA_GPKG_PARROQUIAS, RUTA_LOGO_PREFECTURA

def ejecutar():
    Shape = RUTA_GPKG_PARROQUIAS
    parroquias_gdf = gpd.read_file(Shape)
    parroquias_gdf.set_crs(epsg=32717, inplace=True)

    def get_exif_data(img_path):
        img = Image.open(img_path)
        exif_data = piexif.load(img.info.get('exif', b''))
        gps_data = exif_data.get('GPS', {})
        date_time_raw = exif_data.get('Exif', {}).get(piexif.ExifIFD.DateTimeOriginal)

        if not gps_data or 2 not in gps_data or 4 not in gps_data:
            print(f"❌ Sin datos GPS válidos en {img_path}")
            return None, None, None

        def to_deg(value):
            d, m, s = value
            deg = d[0] / d[1]
            minutes = m[0] / m[1]
            seconds = s[0] / s[1]
            return deg + (minutes / 60) + (seconds / 3600)

        try:
            lat = to_deg(gps_data[2])
            lon = to_deg(gps_data[4])
            if gps_data.get(1) == b'S': lat = -lat
            if gps_data.get(3) == b'W': lon = -lon
        except Exception as e:
            print(f"⚠️ Error convirtiendo coordenadas en {img_path}: {e}")
            return None, None, None

        print(f"📍 Extraído de {os.path.basename(img_path)}: lat={lat:.6f}, lon={lon:.6f}")

        if date_time_raw:
            try:
                dt_obj = datetime.strptime(date_time_raw.decode(), "%Y:%m:%d %H:%M:%S")
                fecha = dt_obj.strftime("%Y-%m-%d %H:%M")
            except Exception:
                fecha = date_time_raw.decode()
        else:
            fecha = "Fecha desconocida"

        if 17 in gps_data:
            val = gps_data[17]
            azimut = f"N{round(val[0]/val[1])}°"
        else:
            azimut = "No disponible"

        return (lat, lon), fecha, azimut

    def convert_wgs84_to_utm17s(lat, lon):
        transformer = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
        easting, northing = transformer.transform(lon, lat)
        return round(easting), round(northing)

    def corregir_orientacion(img):
        try:
            exif = img._getexif()
            if exif:
                for tag, value in exif.items():
                    tag_name = ExifTags.TAGS.get(tag, tag)
                    if tag_name == 'Orientation':
                        if value == 3: img = img.rotate(180, expand=True)
                        elif value == 6: img = img.rotate(270, expand=True)
                        elif value == 8: img = img.rotate(90, expand=True)
                        break
        except Exception as e:
            print("No se pudo corregir orientación:", e)
        return img

    def recortar_a_cuadrado_centrado(img, tamaño_px=945):
        img_ratio = img.width / img.height
        if img_ratio > 1:
            new_height = tamaño_px
            new_width = int(img_ratio * tamaño_px)
        else:
            new_width = tamaño_px
            new_height = int(tamaño_px / img_ratio)
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        izquierda = (new_width - tamaño_px) // 2
        superior = (new_height - tamaño_px) // 2
        derecha = izquierda + tamaño_px
        inferior = superior + tamaño_px
        return img.crop((izquierda, superior, derecha, inferior))

    def obtener_parroquia_canton(utm_e, utm_n):
        punto = Point(utm_e, utm_n).buffer(3)
        fila = parroquias_gdf[parroquias_gdf.geometry.intersects(punto)]
        if not fila.empty:
            parroquia = fila.iloc[0]['PARROQUIA'].title()
            canton = fila.iloc[0]['CANTON'].title()
        else:
            print(f"⚠️ Coordenada fuera del Azuay: UTM {utm_e}, {utm_n}")
            parroquia = "Desconocida"
            canton = "Desconocido"
        return parroquia, canton

    def stamp_photo(img_path, output_folder, autor="Ing. Juan Arévalo", empresa="GAD Provincial del Azuay", logo_path=None):
        resultado = get_exif_data(img_path)
        if resultado[0] is None:
            print(f"No GPS en: {img_path}")
            return
        (lat, lon), fecha, azimut = resultado
        utm_e, utm_n = convert_wgs84_to_utm17s(lat, lon)
        parroquia, canton = obtener_parroquia_canton(utm_e, utm_n)

        img = Image.open(img_path)
        img = corregir_orientacion(img)
        img = recortar_a_cuadrado_centrado(img)
        img = img.convert("RGBA")
        draw = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype("arial.ttf", size=18)
        except Exception:
            font = ImageFont.load_default()

        texto = (
            f"Autor: {autor}\n"
            f"Fecha: {fecha}\n"
            f"Coord. WGS84: {lat:.6f}, {lon:.6f}\n"
            f"Coord. UTM17S: {utm_e}, {utm_n}\n"
            f"Azimut: {azimut}\n"
            f"Parroquia: {parroquia}\n"
            f"Cantón: {canton}\n"
            f"{empresa}"
        )

        text_lines = texto.split("\n")
        max_width = max(draw.textlength(line, font=font) for line in text_lines)
        line_height = font.getbbox("A")[3] - font.getbbox("A")[1] + 4
        box_height = line_height * len(text_lines) + 10
        margin = 10
        x, y = margin, margin

        text_box = Image.new("RGBA", img.size)
        draw_box = ImageDraw.Draw(text_box)
        draw_box.rectangle([x - 5, y - 5, x + max_width + 10, y + box_height], fill=(0, 0, 0, 180))
        for i, line in enumerate(text_lines):
            draw_box.text((x, y + i * line_height), line, font=font, fill="white")
        img = Image.alpha_composite(img, text_box)

        if logo_path and os.path.exists(logo_path):
            logo = Image.open(logo_path).convert("RGBA")
            logo.thumbnail((250, 250), Image.Resampling.LANCZOS)
            img_w, _ = img.size
            logo_w, _ = logo.size
            position = (img_w - logo_w - margin, margin)
            img.paste(logo, position, mask=logo)

        img = img.convert("RGB")
        nombre_salida = os.path.basename(img_path)
        ruta_salida = os.path.join(output_folder, nombre_salida)
        img.save(ruta_salida, dpi=(300, 300))
        print(f"✅ Guardado: {ruta_salida}")

    def procesar_directorio(directorio, logo_path=None):
        output_folder = Path(directorio) / "Recortes"
        output_folder.mkdir(parents=True, exist_ok=True)
        for archivo in os.listdir(directorio):
            if archivo.lower().endswith(('.jpg', '.jpeg')):
                ruta = os.path.join(directorio, archivo)
                stamp_photo(ruta, str(output_folder), logo_path=logo_path)

    directorio = input("Ingrese el directorio: ")
    procesar_directorio(directorio, logo_path=RUTA_LOGO_PREFECTURA)