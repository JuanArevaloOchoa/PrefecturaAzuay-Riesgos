import os
import time
from datetime import datetime
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
from pyproj import Transformer
import piexif
import pyperclip
from openpyxl import load_workbook
from modulos.comun import _ORIG_INPUT, RUTA_BASE_CARPETAS, RUTA_EXCEL_TRAMITES, RUTA_GPKG_PARROQUIAS

def ejecutar():
    def obtener_datos_geograficos(x, y, url_gpkg):
        try:
            transformer = Transformer.from_crs("epsg:32717", "epsg:4326", always_xy=True)
            lon, lat = transformer.transform(float(x), float(y))
            url_maps = f"https://www.google.com/maps?q={lat},{lon}"

            gdf = gpd.read_file(url_gpkg)
            punto = Point(float(x), float(y))
            resultado = gdf[gdf.contains(punto)]

            if not resultado.empty:
                parroquia = resultado.iloc[0].get('PARROQUIA', 'No encontrada')
                canton = resultado.iloc[0].get('CANTON', 'No encontrado')
                return parroquia, canton, url_maps, lon, lat
            else:
                return "Fuera de límites", "Fuera de límites", url_maps, lon, lat
        except Exception as e:
            return f"Error: {e}", "Error", "", "", ""

    def crear_carpeta_tramite(no_registro, sector_actual, parroquia, canton):
        try:
            sector_final = sector_actual
            if sector_actual and parroquia and canton:
                nombre_carpeta = f"{no_registro}. {sector_actual}, {parroquia}, {canton}"
            else:
                print(f"\n📂 No hay datos de ubicación completos.")
                nombre_personalizado = input(f"Ingrese el nombre para la carpeta (El No. {no_registro} se añadirá auto): > ")
                nombre_carpeta = f"{no_registro}. {nombre_personalizado}"
                sector_final = nombre_personalizado

            for char in '<>:"/\\|?*':
                nombre_carpeta = nombre_carpeta.replace(char, '')

            ruta_final = os.path.abspath(os.path.join(RUTA_BASE_CARPETAS, nombre_carpeta))

            if not os.path.exists(ruta_final):
                os.makedirs(ruta_final)
                print(f"📁 Carpeta creada: {nombre_carpeta}")
            else:
                print(f"ℹ️ La carpeta ya existe: {nombre_carpeta}")

            ruta_documentos = ""
            for subcarpeta in ["Fotos", "Documentos"]:
                ruta_sub = os.path.join(ruta_final, subcarpeta)
                if not os.path.exists(ruta_sub):
                    os.makedirs(ruta_sub)
                if subcarpeta == "Documentos":
                    ruta_documentos = ruta_sub

            pyperclip.copy(ruta_documentos)
            print(f"   ∟ ✅ Subcarpetas listas.")
            print(f"   ∟ 📋 Ruta de 'Documentos' copiada al portapapeles.")

            return ruta_final, sector_final
        except Exception as e:
            print(f"❌ Error al crear estructura de carpetas: {e}")
            return "Error", sector_actual

    def limpiar_formato_fecha(texto_fecha):
        if not texto_fecha or str(texto_fecha).lower() == "nan":
            return ""
        try:
            fecha_solo_texto = str(texto_fecha).strip()[:10]
            objeto_fecha = datetime.strptime(fecha_solo_texto, "%Y-%m-%d")
            return objeto_fecha.strftime("%d/%m/%Y")
        except Exception:
            return texto_fecha

    def obtener_lat_lon_foto(img_path):
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

        try:
            exif_dict = piexif.load(img_path)
            gps = exif_dict.get("GPS", {})
            if not gps:
                return None, None
            lat = gps.get(piexif.GPSIFD.GPSLatitude)
            lat_ref = gps.get(piexif.GPSIFD.GPSLatitudeRef)
            lon = gps.get(piexif.GPSIFD.GPSLongitude)
            lon_ref = gps.get(piexif.GPSIFD.GPSLongitudeRef)
            if not all([lat, lat_ref, lon, lon_ref]):
                return None, None
            return _dms_to_dd(lat, lat_ref), _dms_to_dd(lon, lon_ref)
        except Exception as e:
            print(f"⚠️ Error procesando imagen: {e}")
            return None, None

    ruta_excel = RUTA_EXCEL_TRAMITES
    url_gpkg = RUTA_GPKG_PARROQUIAS

    columnas_excel = [
        "NO", "DE", "ASUNTO", "FECHA DOCUMENTO", "NUMERO DOCUMENTO",
        "NO REFERENCIA", "TIPO DOCUMENTO", "NO TRÁMITE",
        "USUARIO ANTERIOR", "FECHA ÚLTIMA ACCIÓN",
        "COORDENADA_X", "COORDENADA_Y", "LATITUD", "LONGITUD", "SECTOR", "PARROQUIA", "CANTÓN",
        "GOOGLE_MAPS", "ESTADO", "DIRECTORIO"
    ]

    print("\n" + "=" * 50)
    print("SISTEMA DE REGISTRO DE TRÁMITES (FECHA CORREGIDA)")
    print("=" * 50)
    print("Pega la línea de texto (DE -> FECHA ACCIÓN) o 'salir':")
    texto_raw = input("> ")

    if texto_raw.lower() == 'salir':
        return

    if not texto_raw.strip():
        print("\n--- INGRESO MANUAL DE DATOS ---")
        de_manual = input("DE (Remitente): > ")
        asunto_manual = input("ASUNTO: > ")
        fecha_solicitud = input("FECHA DE SOLICITUD (DD/MM/AAAA): > ")
        partes = [de_manual, asunto_manual, fecha_solicitud]
    else:
        partes = [p.strip() for p in texto_raw.split('\t')]

    print("\n¿Este trámite tiene ubicación física?")
    print("1. Sí")
    print("2. No")
    while True:
        opcion_ubicacion = _ORIG_INPUT("> ").strip()
        if opcion_ubicacion == "1":
            tiene_ubicacion = True
            break
        elif opcion_ubicacion == "2":
            tiene_ubicacion = False
            break
        else:
            print("Opción no válida. Elige 1 (Sí) o 2 (No).")

    coord_x, coord_y, sector_input, parroquia, canton, url_maps, lat, lon = "", "", "", "", "", "", "", ""

    if tiene_ubicacion:
        print("\n¿Cómo deseas ingresar la ubicación?")
        print("1. Coordenada UTM")
        print("2. Coordenada WGS84")
        print("3. Extraer coordenada de una fotografía (directorio)")

        opcion_coord = input("\nSelecciona una opción (1, 2 o 3): > ").strip()
        transformer_to_utm = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)

        if opcion_coord == '1':
            coord_x = input("Ingrese Coordenada UTM X: > ")
            coord_y = input("Ingrese Coordenada UTM Y: > ")
        elif opcion_coord == '2':
            latitud_input = float(input("Ingrese Latitud (WGS84, ej. -2.90): > "))
            longitud_input = float(input("Ingrese Longitud (WGS84, ej. -79.0): > "))
            coord_x_num, coord_y_num = transformer_to_utm.transform(longitud_input, latitud_input)
            coord_x, coord_y = str(coord_x_num), str(coord_y_num)
            print(f"✅ Convertido a UTM: X={coord_x_num:.2f}, Y={coord_y_num:.2f}")
        elif opcion_coord == '3':
            ruta_foto = input("Ingrese el directorio/ruta de la fotografía: > ").strip().strip("'").strip('"')
            latitud_img, longitud_img = obtener_lat_lon_foto(ruta_foto)

            if latitud_img is not None and longitud_img is not None:
                coord_x_num, coord_y_num = transformer_to_utm.transform(longitud_img, latitud_img)
                coord_x, coord_y = str(coord_x_num), str(coord_y_num)
                print(f"✅ Coordenadas extraídas de la foto:\nLat={latitud_img:.6f}, Lon={longitud_img:.6f} -> UTM X={coord_x_num:.2f}, Y={coord_y_num:.2f}")
            else:
                print("❌ No se pudo obtener la ubicación de la fotografía. Las coordenadas quedarán vacías.")
        else:
            print("❌ Opción no válida. Las coordenadas quedarán vacías.")

        if coord_x and coord_y:
            sector_input = input("Ingrese el Sector: > ")
            print("Consultando base de datos geográfica...")
            parroquia, canton, url_maps, lon, lat = obtener_datos_geograficos(coord_x, coord_y, url_gpkg)
            print(f"Detectado: {parroquia} / {canton}")

    if os.path.exists(ruta_excel):
        try:
            df_temp = pd.read_excel(ruta_excel)
            ultimo_no = pd.to_numeric(df_temp['NO'], errors='coerce').max()
            nuevo_no = 1 if pd.isna(ultimo_no) else int(ultimo_no) + 1

            ruta_generada, sector_definitivo = crear_carpeta_tramite(nuevo_no, sector_input, parroquia, canton)

            datos_mapeados = {
                "NO": nuevo_no,
                "DE": partes[0] if len(partes) > 0 else "",
                "ASUNTO": partes[1] if len(partes) > 1 else "",
                "FECHA DOCUMENTO": limpiar_formato_fecha(partes[2]) if len(partes) > 2 else "",
                "NUMERO DOCUMENTO": partes[3] if len(partes) > 3 else "",
                "NO REFERENCIA": partes[4] if len(partes) > 4 else "",
                "TIPO DOCUMENTO": partes[5] if len(partes) > 5 else "",
                "NO TRÁMITE": partes[6] if len(partes) > 6 else "",
                "USUARIO ANTERIOR": partes[7] if len(partes) > 7 else "",
                "FECHA ÚLTIMA ACCIÓN": limpiar_formato_fecha(partes[8]) if len(partes) > 8 else "",
                "COORDENADA_X": coord_x,
                "COORDENADA_Y": coord_y,
                "LATITUD": lat,
                "LONGITUD": lon,
                "SECTOR": sector_definitivo,
                "PARROQUIA": parroquia,
                "CANTÓN": canton,
                "GOOGLE_MAPS": url_maps,
                "ESTADO": "En Trámite",
                "DIRECTORIO": ruta_generada
            }

            fila_valores = [datos_mapeados.get(col, "") for col in columnas_excel]

            guardado = False
            while not guardado:
                try:
                    book = load_workbook(ruta_excel)
                    sheet = book.active
                    sheet.append(fila_valores)
                    book.save(ruta_excel)
                    guardado = True
                    print(f"\n✅ Registro {nuevo_no} guardado con éxito.")
                except PermissionError:
                    print("\n⚠️ CIERRA EL EXCEL para guardar...")
                    time.sleep(5)
        except Exception as e:
            print(f"\n❌ ERROR: {e}")
    else:
        print("\n❌ No se encontró el Excel en la ruta especificada.")