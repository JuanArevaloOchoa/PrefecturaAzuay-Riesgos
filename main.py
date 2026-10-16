import os
import sys
import time
import re
import shutil
import warnings
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
import builtins
import xml.etree.ElementTree as ET

import cv2
import openpyxl
import pandas as pd
import numpy as np
import requests
from requests.auth import HTTPBasicAuth
import urllib3
import piexif
from PIL import Image, ImageDraw, ImageFont, ExifTags
import pyperclip
from shapely.geometry import Point
from pyproj import Transformer
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from arcgis.gis import GIS
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from moviepy import ImageSequenceClip
import rasterio
import rasterio.plot
from rasterio.warp import reproject, Resampling, calculate_default_transform
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize
import matplotlib.patheffects as path_effects
import matplotlib.image as mpimg
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
import geopandas as gpd
import contextily as cx
from dotenv import load_dotenv

# Cargar variables secretas desde .env
load_dotenv()

# Desactivar avisos SSL del servidor
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# -----------------------------
# Cache en memoria: BD ArcGIS
# -----------------------------
ARCGIS_ITEM_ID = "b255d12ed7424df8914bc7bbd0e05a1d"
ARCGIS_PORTAL_URL = "https://azuay.maps.arcgis.com/"
ARCGIS_USERNAME = os.getenv("ARCGIS_USER", "jarevalo_azuay")
ARCGIS_PASSWORD = os.getenv("ARCGIS_PASS", "Azuay.1234567")

_ARCGIS_CACHE = {
    "df": None,
    "loaded_at": None,
    "error": None,
}

REPORTES_DIR = Path(r"D:\Desktop\Reportes")


def asegurar_reportes_dir() -> Path:
    REPORTES_DIR.mkdir(parents=True, exist_ok=True)
    return REPORTES_DIR


def ruta_en_reportes(nombre_archivo: str, subcarpeta: Optional[str] = None) -> Path:
    base = asegurar_reportes_dir()
    if subcarpeta:
        base = base / subcarpeta
        base.mkdir(parents=True, exist_ok=True)
    return base / nombre_archivo


def cargar_bd_arcgis(force: bool = False) -> Optional[pd.DataFrame]:
    global _ARCGIS_CACHE

    if _ARCGIS_CACHE["df"] is not None and not force:
        return _ARCGIS_CACHE["df"]

    print("\nConectando a ArcGIS y cargando BD en memoria...")
    try:
        try:
            gis = GIS(ARCGIS_PORTAL_URL, ARCGIS_USERNAME, ARCGIS_PASSWORD)
        except Exception:
            gis = GIS()

        item = gis.content.get(ARCGIS_ITEM_ID)
        if item is None:
            raise ValueError("No se pudo encontrar el item en ArcGIS (verifica el Item ID / permisos).")

        layer = item.layers[0]
        sdf = layer.query(where="1=1", out_fields="*").sdf
        df = pd.DataFrame(sdf)

        if "fecha_de_visita" in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df["fecha_de_visita"]):
                pass
            elif pd.api.types.is_numeric_dtype(df["fecha_de_visita"]):
                df["fecha_de_visita"] = pd.to_datetime(df["fecha_de_visita"], unit="ms", errors="coerce")
            else:
                df["fecha_de_visita"] = pd.to_datetime(df["fecha_de_visita"], errors="coerce")

        _ARCGIS_CACHE["df"] = df
        _ARCGIS_CACHE["loaded_at"] = datetime.now()
        _ARCGIS_CACHE["error"] = None

        print(f"✅ BD ArcGIS cargada: {len(df)} registros.")
        return df
    except Exception as e:
        _ARCGIS_CACHE["df"] = None
        _ARCGIS_CACHE["loaded_at"] = None
        _ARCGIS_CACHE["error"] = str(e)
        print(f"⚠️ No se pudo cargar la BD ArcGIS en memoria: {e}")
        return None


def recargar_bd_arcgis():
    cargar_bd_arcgis(force=True)


# -----------------------------
# Salida global con "s" en inputs
# -----------------------------
_ORIG_INPUT = builtins.input


class SalirInput(BaseException):
    def __init__(self, prompt: str = ""):
        self.prompt = prompt or ""
        super().__init__(self.prompt)


def _agregar_s_para_salir(prompt: str) -> str:
    p = prompt or ""
    pl = p.lower()
    if "s para salir" in pl or "o 's'" in pl or 'o "s"' in pl:
        return p
    return f"{p} (s para salir): "


def _input_con_salida(prompt: str = "") -> str:
    valor = _ORIG_INPUT(_agregar_s_para_salir(prompt))
    if isinstance(valor, str) and valor.strip().lower() == "s":
        raise SalirInput(prompt)
    return valor


builtins.input = _input_con_salida


def mostrar_menu():
    print("\n--- SISTEMA DE TRÁMITES ---")
    print("1. Recargar BD Arcgis")
    print("2. Ingresar nuevo trámite")
    print("3. Crear documentos con survey")
    print("4. Timestamp cuadrado")
    print("5. Timestamp rectangular (FDS)")
    print("6. Conversor de Coordenadas")
    print("7. Extractor de Coordenadas de Fotos")
    print("8. Reporte de horas")
    print("9. Reporte Emergencia")
    print("10. Descargar BD Arcgis")
    print("11. Generar Ayuda Memoria")
    print("12. Tablas Resúmenes de puntos")
    print("13. Vídeo pronóstico INHAMI")
    print("14. Timestamp de vídeos")
    print("0. Salir del sistema")


def ejecutar_menu():
    while True:
        try:
            mostrar_menu()
            seleccion = input("Seleccione su opción: ").strip()
        except SalirInput:
            print("Saliendo del sistema... ¡Hasta luego!")
            break

        try:
            if seleccion == "1":
                recargar_bd_arcgis()
            elif seleccion == "2":
                ingresar_nuevo_tramite()
            elif seleccion == "3":
                crear_documento_survey()
            elif seleccion == "4":
                timestamp_cuadrado()
            elif seleccion == "5":
                print("Fuera de servicio")
            elif seleccion == "6":
                conversor_coordenadas()
            elif seleccion == "7":
                extractor_coordenadas_fotos()
            elif seleccion == "8":
                reporte_horas()
            elif seleccion == "9":
                reporte_emergencia()
            elif seleccion == "10":
                descargar_bd_arcgis()
            elif seleccion == "11":
                generar_reporte_desde_arcgis()
            elif seleccion == "12":
                generar_tablas_resumen_puntos()
            elif seleccion == "13":
                procesar_video_wrf()
            elif seleccion == "14":
                timestamp_videos()
            elif seleccion == "0":
                print("Saliendo del sistema... ¡Hasta luego!")
                break
            else:
                print("\n[!] Opción no válida. Intente de nuevo.")
        except SalirInput:
            print("\nOperación cancelada. Volviendo al menú...")
            continue


def ingresar_nuevo_tramite():
    RUTA_BASE_CARPETAS = r"D:\OneDrive - GAD Provincial del Azuay"

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

    ruta_excel = r"D:\OneDrive - GAD Provincial del Azuay\RegistroTrámitesJuan.xlsx"
    url_gpkg = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\AzuayXParroquias.gpkg"

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


def crear_documento_survey():
    try:
        sedf_resultados = cargar_bd_arcgis()
        if sedf_resultados is None or sedf_resultados.empty:
            print("❌ No se pudo cargar la BD de ArcGIS (o está vacía).")
            return
        if "fecha_de_visita" in sedf_resultados.columns and sedf_resultados["fecha_de_visita"].dtype == "object":
            sedf_resultados["fecha_de_visita"] = pd.to_datetime(sedf_resultados["fecha_de_visita"], unit="ms", errors="coerce")
    except Exception as e:
        print(f"❌ Error GIS: {e}")
        return

    sedf_recientes = sedf_resultados.sort_values(by='fecha_de_visita', ascending=False).head(30)
    print("\n--- 10 ENTIDADES MÁS RECIENTES (GIS) ---")
    for _, fila in sedf_recientes.iterrows():
        print(f"{fila['objectid']}. {fila['nombre_del_punto']}")

    id_elegido_gis = int(input("\nIngrese el ObjectID del punto GIS: "))
    punto_sel = sedf_recientes[sedf_recientes['objectid'] == id_elegido_gis].iloc[0]

    for col, val in punto_sel.items():
        nombre_var = col.replace(' ', '_').lower()
        globals()[nombre_var] = val.strip() if isinstance(val, str) else val

    ruta_excel = r"D:\OneDrive - GAD Provincial del Azuay\RegistroTrámitesJuan.xlsx"
    COL_OBJECTID = "OBJECTID"
    COL_INFORME = "NO INFORME TÉCNICO"
    COL_NOMBRE = "NOMBRE"
    COL_FECHA = "FECHA VT"
    COL_DESC = "DESCRIPCIÓN"
    COL_TECNICO = "COMENTARIOS TÉCNICOS"
    COL_VIAL = "INVENTARIO"

    rutafinal1 = None

    try:
        df_temp = pd.read_excel(ruta_excel)
        print("\n--- ÚLTIMOS 10 REGISTROS DEL EXCEL ---")
        df_ordenado_excel = df_temp.sort_values(by='NO', ascending=False).head(20)
        for _, fila in df_ordenado_excel.iterrows():
            no_mostrar = str(fila['NO']).replace('.0', '')
            print(f"{no_mostrar}. {fila['SECTOR / NOMBRE']}")
        print("-" * 30)

        busqueda_no = float(input("Escribe el número (NO) del registro de Excel para destino: "))
        indices = df_temp.index[df_temp['NO'] == busqueda_no].tolist()
        if not indices:
            print("❌ No se encontró ese número en el Excel.")
            return

        fila_excel = indices[0] + 2
        libro = openpyxl.load_workbook(ruta_excel)
        hoja = libro.active
        cabecera = [celda.value for celda in hoja[1]]

        def obtener_col(nombre_col):
            try:
                return cabecera.index(nombre_col) + 1
            except ValueError:
                return None

        def limpiar_para_excel(val):
            if val is None or pd.isna(val) or str(val).strip().lower() in ["<na>", "nan", "none"]:
                return ""
            if hasattr(val, 'strftime'):
                return val.strftime('%d/%m/%Y')
            return val

        mapa_datos = {
            COL_OBJECTID: limpiar_para_excel(id_elegido_gis),
            COL_INFORME: limpiar_para_excel(globals().get('no_informe', '')),
            COL_NOMBRE: limpiar_para_excel(globals().get('nombre_del_punto', '')),
            COL_FECHA: limpiar_para_excel(globals().get('fecha_de_visita', '')),
            COL_DESC: limpiar_para_excel(globals().get('descripci_n_del_evento', '')),
            COL_TECNICO: limpiar_para_excel(globals().get('comentarios_t_cnicos', '')),
            COL_VIAL: limpiar_para_excel(globals().get('inventario_vial', ''))
        }

        for nombre_col, valor in mapa_datos.items():
            col_num = obtener_col(nombre_col)
            if col_num:
                hoja.cell(row=fila_excel, column=col_num).value = valor

        col_dir = obtener_col("DIRECTORIO")
        if col_dir:
            rutafinal1 = hoja.cell(row=fila_excel, column=col_dir).value
        else:
            print("❌ No se encontró la columna 'DIRECTORIO'.")
            return

        if not rutafinal1:
            print("❌ La celda 'DIRECTORIO' en el registro seleccionado está vacía.")
            return

        libro.save(ruta_excel)
        print("✅ Excel actualizado y formato preservado.")
    except Exception as e:
        print(f"❌ Error en Excel: {e}")
        return

    transformer = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
    CoordenadaX, CoordenadaY = transformer.transform(globals().get('longitud', 0), globals().get('latitud', 0))

    meses = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
             7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
    hoy = datetime.now()
    fecha_formateada = f"{hoy.day} DE {meses[hoy.month]} DEL {hoy.year}"

    f_visita = globals().get('fecha_de_visita')
    if hasattr(f_visita, 'strftime'):
        solo_fecha_VT = f"{f_visita.day} de {f_visita.strftime('%B')} del {f_visita.year}"
    else:
        solo_fecha_VT = str(f_visita) if pd.notna(f_visita) else ""

    TecnicoGPA = {"Juan_Arévalo": "Ing. Juan Arévalo Ochoa, Responsable de Riesgos",
                  "Xavier_Sánchez": "Ing. Xavier Sánchez, Técnico de Riesgos"}.get(globals().get('field_50'), "Null")

    ruta_plantilla = r"D:\OneDrive - GAD Provincial del Azuay\1000. Nuevo informe y GIS\PLANTILLA INFORME.docx"
    informe_nombre = globals().get('no_informe', 'INFORME_SIN_NUMERO')
    ruta_salida_word = os.path.join(str(rutafinal1), f"{informe_nombre}.docx")

    temp_series = punto_sel.replace(to_replace='<NA>', value=np.nan)
    filtered_series = temp_series.dropna()
    resultado_final_str = ""
    for key, value in filtered_series[6:-1].items():
        resultado_final_str += f"{key}: {value}\n"

    sustituciones = {
        "#AUTOR#": "Juan Arévalo Ochoa",
        "#FECHAINF#": fecha_formateada.upper(),
        "#TIPOEVENTO#": globals().get('tipo_de_evento', ''),
        "#SUBTIPOEVENTO#": globals().get('subtipo_de_evento', ''),
        "#FECHAVT#": solo_fecha_VT,
        "#COORDX#": int(CoordenadaX),
        "#COORDY#": int(CoordenadaY),
        "#SECTOR#": str(globals().get('sector', '')).title(),
        "#PARROQUIA#": globals().get('parroquia', ''),
        "#CANTÓN#": globals().get('cant_n', ''),
        "#ESTADO#": "",
        "#AREAAFECTADA#": "",
        "#TÉCNICOSGADPA#": TecnicoGPA,
        "#PERSONALGADCANTONAL#": globals().get('personal_gad_cantonal', ''),
        "#PERSONALGADPARROQUIAL#": globals().get('personal_gad_parroquial', ''),
        "#NOTRAMITE#": globals().get('no_tr_mite', ''),
        "#NOINFORME#": informe_nombre,
        "#NOMBREPUNTO#": globals().get('nombre_del_punto', ''),
        "#DESCRIPCION#": globals().get('descripci_n_del_evento', ''),
        "#COMENTARIOSTECNICOS#": globals().get('comentarios_t_cnicos', ''),
        "#FECHAEVENTO#": globals().get('fecha_y_hora_de_evento', ''),
        "#DATOSTOTALES#": resultado_final_str,
    }

    def reemplazar(doc_obj, susts):
        for p in doc_obj.paragraphs:
            for run in p.runs:
                for c, v in susts.items():
                    if c in run.text:
                        run.text = run.text.replace(c, str(v) if v is not None else "")
        for t in doc_obj.tables:
            for fila in t.rows:
                for celda in fila.cells:
                    for p in celda.paragraphs:
                        for run in p.runs:
                            for c, v in susts.items():
                                if c in run.text:
                                    run.text = run.text.replace(c, str(v) if v is not None else "")

    if pd.notna(globals().get('no_informe')):
        doc = Document(ruta_plantilla)
        reemplazar(doc, sustituciones)
        doc.save(ruta_salida_word)
        print("📄 Informe Word guardado en la ruta elegida.")

    DIR_ORIGEN = r"D:\OneDrive - GAD Provincial del Azuay\1000. Nuevo informe y GIS"
    for item in Path(DIR_ORIGEN).iterdir():
        if item.name == "PLANTILLA INFORME.docx":
            continue
        dest = Path(rutafinal1) / item.name
        if item.is_file():
            shutil.copy2(item, dest)
        elif item.is_dir():
            shutil.copytree(item, dest, dirs_exist_ok=True)

    print("\n🚀 PROCESO FINALIZADO CON ÉXITO.")


def timestamp_cuadrado():
    Shape = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\AzuayXParroquias.gpkg"
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
    logo_path = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\logo prefectura.png"
    procesar_directorio(directorio, logo_path=logo_path)


def conversor_coordenadas():
    print("--- CONVERSOR DE COORDENADAS (Azuay - Zona 17S) ---")
    print("1. WGS84 (Lat/Lon) a UTM 17S (E/N)")
    print("2. UTM 17S (E/N) a WGS84 (Lat/Lon)")

    opcion = input("Seleccione una opción (1 o 2): ")

    if opcion == "1":
        transformer = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
        try:
            lat = float(input("Ingrese Latitud (ej. -2.9001): "))
            lon = float(input("Ingrese Longitud (ej. -79.0001): "))
            easting, northing = transformer.transform(lon, lat)
            print("-" * 30)
            print(f"📍 Resultado UTM 17S:\nEste (X): {easting:.2f}\nNorte (Y): {northing:.2f}")
            print("-" * 30)
        except ValueError:
            print("❌ Error: Por favor ingrese números válidos.")
    elif opcion == "2":
        transformer = Transformer.from_crs("EPSG:32717", "EPSG:4326", always_xy=True)
        try:
            easting = float(input("Ingrese Coordenada Este (X): "))
            northing = float(input("Ingrese Coordenada Norte (Y): "))
            lon, lat = transformer.transform(easting, northing)
            print("-" * 30)
            print(f"📍 Resultado WGS84:\nLatitud: {lat:.8f}\nLongitud: {lon:.8f}")
            print("-" * 30)
        except ValueError:
            print("❌ Error: Por favor ingrese números válidos.")
    else:
        print("Opción no válida.")


def extractor_coordenadas_fotos():
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


def reporte_horas():
    print("--- PROCESADOR DE ASISTENCIA ---")
    ruta_archivo = input("Introduce la ruta completa del archivo (ej: C:/Carpeta/asistencia.xlsx): ").strip()
    ruta_archivo = ruta_archivo.replace('"', '')

    if not os.path.exists(ruta_archivo):
        print(f"Error: No se encontró el archivo en {ruta_archivo}")
        return

    try:
        if ruta_archivo.lower().endswith('.csv'):
            try:
                df_raw = pd.read_csv(ruta_archivo, header=None, encoding='utf-8')
            except Exception:
                df_raw = pd.read_csv(ruta_archivo, header=None, encoding='latin1')
        else:
            df_raw = pd.read_excel(ruta_archivo, header=None)

        idx_header = df_raw[df_raw.astype(str).apply(lambda x: x.str.contains('Fecha', case=False)).any(axis=1)].index[0]
        df = df_raw.iloc[idx_header:].copy()
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

        df = df.dropna(subset=['Fecha', 'Hora'])
        df['Fecha_dt'] = pd.to_datetime(df['Fecha'], dayfirst=True)
        df = df.drop_duplicates(subset=['Fecha_dt', 'Hora'])
        df = df.sort_values(by=['Fecha_dt', 'Hora'], ascending=[True, True])
        df['Num'] = df.groupby('Fecha_dt').cumcount() + 1
        df = df[df['Num'] <= 4]

        reporte = df.pivot(index='Fecha_dt', columns='Num', values='Hora')
        columnas_finales = {1: 'ENTRADA', 2: 'SALIDA ALMUERZO', 3: 'ENTRADA ALMUERZO', 4: 'SALIDA'}
        reporte = reporte.rename(columns=columnas_finales)
        reporte = reporte.sort_index(ascending=False)

        nombre_salida = f"{Path(ruta_archivo).stem}_REPORTE_FORMATO_FINAL.xlsx"
        ruta_salida = str(ruta_en_reportes(nombre_salida, subcarpeta="Asistencia"))

        with pd.ExcelWriter(ruta_salida, engine='openpyxl', date_format='DD/MM/YYYY') as writer:
            reporte.to_excel(writer, sheet_name='Asistencia')

        print(f"\n✅ Proceso completado con éxito.\n📂 Archivo guardado en: {ruta_salida}")
    except Exception as e:
        print(f"\n❌ Ocurrió un error: {e}\nRevisa que el archivo no esté abierto en Excel.")


def reporte_emergencia():
    def tiene_valor(dato):
        if pd.isna(dato) or str(dato).lower() in ["<na>", "nan", "none", "null", ""]:
            return False
        return True

    try:
        sedf_base = cargar_bd_arcgis()
        if sedf_base is None or sedf_base.empty:
            print("❌ No se pudo cargar la BD de ArcGIS (o está vacía).")
            return

        while True:
            sedf = sedf_base
            if sedf.empty:
                print("❌ La capa está vacía.")
                break

            col_fecha = 'fecha_de_visita'
            if col_fecha in sedf.columns:
                if pd.api.types.is_numeric_dtype(sedf[col_fecha]):
                    sedf[col_fecha] = pd.to_datetime(sedf[col_fecha], unit='ms', errors='coerce')
                else:
                    sedf[col_fecha] = pd.to_datetime(sedf[col_fecha], errors='coerce')

            recientes = sedf.sort_values(by=col_fecha, ascending=False).head(10)
            print(f"\n--- ÚLTIMOS 10 REGISTROS ---")
            for _, fila in recientes.iterrows():
                f_print = fila[col_fecha].date() if col_fecha in fila else "S/F"
                print(f"ID: {fila['objectid']} | {fila.get('nombre_del_punto', 'S/N')} | {f_print}")

            entrada = input("\nIngrese el ObjectID seleccionado (o 's' para salir): ").lower()
            if entrada == 's': break

            try:
                idx = int(entrada)
                seleccion = recientes[recientes['objectid'] == idx]
                if seleccion.empty:
                    print("⚠️ ID no encontrado.")
                    continue

                d = seleccion.iloc[0].to_dict()
                transformer = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
                long, lat = d.get('longitud', 0), d.get('latitud', 0)
                coord_x, coord_y = transformer.transform(long, lat)

                fecha_limpia = d[col_fecha].strftime('%d-%m-%Y') if tiene_valor(d.get(col_fecha)) else "S/F"
                tecnico_formateado = str(d.get('field_50', '')).replace(",", ", ").replace("_", " ")

                lineas_reporte = [f"*INSP-TEC-{str(d.get('nombre_del_punto', 'S/N')).replace(' ','_').upper()}*"]
                campos = [
                    ("Sector", d.get('sector')),
                    ("Fecha", fecha_limpia),
                    ("Parroquia", d.get('parroquia')),
                    ("Cantón", d.get('cant_n')),
                    ("Longitud", long),
                    ("Latitud", lat),
                    ("UTM X", int(coord_x)),
                    ("UTM Y", int(coord_y)),
                    ("Tipo de evento", d.get('tipo_de_evento')),
                    ("Subtipo de evento", d.get('subtipo_de_evento')),
                    ("Cantidad de Personas Afectadas", d.get('personas_afectadas')),
                    ("Cantidad de Viviendas Afectadas", d.get('viviendas_afectadas')),
                    ("Otras afectaciones", d.get('otras_afectaciones')),
                    ("Descripción del evento", d.get('descripci_n_del_evento')),
                    ("Afectación en vía", d.get('nombre_de_v_a')),
                    ("Pertenece al inventario vial Provincial", d.get('inventario_vial')),
                    ("Comentarios Técnicos", d.get('comentarios_t_cnicos')),
                    ("Técnico GAD Provincial", tecnico_formateado),
                    ("Personal GAD Provincial", d.get('personal_gad_provincial')),
                    ("Personal GAD Cantonal", d.get('personal_gad_cantonal')),
                    ("Personal GAD Parroquial", d.get('personal_gad_parroquial')),
                ]

                for etiqueta, valor in campos:
                    if tiene_valor(valor):
                        lineas_reporte.append(f" - *{etiqueta}:* {valor}")

                area_dron = d.get('_rea_con_dron')
                if tiene_valor(area_dron) and float(area_dron) > 0:
                    lineas_reporte.append(f" - *Área levantada con dron:* {int(float(area_dron)/10000)}Ha")

                print("\n" + "\n".join(lineas_reporte))
            except ValueError:
                print("❌ Ingrese un número válido.")
    except Exception as e:
        print(f"❌ Error: {e}")


def descargar_bd_arcgis():
    print("\n--- TIPO DE BASE A DESCARGAR ---")
    print("1. BD Completa\n2. BD Resumen")
    opcion_bd = input("Seleccione una opción (1 o 2): ").strip()
    if opcion_bd not in ["1", "2"]:
        opcion_bd = "1"

    fecha_hoy = datetime.now().strftime("%y%m%d")
    nombre_archivo = f"{fecha_hoy}_BD_Riesgos_Resumen.xlsx" if opcion_bd == "2" else f"{fecha_hoy}_BD_Riesgos.xlsx"
    ruta_final = ruta_en_reportes(nombre_archivo, subcarpeta="BD_ArcGIS")

    columnas_a_borrar = [
        'globalid', 'CreationDate', 'Creator', 'EditDate', 'Editor',
        'direcci_n_del_gadpa', 'direcci_n_del_gadpa_other',
        'tipo_de_evento_other', 'field_50_other', 'SHAPE'
    ]

    try:
        sedf_completo = cargar_bd_arcgis()
        if sedf_completo is None:
            print("❌ No se pudo cargar la BD de ArcGIS en memoria.")
            return

        if not sedf_completo.empty:
            columnas_presentes = [c for c in columnas_a_borrar if c in sedf_completo.columns]
            sedf_completo = sedf_completo.drop(columns=columnas_presentes)

            if 'fecha_de_visita' in sedf_completo.columns:
                if pd.api.types.is_numeric_dtype(sedf_completo['fecha_de_visita']):
                    sedf_completo['fecha_de_visita'] = pd.to_datetime(sedf_completo['fecha_de_visita'], unit='ms', errors='coerce')
                else:
                    sedf_completo['fecha_de_visita'] = pd.to_datetime(sedf_completo['fecha_de_visita'], errors='coerce')
                sedf_completo = sedf_completo.sort_values(by='fecha_de_visita', ascending=False)
                sedf_completo['fecha_de_visita'] = sedf_completo['fecha_de_visita'].dt.strftime('%d/%m/%Y')

            if opcion_bd == "2":
                columnas_resumen = [
                    "objectid", "nombre_del_punto", "fecha_de_visita", "sector", "parroquia",
                    "cant_n", "tipo_de_evento", "fecha_y_hora_de_evento", "tipo_de_atenci_n",
                    "trailers", "tractores", "excavadoras", "cargadoras", "retro_excavadoras",
                    "volquetes", "cisternas", "drones", "cuadrilla", "operadores_y_conductores",
                    "personal_t_cnico", "en_intervenci_n"
                ]
                columnas_presentes_resumen = [c for c in columnas_resumen if c in sedf_completo.columns]
                sedf_salida = sedf_completo[columnas_presentes_resumen].copy()
            else:
                sedf_salida = sedf_completo

            sedf_salida.to_excel(ruta_final, index=False)
            print("-" * 45)
            print(f"✅ ¡Archivo guardado con éxito!\n📍 Ubicación: {ruta_final}")
            print("-" * 45)
        else:
            print("⚠️ No hay datos disponibles para descargar.")
    except Exception as e:
        print(f"❌ Ocurrió un error: {e}")


def set_cell_shading(cell, color_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def crear_reporte_aptos(lista_tablas, titulo_reporte, periodo_str):
    hoy = datetime.now()
    FechaCreacion = hoy.strftime("%d/%m/%Y %H:%M")
    nombre_archivo = hoy.strftime("%y%m%d%H%M - Ayuda Memoria %H%M") + ".docx"
    ruta_final = str(ruta_en_reportes(nombre_archivo, subcarpeta="AyudaMemoria"))
    ruta_plantilla = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\DocEnBlanco.docx"

    try:
        doc = Document(ruta_plantilla)
    except Exception:
        doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Aptos'
    font.size = Pt(12)
    font.color.rgb = RGBColor(40, 40, 40)

    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_fecha = p_fecha.add_run(f"Fecha de Actualización: {FechaCreacion}")
    run_fecha.font.size = Pt(10)

    p_sec = doc.add_paragraph()
    p_sec.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sec.paragraph_format.space_after = Pt(0)
    run_sec = p_sec.add_run("SECRETARÍA DE INFRAESTRUCTURAS")
    run_sec.bold = True

    p_coor = doc.add_paragraph()
    p_coor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_coor.paragraph_format.space_after = Pt(10)
    run_coor = p_coor.add_run("COORDINACIÓN DE RIESGOS")
    run_coor.bold = True
    run_coor.font.color.rgb = RGBColor(128, 128, 128)

    p_memoria = doc.add_paragraph()
    p_memoria.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_memoria.paragraph_format.space_after = Pt(0)
    run_memoria = p_memoria.add_run("AYUDA MEMORIA DE EVENTOS ATENDIDOS")
    run_memoria.font.size = Pt(14)
    run_memoria.bold = True

    p_periodo = doc.add_paragraph()
    p_periodo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_periodo.paragraph_format.space_after = Pt(10)
    run_periodo = p_periodo.add_run(f"PERÍODO: {periodo_str}")
    run_periodo.bold = True

    p2 = doc.add_paragraph(
        "Ayuda memoria correspondiente a los eventos atendidos por la Secretaría de Infraestructuras en la provincia del Azuay para el período correspondiente. "
        "A continuación, se detallan los recursos utilizados y las categorías de eventos."
    )
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for titulo_tabla, datos_matriz in lista_tablas:
        p_tabla = doc.add_paragraph()
        run_tabla = p_tabla.add_run(titulo_tabla)
        run_tabla.bold = True
        run_tabla.font.name = 'Aptos'

        filas = len(datos_matriz)
        columnas = len(datos_matriz[0])
        tabla = doc.add_table(rows=filas, cols=columnas)
        tabla.style = 'Table Grid'

        for i, fila in enumerate(datos_matriz):
            celdas = tabla.rows[i].cells
            for j, valor in enumerate(fila):
                texto = str(valor).upper() if i == 0 else str(valor)
                p = celdas[j].paragraphs[0]
                run = p.add_run(texto)
                run.font.name = 'Aptos'
                run.font.size = Pt(10)
                if i == 0:
                    run.bold = True
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_shading(celdas[j], "D9D9D9")
        doc.add_paragraph()

    num_siguiente = len(lista_tablas) + 1
    p_coes = doc.add_paragraph()
    run_coes = p_coes.add_run(f"{num_siguiente}. COES ACTIVADOS")
    run_coes.bold = True
    run_coes.font.name = 'Aptos'
    doc.add_paragraph()

    num_siguiente += 1
    p_res = doc.add_paragraph()
    run_res = p_res.add_run(f"{num_siguiente}. RESOLUCIONES EMITIDAS POR LA SNGR")
    run_res.bold = True
    run_res.font.name = 'Aptos'
    doc.add_paragraph()

    doc.save(ruta_final)
    print(f"\n[ÉXITO] Reporte generado correctamente en:\n{ruta_final}")


def generar_reporte_desde_arcgis():
    try:
        df = cargar_bd_arcgis()
        if df is None or df.empty:
            print("❌ No se pudo cargar la BD de ArcGIS (o está vacía).")
            return
        df = df.copy()

        if 'fecha_de_visita' in df.columns:
            if pd.api.types.is_numeric_dtype(df['fecha_de_visita']):
                df['fecha_de_visita'] = pd.to_datetime(df['fecha_de_visita'], unit='ms', errors='coerce')
            else:
                df['fecha_de_visita'] = pd.to_datetime(df['fecha_de_visita'], errors='coerce')
        df['anio'] = df['fecha_de_visita'].dt.year

        maquinaria_pesada = ['trailers', 'tractores', 'excavadoras', 'cargadoras',
                             'retro_excavadoras', 'mini_cargadoras', 'volquetes', 'cisternas']
        maquinaria_liviana = ['camionetas', 'drones']
        personal_operativo = ['cuadrilla', 'operadores_y_conductores']
        personal_tecnico = ['personal_t_cnico']

        cols_existentes = [col for col in (maquinaria_pesada + maquinaria_liviana + personal_operativo + personal_tecnico) if col in df.columns]
        df[cols_existentes] = df[cols_existentes].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0)

        print("--- FILTROS DEL REPORTE DE RIESGOS ---")
        print("1. 2025\n2. 2026\n3. Fecha personalizada")
        opcion_fecha = input("Seleccione el período (1/2/3): ")

        if opcion_fecha == "1":
            df_fecha = df[df['anio'] == 2025]
            periodo_str = "2025"
        elif opcion_fecha == "2":
            df_fecha = df[df['anio'] == 2026]
            periodo_str = "2026"
        elif opcion_fecha == "3":
            fecha_inicio_str = input("Ingrese la fecha de inicio (YYYY-MM-DD): ")
            fecha_fin_str = input("Ingrese la fecha de fin (YYYY-MM-DD): ")
            try:
                fecha_inicio = pd.to_datetime(fecha_inicio_str)
                fecha_fin = pd.to_datetime(fecha_fin_str + ' 23:59:59')
                df_fecha = df[(df['fecha_de_visita'] >= fecha_inicio) & (df['fecha_de_visita'] <= fecha_fin)]
                periodo_str = f"DEL {fecha_inicio.strftime('%d-%m-%Y')} AL {fecha_fin.strftime('%d-%m-%Y')}"
            except Exception:
                df_fecha = df[df['anio'] == 2025]
                periodo_str = "2025"
        else:
            df_fecha = df[df['anio'] == 2025]
            periodo_str = "2025"

        col_tipo_atencion = 'tipo_de_atenci_n' if 'tipo_de_atenci_n' in df.columns else 'tipo_de_atencion'
        tipos_disponibles = sorted(df[col_tipo_atencion].dropna().unique().tolist())
        print("\nSeleccione el tipo de atención:\n0. Todos")
        for i, tipo in enumerate(tipos_disponibles, 1):
            print(f"{i}. {tipo}")

        opcion_tipo = input("Ingrese el número de su opción: ")
        if opcion_tipo in ["0", ""]:
            tipo_filtro = "Todos"
            df_filtrado = df_fecha
        else:
            try:
                idx = int(opcion_tipo) - 1
                tipo_filtro = tipos_disponibles[idx]
                df_filtrado = df_fecha[df_fecha[col_tipo_atencion] == tipo_filtro]
            except Exception:
                tipo_filtro = "Todos"
                df_filtrado = df_fecha

        if df_filtrado.empty:
            print(f"\nNo se encontraron datos para el período '{periodo_str}' y tipo: {tipo_filtro}")
            return

        maq_pesada_validas = [col for col in maquinaria_pesada if col in df.columns]
        maq_liviana_validas = [col for col in maquinaria_liviana if col in df.columns]
        pers_operativo_validos = [col for col in personal_operativo if col in df.columns]
        pers_tecnico_validos = [col for col in personal_tecnico if col in df.columns]

        resumen_final = {
            'Concepto': [
                'Cantidad Total de Eventos', 'Total Maquinaria Pesada',
                'Total Maquinaria Liviana', 'Total Operadores/Cuadrilla', 'Total Personal Técnico'
            ],
            'Valor': [
                len(df_filtrado),
                int(df_filtrado[maq_pesada_validas].values.sum()) if maq_pesada_validas else 0,
                int(df_filtrado[maq_liviana_validas].values.sum()) if maq_liviana_validas else 0,
                int(df_filtrado[pers_operativo_validos].values.sum()) if pers_operativo_validos else 0,
                int(df_filtrado[pers_tecnico_validos].values.sum()) if pers_tecnico_validos else 0
            ]
        }
        df_tabla_totales = pd.DataFrame(resumen_final)
        matriz_totales = [df_tabla_totales.columns.tolist()] + df_tabla_totales.values.tolist()

        col_evento = 'tipo_de_evento' if 'tipo_de_evento' in df.columns else 'tipo_evento'
        resumen_eventos = df_filtrado[col_evento].value_counts().reset_index()
        resumen_eventos.columns = ['Tipo de Evento', 'Cantidad']
        matriz_eventos = [resumen_eventos.columns.tolist()] + resumen_eventos.values.tolist()

        col_canton = 'cant_n' if 'cant_n' in df.columns else 'canton'
        resumen_geo = df_filtrado.groupby([col_canton, 'parroquia']).size().reset_index(name='No. DE REGISTROS')
        resumen_geo.columns = ['CANTÓN', 'PARROQUIA', 'No. DE REGISTROS']
        matriz_geo = [resumen_geo.columns.tolist()] + resumen_geo.values.tolist()

        lista_de_tablas = [
            ("1. RESUMEN DE CANTIDADES TOTALES", matriz_totales),
            ("2. CANTIDAD POR TIPO DE EVENTO", matriz_eventos),
            ("3. REGISTROS POR UBICACIÓN GEOGRÁFICA", matriz_geo)
        ]

        if tipo_filtro == "Todos":
            resumen_atencion = df_filtrado[col_tipo_atencion].value_counts().reset_index()
            resumen_atencion.columns = ['Tipo de Atención', 'Cantidad']
            matriz_atencion = [resumen_atencion.columns.tolist()] + resumen_atencion.values.tolist()
            lista_de_tablas.append(("4. CANTIDAD POR TIPO DE ATENCIÓN", matriz_atencion))

        titulo_doc = f"PERIODO: {periodo_str} - ATENCIÓN: {tipo_filtro}"
        crear_reporte_aptos(lista_de_tablas, titulo_doc, periodo_str)
    except Exception as e:
        print(f"Ocurrió un error al procesar los datos de ArcGIS: {e}")


def generar_tablas_resumen_puntos():
    warnings.filterwarnings("ignore")
    print("\n--- OPCIONES DE FILTRADO ---")
    print("1. Filtrar solo el año 2025\n2. Filtrar solo el año 2026\n3. Rango de fechas personalizado")
    print("4. Frentes Activos (Solo en intervención: 'Sí')\n5. Solo emergencias (Filtro por columna tipo_de_atenci_n)")
    opcion_filtro = input("Elige una opción (1, 2, 3, 4 o 5): ").strip()

    fecha_inicio_dt = None
    fecha_fin_dt = None

    if opcion_filtro == "3":
        fecha_inicio_str = input("Ingresa la fecha de INICIO (formato DD/MM/AAAA): ").strip()
        fecha_fin_str = input("Ingresa la fecha de FIN (formato DD/MM/AAAA): ").strip()
        fecha_inicio_dt = pd.to_datetime(fecha_inicio_str, format="%d/%m/%Y", errors="coerce")
        fecha_fin_dt = pd.to_datetime(fecha_fin_str, format="%d/%m/%Y", errors="coerce")
        if pd.isna(fecha_inicio_dt) or pd.isna(fecha_fin_dt):
            opcion_filtro = "2"
    elif opcion_filtro not in ["1", "2", "3", "4", "5"]:
        opcion_filtro = "2"

    try:
        df = cargar_bd_arcgis()
        if df is None or df.empty:
            print("❌ No se pudo cargar la BD de ArcGIS (o está vacía).")
            return
        df = df.copy()

        if "fecha_de_visita" in df.columns:
            if pd.api.types.is_numeric_dtype(df["fecha_de_visita"]):
                df["fecha_de_visita"] = pd.to_datetime(df["fecha_de_visita"], unit="ms", errors="coerce")
            else:
                df["fecha_de_visita"] = pd.to_datetime(df["fecha_de_visita"], errors="coerce")

        df_filtrado = df.copy()
        mascara = None
        if opcion_filtro == "1":
            mascara = df_filtrado["fecha_de_visita"].dt.year == 2025
        elif opcion_filtro == "2":
            mascara = df_filtrado["fecha_de_visita"].dt.year == 2026
        elif opcion_filtro == "3":
            mascara = (df_filtrado["fecha_de_visita"] >= fecha_inicio_dt) & (
                df_filtrado["fecha_de_visita"] <= fecha_fin_dt + pd.Timedelta(days=1, seconds=-1)
            )
        elif opcion_filtro == "4":
            if "en_intervenci_n" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["en_intervenci_n"].astype(str).str.strip().str.lower().isin(["sí", "si"])]
            mascara = None
        elif opcion_filtro == "5":
            if "tipo_de_atenci_n" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["tipo_de_atenci_n"].astype(str).str.strip().str.lower() == "emergencia"]
            mascara = None

        if opcion_filtro in ["1", "2", "3"] and mascara is not None:
            if "en_intervenci_n" in df_filtrado.columns:
                df_filtrado = df_filtrado[mascara | df_filtrado["en_intervenci_n"].astype(str).str.strip().str.lower().isin(["sí", "si"])]
            else:
                df_filtrado = df_filtrado[mascara]

        columnas_equipos = {
            "trailers": ("Trailer", "Trailers"), "tractores": ("Tractor", "Tractores"),
            "excavadoras": ("Excavadora", "Excavadoras"), "cargadoras": ("Cargadora", "Cargadoras"),
            "retro_excavadoras": ("Retroexcavadora", "Retroexcavadoras"), "mini_cargadoras": ("Minicargadora", "Minicargadoras"),
            "volquetes": ("Volquete", "Volquetes"), "cisternas": ("Cisterna", "Cisternas"),
            "camionetas": ("Camioneta", "Camionetas"), "drones": ("Dron", "Drones")
        }

        def juntar_equipos(row):
            equipos = []
            for col, (singular, plural) in columnas_equipos.items():
                if col in row and pd.notna(row[col]):
                    try:
                        cant = int(float(row[col]))
                        if cant == 1: equipos.append(f"1 {singular}")
                        elif cant > 1: equipos.append(f"{cant} {plural}")
                    except Exception: pass
            return ", ".join(equipos) if equipos else "Sin equipos"

        df_filtrado["equipos"] = df_filtrado.apply(juntar_equipos, axis=1)

        columnas_personal = {
            "cuadrilla": ("Cuadrilla", "Cuadrillas"),
            "operadores_y_conductores": ("Operador", "Operadores"),
            "personal_t_cnico": ("Técnico", "Técnicos")
        }

        def juntar_personal(row):
            personal = []
            for col, (singular, plural) in columnas_personal.items():
                if col in row and pd.notna(row[col]):
                    try:
                        cant = int(float(row[col]))
                        if cant == 1: personal.append(f"1 {singular}")
                        elif cant > 1: personal.append(f"{cant} {plural}")
                    except Exception: pass
            return ", ".join(personal) if personal else "Sin personal"

        df_filtrado["personal"] = df_filtrado.apply(juntar_personal, axis=1)

        columnas_finales = [
            col for col in [
                "nombre_del_punto", "parroquia", "cant_n", "fecha_de_visita",
                "subtipo_de_evento", "equipos", "personal"
            ] if col in df_filtrado.columns
        ]
        if "en_intervenci_n" in df_filtrado.columns: columnas_finales.append("en_intervenci_n")
        if "descripci_n_intervenci_n" in df_filtrado.columns: columnas_finales.append("descripci_n_intervenci_n")

        df_resultado = df_filtrado[columnas_finales].copy()
        if "fecha_de_visita" in df_resultado.columns:
            df_resultado = df_resultado.sort_values(by="fecha_de_visita", ascending=True)
            df_resultado["fecha_de_visita"] = df_resultado["fecha_de_visita"].dt.strftime("%d/%m/%Y")

        df_resultado.rename(columns={
            "nombre_del_punto": "Sector", "parroquia": "Parroquia", "cant_n": "Cantón",
            "fecha_de_visita": "Fecha de visita", "subtipo_de_evento": "Subtipo de evento",
            "en_intervenci_n": "En intervención", "descripci_n_intervenci_n": "Actividades realizadas"
        }, inplace=True)
        df_resultado.insert(0, "No.", range(1, len(df_resultado) + 1))

        if "cant_n" in df_filtrado.columns and "parroquia" in df_filtrado.columns:
            df_atenciones = df_filtrado.groupby(["cant_n", "parroquia"]).size().reset_index(name="Total Atenciones")
        else:
            df_atenciones = pd.DataFrame()

        cols_a_sumar = list(columnas_equipos.keys()) + list(columnas_personal.keys())
        for col in cols_a_sumar:
            if col in df_filtrado.columns:
                df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce").fillna(0)
            else:
                df_filtrado[col] = 0

        totales = df_filtrado[cols_a_sumar].sum()
        df_totales = pd.DataFrame([
            {"Recursos": "Trailers", "Cantidad": totales.get("trailers", 0)},
            {"Recursos": "Tractores", "Cantidad": totales.get("tractores", 0)},
            {"Recursos": "Excavadoras", "Cantidad": totales.get("excavadoras", 0)},
            {"Recursos": "Cargadoras", "Cantidad": totales.get("cargadoras", 0)},
            {"Recursos": "Retroexcavadoras", "Cantidad": totales.get("retro_excavadoras", 0)},
            {"Recursos": "Volquetes", "Cantidad": totales.get("volquetes", 0)},
            {"Recursos": "Cisternas", "Cantidad": totales.get("cisternas", 0)},
            {"Recursos": "Camionetas", "Cantidad": totales.get("camionetas", 0)},
            {"Recursos": "No. vuelo de drones", "Cantidad": totales.get("drones", 0)},
            {"Recursos": "Personal Operativo", "Cantidad": totales.get("operadores_y_conductores", 0) + totales.get("cuadrilla", 0)},
            {"Recursos": "Personal Técnico", "Cantidad": totales.get("personal_t_cnico", 0)},
        ])

        fecha_ejecucion_str = datetime.now().strftime("%y%m%d")
        if opcion_filtro == "1": nombre_filtro = "Año2025"
        elif opcion_filtro == "2": nombre_filtro = "Año2026"
        elif opcion_filtro == "3":
            fi_str = fecha_inicio_dt.strftime("%y%m%d") if fecha_inicio_dt is not None else "000000"
            ff_str = fecha_fin_dt.strftime("%y%m%d") if fecha_fin_dt is not None else "000000"
            nombre_filtro = f"Rango{fi_str}-{ff_str}"
        elif opcion_filtro == "4": nombre_filtro = "FrentesActivos"
        elif opcion_filtro == "5": nombre_filtro = "SoloEmergencias"
        else: nombre_filtro = "SinFiltro"

        ruta_salida = str(ruta_en_reportes(f"{fecha_ejecucion_str}-{nombre_filtro}.xlsx", subcarpeta="TablasResumen"))
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, sheet_name="Resumen General", index=False)
            df_atenciones.to_excel(writer, sheet_name="Atenciones por Sector", index=False)
            df_totales.to_excel(writer, sheet_name="Total Equipos", index=False)

        wb = load_workbook(ruta_salida)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 25

        wb.save(ruta_salida)
        print(f"\n¡PROCESO COMPLETADO! Archivo generado en: {ruta_salida}")
    except Exception as e:
        print(f"\nError al generar las tablas de resumen: {e}")


# ==========================================
# VIDEO PRONOSTICO INHAMI
# ==========================================
USUARIO = os.getenv("INAMHI_USER", "ExternoINAMHI")
CLAVE = os.getenv("INAMHI_PASS", "externoinamhi2025*")
BASE_URL = "https://infohm.inamhi.gob.ec/remote.php/dav/files/ExternoINAMHI"

ruta_mascara = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\MascaraBlancaFueraAzuay2.gpkg"
ruta_shp_cantones = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\AzuayXCantones.gpkg"
ruta_extension = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\ExtensionAzuay.gpkg"
ruta_poblados = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\PobladosPrincipalesAzuay2.gpkg"
ruta_logo_1 = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\logo prefectura.png"
ruta_logo_2 = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\Diseño sin título.png"

CARPETA_TIFS = "temp_tifs_descargados"
CARPETA_PNGS = "Procesados_PNG_1080p_Final"

ancho_px, alto_px = 1920, 1080
dpi_salida = 100

colors_rgb_norm = [
    (0 / 50, (49 / 255, 130 / 255, 189 / 255)),
    (5 / 50, (99 / 255, 172 / 255, 186 / 255)),
    (10 / 50, (145 / 255, 203 / 255, 168 / 255)),
    (15 / 50, (189 / 255, 228 / 255, 157 / 255)),
    (20 / 50, (228 / 255, 244 / 255, 178 / 255)),
    (25 / 50, (254 / 255, 255 / 255, 191 / 255)),
    (30 / 50, (254 / 255, 217 / 255, 139 / 255)),
    (35 / 50, (253 / 255, 181 / 255, 106 / 255)),
    (40 / 50, (246 / 255, 131 / 255, 78 / 255)),
    (45 / 50, (230 / 255, 85 / 255, 55 / 255)),
    (50 / 50, (200 / 255, 30 / 255, 30 / 255)),
]
precip_cmap = LinearSegmentedColormap.from_list("precip_azuay", colors_rgb_norm)
precip_norm = Normalize(vmin=0, vmax=50)


def verificar_carpeta_tifs(fecha):
    ruta = f"/WRF/{fecha}/Azuay/prec/tif/"
    url = BASE_URL + ruta
    try:
        res = requests.request("PROPFIND", url, auth=HTTPBasicAuth(USUARIO, CLAVE), headers={"Depth": "1"}, verify=False, timeout=10)
        if res.status_code == 207:
            root = ET.fromstring(res.content)
            ns = {"d": "DAV:"}
            tifs = [h.text for h in root.findall(".//d:href", ns) if h.text and h.text.lower().endswith((".tif", ".tiff"))]
            return True, tifs, url
        return False, [], url
    except Exception as e:
        print(f"Error al conectar: {e}")
        return False, [], url


def descargar_archivos(lista_enlaces, carpeta_destino):
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)
    archivos_locales = []
    for enlace in sorted(lista_enlaces):
        nombre = os.path.basename(enlace)
        ruta_local = os.path.join(carpeta_destino, nombre)
        print(f"📥 Descargando {nombre}...", end="\r")
        r = requests.get(f"https://infohm.inamhi.gob.ec{enlace}", auth=HTTPBasicAuth(USUARIO, CLAVE), verify=False)
        if r.status_code == 200:
            with open(ruta_local, "wb") as f:
                f.write(r.content)
            archivos_locales.append(ruta_local)
    print("\n✅ Descarga de TIFs completada.")
    return archivos_locales


def procesar_mapas(archivos_tif_locales, fecha_str):
    if not os.path.exists(CARPETA_PNGS):
        os.makedirs(CARPETA_PNGS)

    tiempos_archivos = []
    patron = r"(\d{4}-\d{2}-\d{2})_(\d{2})h(\d{2})-(\d{2})h(\d{2})"
    for ruta in archivos_tif_locales:
        match = re.search(patron, os.path.basename(ruta))
        if match:
            f_s, h1, m1, _, _ = match.groups()
            tiempos_archivos.append(datetime.strptime(f"{f_s} {h1}:{m1}", "%Y-%m-%d %H:%M"))

    if not tiempos_archivos:
        print("❌ Error: No se pudieron extraer fechas para la barra de progreso.")
        return []

    fecha_min = min(tiempos_archivos)
    fecha_max = max(tiempos_archivos)
    if fecha_min == fecha_max:
        fecha_max += timedelta(hours=1)

    dias_unicos = sorted(list(set(d.replace(hour=0, minute=0) for d in tiempos_archivos)))

    try:
        mascara_gdf = gpd.read_file(ruta_mascara)
        cantones_gdf = gpd.read_file(ruta_shp_cantones)
        extension_gdf = gpd.read_file(ruta_extension)
        poblados_gdf = gpd.read_file(ruta_poblados)
        img_logo_1_raw = mpimg.imread(ruta_logo_1) if os.path.exists(ruta_logo_1) else None
        img_logo_2_raw = mpimg.imread(ruta_logo_2) if os.path.exists(ruta_logo_2) else None
    except Exception as e:
        print(f"❌ Error al cargar capas base: {e}")
        return []

    pngs_generados = []
    dias_semana = {0: "lunes", 1: "martes", 2: "miércoles", 3: "jueves", 4: "viernes", 5: "sábado", 6: "domingo"}
    dt_fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
    nombre_dia = dias_semana[dt_fecha.weekday()]
    fecha_formateada = dt_fecha.strftime("%d/%m/%Y")
    texto_subtitulo = f"act: {nombre_dia} {fecha_formateada}"

    for i, ruta_in in enumerate(archivos_tif_locales):
        archivo = os.path.basename(ruta_in)
        ruta_out = os.path.join(CARPETA_PNGS, f"Mapa_{os.path.splitext(archivo)[0]}.png")
        dt_actual = tiempos_archivos[i]

        try:
            with rasterio.open(ruta_in) as src:
                crs_raster = src.crs
                resolucion_nueva = 0.001
                dst_transform, dst_width, dst_height = calculate_default_transform(
                    src.crs, crs_raster, src.width, src.height, *src.bounds,
                    resolution=(resolucion_nueva, resolucion_nueva),
                )
                data = np.zeros((dst_height, dst_width), dtype=src.dtypes[0])
                reproject(
                    source=rasterio.band(src, 1),
                    destination=data,
                    src_transform=src.transform,
                    src_crs=src.crs,
                    dst_transform=dst_transform,
                    dst_crs=crs_raster,
                    resampling=Resampling.lanczos,
                )
                minx = dst_transform.c
                maxy = dst_transform.f
                maxx = minx + dst_width * dst_transform.a
                miny = maxy + dst_height * dst_transform.e
                extent_raster = (minx, maxx, miny, maxy)

                mascara_proj = mascara_gdf.to_crs(crs_raster)
                cantones_proj = cantones_gdf.to_crs(crs_raster)
                extension_proj = extension_gdf.to_crs(crs_raster)
                poblados_proj = poblados_gdf.to_crs(crs_raster)

            fig, ax = plt.subplots(figsize=(ancho_px / dpi_salida, alto_px / dpi_salida), dpi=dpi_salida)
            fig.patch.set_facecolor("white")
            ax.set_axis_off()

            img = ax.imshow(data, cmap=precip_cmap, norm=precip_norm, extent=extent_raster, interpolation="nearest", zorder=1)
            mascara_proj.plot(ax=ax, facecolor="white", edgecolor="none", zorder=2)
            poblados_proj.plot(ax=ax, color="black", edgecolor="white", linewidth=0.5, markersize=15, zorder=3)

            for x, y, label in zip(poblados_proj.geometry.centroid.x, poblados_proj.geometry.centroid.y, poblados_proj["nombre"]):
                if label:
                    txt_pob = ax.annotate(label, xy=(x, y), xytext=(0, 4), textcoords="offset points", fontsize=5, color="black", ha="center", va="bottom", zorder=4)
                    txt_pob.set_path_effects([path_effects.withStroke(linewidth=0, foreground="white")])

            cantones_proj.plot(ax=ax, facecolor="none", edgecolor="black", linewidth=0.5, alpha=0.3, zorder=5)
            for x, y, label in zip(cantones_proj.geometry.centroid.x, cantones_proj.geometry.centroid.y, cantones_proj["CANTON"]):
                txt = ax.text(x, y, label, fontsize=8, fontweight="bold", color="white", ha="center", va="center", zorder=6)
                txt.set_path_effects([path_effects.withStroke(linewidth=2, foreground="black")])

            titulo_obj = ax.text(0.5, 0.96, "MAPA DE PRONÓSTICO DE PRECIPITACIONES", transform=ax.transAxes, fontsize=24, fontweight="bold", color="black", ha="center", va="top", zorder=15)
            titulo_obj.set_path_effects([path_effects.withStroke(linewidth=5, foreground="white")])

            subtitulo_obj = ax.text(0.5, 0.92, texto_subtitulo, transform=ax.transAxes, fontsize=16, color="black", ha="center", va="top", zorder=15)
            subtitulo_obj.set_path_effects([path_effects.withStroke(linewidth=3, foreground="white")])

            ax_prog = fig.add_axes([0.25, 0.88, 0.5, 0.015])
            ax_prog.set_xlim(fecha_min, fecha_max)
            ax_prog.set_yticks([])
            ax_prog.set_xticks(dias_unicos)
            ax_prog.set_xticklabels([d.strftime("%d/%m") for d in dias_unicos], fontsize=10, fontweight="bold", color="#333333")
            ax_prog.axhline(y=0.5, color="#e0e0e0", linewidth=8, zorder=1, solid_capstyle="round")
            ax_prog.scatter(dt_actual, 0.5, color="#e74c3c", s=100, zorder=3, edgecolor="white", linewidth=1.5)
            ax_prog.text(dt_actual, 1.8, dt_actual.strftime("%H:%M"), color="#e74c3c", fontsize=11, fontweight="bold", ha="center", transform=ax_prog.get_xaxis_transform())
            for s in ax_prog.spines.values(): s.set_visible(False)
            ax_prog.tick_params(axis="x", which="both", length=0, pad=8)

            match = re.search(patron, archivo)
            if match:
                f_str, h1, m1, h2, m2 = match.groups()
                dt_ini = datetime.strptime(f"{f_str} {h1}:{m1}", "%Y-%m-%d %H:%M")
                dt_fin = datetime.strptime(f"{f_str} {h2}:{m2}", "%Y-%m-%d %H:%M")
                if dt_fin <= dt_ini: dt_fin += timedelta(days=1)
                texto_display = f"Inicio: {dt_ini.strftime('%d/%m/%Y %H:%M')}\nFin: {dt_fin.strftime('%d/%m/%Y %H:%M')}"
                txt_fecha = ax.text(0.02, 0.8, texto_display, transform=ax.transAxes, fontsize=12, fontweight="bold", family="monospace", verticalalignment="top", color="black", zorder=10)
                txt_fecha.set_path_effects([path_effects.withStroke(linewidth=3, foreground="white")])

            cax = fig.add_axes([0.35, 0.12, 0.3, 0.025])
            cb = fig.colorbar(img, cax=cax, orientation="horizontal", ticks=[0, 10, 20, 30, 40, 50])
            cb.set_label("Precipitación (mm)", fontsize=10, fontweight="bold")
            cb.ax.tick_params(labelsize=9)

            minx, miny, maxx, maxy = extension_proj.total_bounds
            ax.set_xlim(minx, maxx)
            ax.set_ylim(miny, maxy)
            ax.set_aspect("equal")

            if img_logo_1_raw is not None:
                zoom_1 = (alto_px * 0.07) / img_logo_1_raw.shape[0]
                ab1 = AnnotationBbox(OffsetImage(img_logo_1_raw, zoom=zoom_1, resample=True), (0.98, 0.35), xycoords="axes fraction", box_alignment=(1.0, 1.0), frameon=False, pad=0.0, zorder=20)
                ax.add_artist(ab1)

            if img_logo_2_raw is not None:
                zoom_2 = (alto_px * 0.05) / img_logo_2_raw.shape[0]
                ab2 = AnnotationBbox(OffsetImage(img_logo_2_raw, zoom=zoom_2, resample=True), (0.98, 0.23), xycoords="axes fraction", box_alignment=(1.0, 1.0), frameon=False, pad=0.0, zorder=20)
                ax.add_artist(ab2)

            plt.savefig(ruta_out, dpi=dpi_salida, bbox_inches="tight", pad_inches=0.1)
            plt.close(fig)
            pngs_generados.append(ruta_out)
            print(f"[{i + 1}/{len(archivos_tif_locales)}] Listo: {os.path.basename(ruta_out)}")
        except Exception as e:
            if "fig" in locals(): plt.close(fig)
            print(f"❌ Error procesando {archivo}: {e}")

    return sorted(pngs_generados)


def procesar_video_wrf():
    hoy = datetime.now().strftime("%Y-%m-%d")
    print(f"🔍 Verificando datos de hoy ({hoy}) en el servidor del INAMHI...")
    existe_hoy, lista_hoy, _ = verificar_carpeta_tifs(hoy)

    print("\n--- MENÚ DE DESCARGA Y MAPEO WRF ---")
    if existe_hoy and lista_hoy:
        print(f"✅ ¡Buenas noticias! Hay {len(lista_hoy)} archivos TIF de HOY disponibles.\n1. Descargar y procesar datos de HOY ({hoy})")
    else:
        print(f"❌ No se encontraron TIFs para hoy ({hoy}) todavía.\n1. Reintentar hoy (por si acaso)")

    print("2. Procesar datos de AYER\n3. Ingresar una FECHA específica (YYYY-MM-DD)")
    opcion = input("\nSelecciona una opción (1, 2 o 3): ").strip()

    if opcion == "1": fecha_final = hoy
    elif opcion == "2": fecha_final = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    elif opcion == "3": fecha_final = input("⌨️ Introduce la fecha (ej. 2026-03-25): ").strip()
    else: return

    ejecutar_flujo(fecha_final)


def ejecutar_flujo(fecha):
    existe, archivos_tifs, url_final = verificar_carpeta_tifs(fecha)
    if not existe or not archivos_tifs:
        print(f"❌ Error: No se encontraron archivos TIF en el servidor para la fecha {fecha}.\nRuta buscada: {url_final}")
        return

    print(f"\n🚀 Iniciando descarga de {len(archivos_tifs)} TIFs para {fecha}...")
    tifs_locales = descargar_archivos(archivos_tifs, CARPETA_TIFS)
    pngs_finales = procesar_mapas(tifs_locales, fecha)

    if not pngs_finales:
        print("❌ No se generaron mapas PNG. No se puede crear el video.")
        return

    print("\n🎬 Ensamblando video final...")
    try:
        clip = ImageSequenceClip(pngs_finales, fps=0.5)
        nombre_video = f"Pronostico_Azuay_{fecha}_Final.mp4"
        ruta_video = str(ruta_en_reportes(nombre_video, subcarpeta="Pronósticos INHAMI"))
        clip.write_videofile(ruta_video, codec="libx264", audio=False, fps=1)
        print(f"\n✅ ¡PROCESO COMPLETADO! Video listo: {ruta_video}")
    except Exception as e:
        print(f"❌ Error al crear el video: {e}")


def timestamp_videos():
    RUTA_CARPETA = input("Ingrese el directorio del proyecto: ").strip().strip('"').strip("'")
    print("\n--- SELECCIÓN DE MAPA BASE ---")
    print("1. Descarga normal por internet (Esri World Imagery)")
    print("2. Archivo local (GeoTIFF .tif)")
    print("3. Link de servicio ArcGIS (MapServer o capa de teselas)")
    opcion_mapa = input("Seleccione una opción (1/2/3): ").strip()

    RUTA_MAPA_LOCAL = ""
    LINK_ARCGIS = ""

    if opcion_mapa == '2':
        RUTA_MAPA_LOCAL = input("Ingrese la ruta completa del archivo .tif: ").strip().strip('"').strip("'")
        if not os.path.exists(RUTA_MAPA_LOCAL):
            opcion_mapa = '1'
    elif opcion_mapa == '3':
        LINK_ARCGIS = input("Pegue el link del servicio de ArcGIS: ").strip()
        if "MapServer" in LINK_ARCGIS and "{z}" not in LINK_ARCGIS:
            if not LINK_ARCGIS.endswith("/"): LINK_ARCGIS += "/"
            LINK_ARCGIS += "tile/{z}/{y}/{x}"

    DISTANCIA_ZOOM = 0.0015
    NUEVO_ALTO = 720
    SALTAR_CUADROS_MAPA = 3

    PATH_SHAPE = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\AzuayXParroquias.gpkg"
    PATH_LOGO = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\logo prefectura.png"

    if not os.path.exists(RUTA_CARPETA):
        raise FileNotFoundError(f"❌ La ruta especificada no existe: {RUTA_CARPETA}")

    archivos = os.listdir(RUTA_CARPETA)
    srt_files = [f for f in archivos if f.lower().endswith('.srt')]
    mp4_files = [f for f in archivos if f.lower().endswith('.mp4')]

    if not srt_files or not mp4_files:
        raise FileNotFoundError("❌ No se encontró el archivo .SRT o .MP4")

    path_video_origen = os.path.join(RUTA_CARPETA, mp4_files[0])
    path_srt = os.path.join(RUTA_CARPETA, srt_files[0])
    path_video_destino = os.path.join(RUTA_CARPETA, "vuelo_75_25_cinematico_logo_video.mp4")
    path_miniatura = os.path.join(RUTA_CARPETA, "miniatura_75_25_logo_video.png")

    parroquias_gdf = gpd.read_file(PATH_SHAPE)
    parroquias_gdf.set_crs(epsg=32717, inplace=True)

    logo_img = None
    if os.path.exists(PATH_LOGO):
        logo_img = Image.open(PATH_LOGO).convert("RGBA")
        logo_img.thumbnail((180, 180), Image.Resampling.LANCZOS)

    try:
        font_stamp = ImageFont.truetype("arial.ttf", size=14)
    except Exception:
        font_stamp = ImageFont.load_default()

    def convert_wgs84_to_utm17s(lat, lon):
        transformer = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
        easting, northing = transformer.transform(lon, lat)
        return round(easting), round(northing)

    def obtener_parroquia_canton(utm_e, utm_n):
        punto = Point(utm_e, utm_n).buffer(3)
        fila = parroquias_gdf[parroquias_gdf.geometry.intersects(punto)]
        if not fila.empty:
            parroquia = fila.iloc[0]['PARROQUIA'].title()
            canton = fila.iloc[0]['CANTON'].title()
        else:
            parroquia = "Desconocida"
            canton = "Desconocido"
        return parroquia, canton

    def estampar_cuadro_video(frame_bgr, lat, lon, alt, frame_idx, fps, empresa="GAD Provincial del Azuay"):
        img_pil = Image.fromarray(cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)).convert("RGBA")
        draw = ImageDraw.Draw(img_pil)
        utm_e, utm_n = convert_wgs84_to_utm17s(lat, lon)
        parroquia, canton = obtener_parroquia_canton(utm_e, utm_n)

        segundos_totales = int(frame_idx / fps)
        horas = segundos_totales // 3600
        minutos = (segundos_totales % 3600) // 60
        segundos = segundos_totales % 60
        tiempo_video = f"{horas:02d}:{minutos:02d}:{segundos:02d}"

        texto = (
            f"Tiempo: {tiempo_video}\n"
            f"Altura: {alt:.1f} m\n"
            f"Coord. WGS84: {lat:.6f}, {lon:.6f}\n"
            f"Coord. UTM17S: {utm_e}, {utm_n}\n"
            f"Parroquia: {parroquia}\n"
            f"Cantón: {canton}\n"
            f"{empresa}"
        )

        text_lines = texto.split("\n")
        max_width = max(draw.textlength(line, font=font_stamp) for line in text_lines)
        line_height = font_stamp.getbbox("A")[3] - font_stamp.getbbox("A")[1] + 4
        box_height = line_height * len(text_lines) + 10
        margin = 15

        text_box = Image.new("RGBA", img_pil.size)
        draw_box = ImageDraw.Draw(text_box)
        draw_box.rectangle([margin - 5, margin - 5, margin + max_width + 10, margin + box_height], fill=(0, 0, 0, 170))
        for i, line in enumerate(text_lines):
            draw_box.text((margin, margin + i * line_height), line, font=font_stamp, fill="white")
        img_pil = Image.alpha_composite(img_pil, text_box)

        if logo_img:
            img_w, _ = img_pil.size
            logo_w, _ = logo_img.size
            position = (img_w - logo_w - margin, margin)
            img_pil.paste(logo_img, position, mask=logo_img)

        return cv2.cvtColor(np.array(img_pil.convert("RGB")), cv2.COLOR_RGB2BGR)

    with open(path_srt, "r", encoding="utf-8") as f:
        contenido = f.read()

    patron = r"FrameCnt:\s*(\d+).*?\[latitude:\s*([\d.-]+)\]\s*\[longitude:\s*([\d.-]+)\]\s*\[rel_alt:\s*([\d.-]+)"
    matches = re.findall(patron, contenido, re.DOTALL)
    telemetria = {int(m[0]): (float(m[1]), float(m[2]), float(m[3])) for m in matches}
    latitudes = [float(m[1]) for m in matches]
    longitudes = [float(m[2]) for m in matches]

    cap = cv2.VideoCapture(path_video_origen)
    fps = int(cap.get(cv2.CAP_PROP_FPS))
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    video_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    video_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

    if total_frames == 0:
        raise ValueError("❌ No se pudo abrir el video origen.")

    nuevo_ancho_video = int(video_width * (NUEVO_ALTO / video_height))
    ancho_mapa = int(nuevo_ancho_video / 3)
    ancho_final_video = nuevo_ancho_video + ancho_mapa
    map_aspect_ratio = ancho_mapa / NUEVO_ALTO

    plt.ioff()
    fig, ax = plt.subplots(figsize=(6 * map_aspect_ratio, 6), dpi=120)
    ax.set_position([0, 0, 1, 1])
    ax.axis('off')
    ax.plot(longitudes, latitudes, color='#00f3ff', linewidth=3, alpha=0.7, zorder=2)
    margen_carga = 0.003
    ax.set_xlim(min(longitudes) - margen_carga, max(longitudes) + margen_carga)
    ax.set_ylim(min(latitudes) - margen_carga, max(latitudes) + margen_carga)

    if opcion_mapa == '2' and RUTA_MAPA_LOCAL:
        try:
            src_raster = rasterio.open(RUTA_MAPA_LOCAL)
            extent = [src_raster.bounds.left, src_raster.bounds.right, src_raster.bounds.bottom, src_raster.bounds.top]
            if src_raster.count >= 3:
                r = src_raster.read(1)
                g = src_raster.read(2)
                b = src_raster.read(3)
                rgb_render = np.dstack((r, g, b))
            else:
                band1 = src_raster.read(1)
                try:
                    cmap = src_raster.colormap(1)
                    r = np.zeros(band1.shape, dtype=np.uint8)
                    g = np.zeros(band1.shape, dtype=np.uint8)
                    b = np.zeros(band1.shape, dtype=np.uint8)
                    for val, rgba in cmap.items():
                        mask = (band1 == val)
                        r[mask] = rgba[0]
                        g[mask] = rgba[1]
                        b[mask] = rgba[2]
                    rgb_render = np.dstack((r, g, b))
                except ValueError:
                    rgb_render = np.dstack((band1, band1, band1))
            ax.imshow(rgb_render, extent=extent, zorder=1)
        except Exception:
            cx.add_basemap(ax, crs='EPSG:4326', source=cx.providers.Esri.WorldImagery, zoom=17, zorder=1)
    elif opcion_mapa == '3' and LINK_ARCGIS:
        try:
            cx.add_basemap(ax, crs='EPSG:4326', source=LINK_ARCGIS, zorder=1)
        except Exception:
            cx.add_basemap(ax, crs='EPSG:4326', source=cx.providers.Esri.WorldImagery, zoom=17, zorder=1)
    else:
        cx.add_basemap(ax, crs='EPSG:4326', source=cx.providers.Esri.WorldImagery, zoom=17, zorder=1)

    drone_dot = ax.scatter([longitudes[0]], [latitudes[0]], color='#ff0055', s=180, edgecolors='white', linewidths=1.5, zorder=3)

    ret, primer_frame = cap.read()
    miniatura_guardada = False

    if ret:
        start_lat, start_lon, start_alt = telemetria.get(1, (latitudes[0], longitudes[0], 0.0))
        ax.set_xlim(start_lon - DISTANCIA_ZOOM, start_lon + DISTANCIA_ZOOM)
        ax.set_ylim(start_lat - (DISTANCIA_ZOOM / map_aspect_ratio), start_lat + (DISTANCIA_ZOOM / map_aspect_ratio))
        fig.canvas.draw()
        mapa_img = np.asarray(fig.canvas.buffer_rgba())
        mapa_img = cv2.cvtColor(mapa_img, cv2.COLOR_RGBA2BGR)
        mapa_img = cv2.resize(mapa_img, (ancho_mapa, NUEVO_ALTO))

        frame_resizado = cv2.resize(primer_frame, (nuevo_ancho_video, NUEVO_ALTO))
        frame_resizado_estampado = estampar_cuadro_video(frame_resizado, start_lat, start_lon, start_alt, 1, fps)
        miniatura_final = np.hstack((frame_resizado_estampado, mapa_img))
        is_success, im_buf_arr = cv2.imencode(".png", miniatura_final)
        if is_success:
            try:
                im_buf_arr.tofile(path_miniatura)
                miniatura_guardada = True
                webbrowser.open(path_miniatura)
            except Exception: pass

    if miniatura_guardada:
        confirmacion = input("\n¿El diseño es correcto? ¿Proceder con el render? (s/n): ").strip().lower()
    else:
        confirmacion = input("¿Deseas intentar el renderizado de todos modos? (s/n): ").strip().lower()

    if confirmacion != 's':
        cap.release()
        plt.close(fig)
        return

    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    fourcc = cv2.VideoWriter_fourcc(*'mp4v')
    out = cv2.VideoWriter(path_video_destino, fourcc, fps, (ancho_final_video, NUEVO_ALTO))

    frame_idx = 1
    mapa_img_cache = None

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret: break

        if frame_idx in telemetria:
            current_lat, current_lon, current_alt = telemetria[frame_idx]
        else:
            current_lat, current_lon, current_alt = latitudes[-1], longitudes[-1], 0.0

        if frame_idx == 1 or frame_idx % SALTAR_CUADROS_MAPA == 0 or mapa_img_cache is None:
            drone_dot.set_offsets([[current_lon, current_lat]])
            ax.set_xlim(current_lon - DISTANCIA_ZOOM, current_lon + DISTANCIA_ZOOM)
            ax.set_ylim(current_lat - (DISTANCIA_ZOOM / map_aspect_ratio), current_lat + (DISTANCIA_ZOOM / map_aspect_ratio))
            fig.canvas.draw()
            mapa_img_raw = np.asarray(fig.canvas.buffer_rgba())
            mapa_img_cv = cv2.cvtColor(mapa_img_raw, cv2.COLOR_RGBA2BGR)
            mapa_img_cache = cv2.resize(mapa_img_cv, (ancho_mapa, NUEVO_ALTO))

        frame_resizado = cv2.resize(frame, (nuevo_ancho_video, NUEVO_ALTO))
        frame_resizado_estampado = estampar_cuadro_video(frame_resizado, current_lat, current_lon, current_alt, frame_idx, fps)
        frame_final = np.hstack((frame_resizado_estampado, mapa_img_cache))
        out.write(frame_final)

        if frame_idx % 150 == 0 or frame_idx == total_frames:
            print(f"    Progreso: {frame_idx}/{total_frames} cuadros ({(frame_idx/total_frames)*100:.1f}%)")

        frame_idx += 1

    cap.release()
    out.release()
    plt.close(fig)
    print(f"\n🎉 ¡Video completado con éxito!\n💾 Guardado en: {path_video_destino}")


if __name__ == "__main__":
    cargar_bd_arcgis(force=True)
    asegurar_reportes_dir()
    ejecutar_menu()