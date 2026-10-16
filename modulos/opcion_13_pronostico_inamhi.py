import os
import re
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import requests
from requests.auth import HTTPBasicAuth
import numpy as np
import rasterio
from rasterio.warp import reproject, Resampling, calculate_default_transform
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize
import matplotlib.patheffects as path_effects
import matplotlib.image as mpimg
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
import geopandas as gpd
from moviepy import ImageSequenceClip
from modulos.comun import ruta_en_reportes, RUTA_LOGO_PREFECTURA

USUARIO = os.getenv("INAMHI_USER", "ExternoINAMHI")
CLAVE = os.getenv("INAMHI_PASS", "externoinamhi2025*")
BASE_URL = "https://infohm.inamhi.gob.ec/remote.php/dav/files/ExternoINAMHI"

ruta_mascara = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\MascaraBlancaFueraAzuay2.gpkg"
ruta_shp_cantones = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\AzuayXCantones.gpkg"
ruta_extension = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\ExtensionAzuay.gpkg"
ruta_poblados = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\PobladosPrincipalesAzuay2.gpkg"
ruta_logo_1 = RUTA_LOGO_PREFECTURA
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


def ejecutar():
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