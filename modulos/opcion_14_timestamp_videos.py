import os
import re
import webbrowser
import cv2
import numpy as np
import rasterio
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
from pyproj import Transformer
import geopandas as gpd
from shapely.geometry import Point
import contextily as cx
from modulos.comun import RUTA_GPKG_PARROQUIAS, RUTA_LOGO_PREFECTURA


def ejecutar():
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
            if not LINK_ARCGIS.endswith("/"):
                LINK_ARCGIS += "/"
            LINK_ARCGIS += "tile/{z}/{y}/{x}"

    DISTANCIA_ZOOM = 0.0015
    NUEVO_ALTO = 720
    SALTAR_CUADROS_MAPA = 3

    PATH_SHAPE = RUTA_GPKG_PARROQUIAS
    PATH_LOGO = RUTA_LOGO_PREFECTURA

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
            except Exception:
                pass

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
        if not ret:
            break

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