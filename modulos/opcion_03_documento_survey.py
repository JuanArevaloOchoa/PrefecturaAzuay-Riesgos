import os
import shutil
from datetime import datetime
from pathlib import Path
import openpyxl
import pandas as pd
import numpy as np
from docx import Document
from pyproj import Transformer
from modulos.comun import cargar_bd_arcgis, RUTA_EXCEL_TRAMITES

def ejecutar():
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

    datos_punto = {}
    for col, val in punto_sel.items():
        nombre_var = col.replace(' ', '_').lower()
        datos_punto[nombre_var] = val.strip() if isinstance(val, str) else val

    ruta_excel = RUTA_EXCEL_TRAMITES
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
            COL_INFORME: limpiar_para_excel(datos_punto.get('no_informe', '')),
            COL_NOMBRE: limpiar_para_excel(datos_punto.get('nombre_del_punto', '')),
            COL_FECHA: limpiar_para_excel(datos_punto.get('fecha_de_visita', '')),
            COL_DESC: limpiar_para_excel(datos_punto.get('descripci_n_del_evento', '')),
            COL_TECNICO: limpiar_para_excel(datos_punto.get('comentarios_t_cnicos', '')),
            COL_VIAL: limpiar_para_excel(datos_punto.get('inventario_vial', ''))
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
    CoordenadaX, CoordenadaY = transformer.transform(datos_punto.get('longitud', 0), datos_punto.get('latitud', 0))

    meses = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
             7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
    hoy = datetime.now()
    fecha_formateada = f"{hoy.day} DE {meses[hoy.month]} DEL {hoy.year}"

    f_visita = datos_punto.get('fecha_de_visita')
    if hasattr(f_visita, 'strftime'):
        solo_fecha_VT = f"{f_visita.day} de {f_visita.strftime('%B')} del {f_visita.year}"
    else:
        solo_fecha_VT = str(f_visita) if pd.notna(f_visita) else ""

    TecnicoGPA = {"Juan_Arévalo": "Ing. Juan Arévalo Ochoa, Responsable de Riesgos",
                  "Xavier_Sánchez": "Ing. Xavier Sánchez, Técnico de Riesgos"}.get(datos_punto.get('field_50'), "Null")

    ruta_plantilla = r"D:\OneDrive - GAD Provincial del Azuay\1000. Nuevo informe y GIS\PLANTILLA INFORME.docx"
    informe_nombre = datos_punto.get('no_informe', 'INFORME_SIN_NUMERO')
    ruta_salida_word = os.path.join(str(rutafinal1), f"{informe_nombre}.docx")

    temp_series = punto_sel.replace(to_replace='<NA>', value=np.nan)
    filtered_series = temp_series.dropna()
    resultado_final_str = ""
    for key, value in filtered_series[6:-1].items():
        resultado_final_str += f"{key}: {value}\n"

    sustituciones = {
        "#AUTOR#": "Juan Arévalo Ochoa",
        "#FECHAINF#": fecha_formateada.upper(),
        "#TIPOEVENTO#": datos_punto.get('tipo_de_evento', ''),
        "#SUBTIPOEVENTO#": datos_punto.get('subtipo_de_evento', ''),
        "#FECHAVT#": solo_fecha_VT,
        "#COORDX#": int(CoordenadaX),
        "#COORDY#": int(CoordenadaY),
        "#SECTOR#": str(datos_punto.get('sector', '')).title(),
        "#PARROQUIA#": datos_punto.get('parroquia', ''),
        "#CANTÓN#": datos_punto.get('cant_n', ''),
        "#ESTADO#": "",
        "#AREAAFECTADA#": "",
        "#TÉCNICOSGADPA#": TecnicoGPA,
        "#PERSONALGADCANTONAL#": datos_punto.get('personal_gad_cantonal', ''),
        "#PERSONALGADPARROQUIAL#": datos_punto.get('personal_gad_parroquial', ''),
        "#NOTRAMITE#": datos_punto.get('no_tr_mite', ''),
        "#NOINFORME#": informe_nombre,
        "#NOMBREPUNTO#": datos_punto.get('nombre_del_punto', ''),
        "#DESCRIPCION#": datos_punto.get('descripci_n_del_evento', ''),
        "#COMENTARIOSTECNICOS#": datos_punto.get('comentarios_t_cnicos', ''),
        "#FECHAEVENTO#": datos_punto.get('fecha_y_hora_de_evento', ''),
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

    if pd.notna(datos_punto.get('no_informe')):
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