import pandas as pd
from pyproj import Transformer
from modulos.comun import cargar_bd_arcgis

def ejecutar():
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