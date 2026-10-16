from modulos.comun import (
    cargar_bd_arcgis,
    asegurar_reportes_dir,
    activar_input_personalizado,
    SalirInput,
)
from modulos import (
    opcion_01_recargar_arcgis,
    opcion_02_nuevo_tramite,
    opcion_03_documento_survey,
    opcion_04_timestamp_cuadrado,
    opcion_05_timestamp_rectangular,
    opcion_06_conversor_coordenadas,
    opcion_07_extractor_coordenadas,
    opcion_08_reporte_horas,
    opcion_09_reporte_emergencia,
    opcion_10_descargar_bd_arcgis,
    opcion_11_ayuda_memoria,
    opcion_12_tablas_resumen,
    opcion_13_pronostico_inamhi,
    opcion_14_timestamp_videos,
)

# Activar intercepción global de 's' para volver al menú o salir
activar_input_personalizado()

ACCIONES = {
    "1": opcion_01_recargar_arcgis.ejecutar,
    "2": opcion_02_nuevo_tramite.ejecutar,
    "3": opcion_03_documento_survey.ejecutar,
    "4": opcion_04_timestamp_cuadrado.ejecutar,
    "5": opcion_05_timestamp_rectangular.ejecutar,
    "6": opcion_06_conversor_coordenadas.ejecutar,
    "7": opcion_07_extractor_coordenadas.ejecutar,
    "8": opcion_08_reporte_horas.ejecutar,
    "9": opcion_09_reporte_emergencia.ejecutar,
    "10": opcion_10_descargar_bd_arcgis.ejecutar,
    "11": opcion_11_ayuda_memoria.ejecutar,
    "12": opcion_12_tablas_resumen.ejecutar,
    "13": opcion_13_pronostico_inamhi.ejecutar,
    "14": opcion_14_timestamp_videos.ejecutar,
}


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

        if seleccion == "0":
            print("Saliendo del sistema... ¡Hasta luego!")
            break

        accion = ACCIONES.get(seleccion)
        if accion:
            try:
                accion()
            except SalirInput:
                print("\nOperación cancelada. Volviendo al menú...")
            except Exception as e:
                print(f"\n❌ Error al ejecutar la opción {seleccion}: {e}")
        else:
            print("\n[!] Opción no válida. Intente de nuevo.")


if __name__ == "__main__":
    cargar_bd_arcgis(force=True)
    asegurar_reportes_dir()
    ejecutar_menu()