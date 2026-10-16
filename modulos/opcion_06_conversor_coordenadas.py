from pyproj import Transformer

def ejecutar():
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