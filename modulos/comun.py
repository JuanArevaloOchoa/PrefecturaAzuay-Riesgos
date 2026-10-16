import os
import builtins
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd
import urllib3
from arcgis.gis import GIS
from dotenv import load_dotenv

load_dotenv()
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuración ArcGIS
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

# Rutas estándar compartidas
RUTA_BASE_CARPETAS = r"D:\OneDrive - GAD Provincial del Azuay"
RUTA_EXCEL_TRAMITES = r"D:\OneDrive - GAD Provincial del Azuay\RegistroTrámitesJuan.xlsx"
RUTA_GPKG_PARROQUIAS = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\AzuayXParroquias.gpkg"
RUTA_LOGO_PREFECTURA = r"D:\OneDrive - GAD Provincial del Azuay\0.1 Recursos\logo prefectura.png"


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


# Gestión de salida con 's'
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


def activar_input_personalizado():
    builtins.input = _input_con_salida