import math
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


NUM_RE = re.compile(r"(\d{6,12})")
PAREN_NUM_RE = re.compile(r"\((\d{6,12})\)\s*$")


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    try:
        if isinstance(value, float) and math.isnan(value):
            return True
    except Exception:
        pass
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def limpiar_texto(value: Any) -> str:
    if _is_empty(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().isoformat(sep=" ")
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    txt = str(value).replace("\xa0", " ").strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def sin_tildes(txt: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", txt or "")
        if unicodedata.category(c) != "Mn"
    )


def normalizar_clave(txt: str) -> str:
    t = sin_tildes(str(txt or "")).upper()
    t = re.sub(r"[^A-Z0-9]+", "_", t).strip("_")
    return t


def obtener(row: Dict[str, Any], *nombres: str) -> Any:
    for n in nombres:
        if n in row:
            return row.get(n)
    # fallback por normalización, para columnas con tildes/espacios raros
    mapa = {normalizar_clave(k): k for k in row.keys()}
    for n in nombres:
        key = normalizar_clave(n)
        if key in mapa:
            return row.get(mapa[key])
    return None


def extraer_numero_y_descripcion(value: Any) -> Tuple[str, str]:
    """
    SAP suele exportar campos así:
      FALLA VÁLVULA CHECK BOMBA (10012648)
    Retorna:
      numero='10012648', descripcion='FALLA VÁLVULA CHECK BOMBA'
    """
    txt = limpiar_texto(value)
    if not txt:
        return "", ""

    # Si el texto termina con número entre paréntesis, separamos ese número.
    m = PAREN_NUM_RE.search(txt)
    if m:
        numero = m.group(1)
        descripcion = PAREN_NUM_RE.sub("", txt).strip()
        return numero, descripcion

    # Si es numérico puro o contiene un número SAP aislado.
    nums = NUM_RE.findall(txt)
    numero = nums[-1] if nums else ""
    descripcion = txt
    if numero and txt.strip() == numero:
        descripcion = ""
    return numero, descripcion


def fecha_a_texto(value: Any) -> Optional[str]:
    if _is_empty(value):
        return None
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return value.to_pydatetime().isoformat(sep=" ")
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return limpiar_texto(value) or None


def normalizar_central(value: Any) -> str:
    txt = sin_tildes(limpiar_texto(value)).upper()
    if "SANTA ROSA" in txt or "PTSR" in txt:
        return "SANTA ROSA"
    if "VENTANILLA" in txt or "PTVE" in txt:
        return "VENTANILLA"
    if "WAYRA" in txt:
        return "WAYRA"
    if "CALLAHUANCA" in txt:
        return "CALLAHUANCA"
    return limpiar_texto(value)


def clasificar_estado(estado_orden: Any, estado_sistema: Any, estado_usuario: Any = None) -> str:
    eo = sin_tildes(limpiar_texto(estado_orden)).upper()
    es = sin_tildes(limpiar_texto(estado_sistema)).upper()
    eu = sin_tildes(limpiar_texto(estado_usuario)).upper()
    combo = f"{eo} {es} {eu}"

    if "BORR" in combo or "MARCADO" in combo:
        return "BORRADO"
    if "CTEC" in combo or "CERR" in combo or "CERRADO TEC" in combo:
        return "CERRADO_TEC"
    if "COMPLETADO" in combo:
        return "COMPLETADO_EMPRESA"
    if "LIBR" in combo or "LIB." in combo or "LIBERADO" in combo:
        return "LIBERADO"
    if "ABIE" in combo or "PENDIENTE" in combo or "ABIERTO" in combo:
        return "PENDIENTE_ABIERTO"
    return "DESCONOCIDO"


def clasificar_tipo_mant(clase_orden: Any, descripcion: Any = None) -> str:
    txt = sin_tildes(f"{limpiar_texto(clase_orden)} {limpiar_texto(descripcion)}").upper()
    if "PROY" in txt or "PROYECT" in txt:
        return "PROY"
    if "CORR" in txt or "CORRECT" in txt or "M2" in txt:
        return "CORR"
    if "PLANIFICADO" in txt or "PREV" in txt or "PREVENT" in txt or "M1" in txt:
        return "PREV"
    return ""


def serializable(value: Any) -> Any:
    if _is_empty(value):
        return None
    if isinstance(value, pd.Timestamp):
        return fecha_a_texto(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, float):
        if math.isnan(value):
            return None
        return value
    if isinstance(value, (str, int, bool)):
        return value
    return limpiar_texto(value)


def parsear_ordenes_sap(ruta_excel: str, archivo_fuente: Optional[str] = None) -> Dict[str, Any]:
    path = Path(ruta_excel)
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo SAP: {ruta_excel}")

    # SAPUI5 suele exportar todo en una hoja llamada Exportación SAPUI5.
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df = df.dropna(how="all")

    registros: List[Dict[str, Any]] = []
    estados: Dict[str, int] = {}
    centrales: Dict[str, int] = {}

    for idx, row_series in df.iterrows():
        row = row_series.to_dict()

        orden_raw = obtener(row, "Orden")
        aviso_raw = obtener(row, "Aviso")
        numero_ot, descripcion_ot = extraer_numero_y_descripcion(orden_raw)
        numero_aviso, descripcion_aviso = extraer_numero_y_descripcion(aviso_raw)

        # Fallbacks útiles si SAP trae descripción limpia en otras columnas.
        if not descripcion_aviso:
            descripcion_aviso = limpiar_texto(obtener(row, "Descripción del aviso"))
        if not descripcion_ot:
            descripcion_ot = limpiar_texto(obtener(row, "Descripción"))

        clase_orden = limpiar_texto(obtener(row, "Clase de orden"))
        centro_puesto = limpiar_texto(obtener(row, "Centro de puesto de trabajo principal"))
        central = normalizar_central(
            centro_puesto
            or obtener(row, "Nombre de centro de puesto de trabajo principal")
            or obtener(row, "Centro de planificación")
            or obtener(row, "Nombre de centro de planificación")
        )

        estado_orden = limpiar_texto(obtener(row, "Estado de la orden"))
        descripcion_estado_orden = limpiar_texto(obtener(row, "Descripción de estado de orden"))
        estado_usuario = limpiar_texto(obtener(row, "Estado de usuario"))
        estado_sistema = limpiar_texto(obtener(row, "Estado del sistema"))
        estado_control = clasificar_estado(estado_orden, estado_sistema, estado_usuario)

        objeto_tecnico = limpiar_texto(obtener(row, "Objeto técnico"))
        descripcion_objeto_tecnico = limpiar_texto(obtener(row, "Descripción del objeto técnico"))
        equipo = limpiar_texto(obtener(row, "Equipo"))
        ubicacion_tecnica = limpiar_texto(obtener(row, "Ubicación técnica"))
        puesto = limpiar_texto(obtener(row, "Puesto de trabajo principal"))
        prioridad = limpiar_texto(obtener(row, "Prioridad", "Texto prioridad"))

        registro = {
            "numero_ot": numero_ot or None,
            "descripcion_ot": descripcion_ot or None,
            "numero_aviso": numero_aviso or None,
            "descripcion_aviso": descripcion_aviso or None,
            "clase_orden": clase_orden or None,
            "tipo_mant_sap": clasificar_tipo_mant(clase_orden, descripcion_ot) or None,
            "central": central or None,
            "centro_puesto": centro_puesto or None,
            "objeto_tecnico": objeto_tecnico or None,
            "descripcion_objeto_tecnico": descripcion_objeto_tecnico or None,
            "equipo": equipo or None,
            "ubicacion_tecnica": ubicacion_tecnica or None,
            "estado_orden": estado_orden or None,
            "descripcion_estado_orden": descripcion_estado_orden or None,
            "estado_usuario": estado_usuario or None,
            "estado_sistema": estado_sistema or None,
            "estado_control": estado_control,
            "inicio_planificado": fecha_a_texto(obtener(row, "Inicio planificado (América, Lima)", "Inicio planificado")),
            "fin_programado": fecha_a_texto(obtener(row, "Fin programado (América, Lima)", "Fin programado")),
            "fecha_creacion": fecha_a_texto(obtener(row, "Fecha y hora de creación (América, Lima)", "Fecha creación")),
            "puesto_trabajo_principal": puesto or None,
            "prioridad": prioridad or None,
            "archivo_fuente": archivo_fuente or path.name,
            "fila_sap": int(idx) + 2,  # +2 por encabezado Excel
            "raw_data": {str(k): serializable(v) for k, v in row.items()},
        }

        # Saltar filas sin OT ni aviso: no sirven para el maestro.
        if not registro["numero_ot"] and not registro["numero_aviso"]:
            continue

        registros.append(registro)
        estados[estado_control] = estados.get(estado_control, 0) + 1
        central_key = central or "SIN CENTRAL"
        centrales[central_key] = centrales.get(central_key, 0) + 1

    ots_unicas = len({r["numero_ot"] for r in registros if r.get("numero_ot")})
    avisos_unicos = len({r["numero_aviso"] for r in registros if r.get("numero_aviso")})

    return {
        "ok": True,
        "archivo_fuente": archivo_fuente or path.name,
        "filas_excel": int(len(df)),
        "registros": registros,
        "resumen": {
            "filas_procesadas": int(len(df)),
            "registros_validos": len(registros),
            "ots_extraidas": ots_unicas,
            "avisos_extraidos": avisos_unicos,
            "estados": estados,
            "centrales": centrales,
        },
    }
