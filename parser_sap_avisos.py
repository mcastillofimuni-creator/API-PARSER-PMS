import math
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


# ============================================================
# Parser inicial SAP Avisos / Notificaciones
# ============================================================
# Objetivo:
# - Leer el Excel de avisos tal como sale de SAP.
# - Extraer numero_aviso, descripcion_aviso, OT asociada si existe,
#   estado, central, equipo, ubicación técnica y raw_data.
# - No requiere Supabase. Devuelve un diccionario con registros y resumen.
#
# Uso esperado desde main.py:
#   from parser_sap_avisos import parsear_avisos_sap
#   resultado = parsear_avisos_sap(ruta_excel, archivo_fuente="Avisos.xlsx")
#
# Salida principal:
#   resultado["registros"] -> lista de dicts listos para cruzar en memoria
# ============================================================

NUM_RE = re.compile(r"(\d{6,12})")
PAREN_NUM_RE = re.compile(r"\((\d{6,12})\)\s*$")
PLAN_RE = re.compile(r"\b([A-Z]{2,8}\d[A-Z0-9]{3,20})\b", re.IGNORECASE)


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
    if txt.upper() in {"NAN", "NONE", "NULL"}:
        return ""
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
    """Obtiene una columna por nombre exacto o por nombre normalizado."""
    for n in nombres:
        if n in row:
            return row.get(n)

    mapa = {normalizar_clave(k): k for k in row.keys()}
    for n in nombres:
        key = normalizar_clave(n)
        if key in mapa:
            return row.get(mapa[key])
    return None


def buscar_por_alias(row: Dict[str, Any], aliases: List[str]) -> Any:
    return obtener(row, *aliases)


def extraer_numero_y_descripcion(value: Any) -> Tuple[str, str]:
    """
    SAP suele exportar campos así:
      FALLA BOMBA AGUA CRUDA (10002122)
    Retorna:
      numero='10002122', descripcion='FALLA BOMBA AGUA CRUDA'
    """
    txt = limpiar_texto(value)
    if not txt:
        return "", ""

    m = PAREN_NUM_RE.search(txt)
    if m:
        numero = m.group(1)
        descripcion = PAREN_NUM_RE.sub("", txt).strip()
        return numero, descripcion

    nums = NUM_RE.findall(txt)
    numero = nums[-1] if nums else ""
    descripcion = txt
    if numero and txt.strip() == numero:
        descripcion = ""
    return numero, descripcion


def extraer_primer_numero(value: Any) -> str:
    txt = limpiar_texto(value)
    nums = NUM_RE.findall(txt)
    return nums[0] if nums else ""


def extraer_plan_pm_de_texto(*values: Any) -> str:
    """
    Detecta planes/códigos PM tipo:
    - PTSRW7ME0349
    - PTSRW7EL0001
    - PW123 / PM123, si existieran
    """
    texto = " ".join(limpiar_texto(v) for v in values if not _is_empty(v)).upper()
    if not texto:
        return ""

    # Patrones largos tipo PTSRW7ME0349
    candidatos = PLAN_RE.findall(texto)
    for c in candidatos:
        c = c.upper().strip()
        # Evitar tomar palabras comunes como si fueran plan.
        if any(c.startswith(pref) for pref in ["PTSR", "PTVE", "PW", "PM", "MPLAN", "PLAN"]):
            return c

    return ""


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
    if "SANTA ROSA" in txt or "PTSR" in txt or "CTSR" in txt:
        return "SANTA ROSA"
    if "VENTANILLA" in txt or "PTVE" in txt:
        return "VENTANILLA"
    if "WAYRA" in txt:
        return "WAYRA"
    if "CALLAHUANCA" in txt:
        return "CALLAHUANCA"
    return limpiar_texto(value)


def clasificar_estado_aviso(estado_aviso: Any, estado_sistema: Any = None, estado_usuario: Any = None) -> str:
    ea = sin_tildes(limpiar_texto(estado_aviso)).upper()
    es = sin_tildes(limpiar_texto(estado_sistema)).upper()
    eu = sin_tildes(limpiar_texto(estado_usuario)).upper()
    combo = f"{ea} {es} {eu}"

    if "BORR" in combo or "MARCADO" in combo:
        return "BORRADO"
    if "CERR" in combo or "CTEC" in combo or "CERRADO" in combo:
        return "CERRADO"
    if "CONCL" in combo or "COMPLET" in combo:
        return "CONCLUIDO"
    if "LIBR" in combo or "LIBERADO" in combo:
        return "LIBERADO"
    if "ABIE" in combo or "ABIERTO" in combo or "PEND" in combo or "NUEVO" in combo:
        return "ABIERTO"
    if "TRAT" in combo or "PROCES" in combo or "EN CURSO" in combo:
        return "EN_TRATAMIENTO"
    return "DESCONOCIDO"


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


# Aliases flexibles para diferentes exportaciones SAP
ALIAS_AVISO = [
    "Aviso",
    "Notificación",
    "Notificacion",
    "Número de aviso",
    "Numero de aviso",
    "Nº aviso",
    "N° aviso",
    "Nro aviso",
    "Aviso de mantenimiento",
]

ALIAS_DESCRIPCION_AVISO = [
    "Descripción del aviso",
    "Descripcion del aviso",
    "Descripción",
    "Descripcion",
    "Texto breve",
    "Texto breve aviso",
    "Denominación",
    "Denominacion",
]

ALIAS_OT_ASOCIADA = [
    "Orden",
    "Orden asociada",
    "Orden de mantenimiento",
    "Orden PM",
    "OT",
    "N° OT",
    "Nº OT",
    "Nro OT",
    "Orden de trabajo",
]

ALIAS_ESTADO_AVISO = [
    "Estado del aviso",
    "Estado de aviso",
    "Estado aviso",
    "Status aviso",
    "Estado",
]

ALIAS_ESTADO_USUARIO = [
    "Estado de usuario",
    "Status usuario",
    "Usuario status",
]

ALIAS_ESTADO_SISTEMA = [
    "Estado del sistema",
    "Estado sistema",
    "Status sistema",
    "Sistema status",
]

ALIAS_CLASE_AVISO = [
    "Clase de aviso",
    "Clase aviso",
    "Tipo de aviso",
    "Tipo aviso",
]

ALIAS_PRIORIDAD = [
    "Prioridad",
    "Texto prioridad",
    "Prioridad aviso",
]

ALIAS_CENTRO = [
    "Centro de puesto de trabajo principal",
    "Centro puesto trabajo principal",
    "Centro de planificación",
    "Centro de planificacion",
    "Nombre de centro de planificación",
    "Nombre de centro de planificacion",
    "Centro",
    "Planta",
]

ALIAS_OBJETO = [
    "Objeto técnico",
    "Objeto tecnico",
    "Ubicación técnica",
    "Ubicacion tecnica",
    "Equipo",
]

ALIAS_DESC_OBJETO = [
    "Descripción del objeto técnico",
    "Descripcion del objeto tecnico",
    "Descripción objeto técnico",
    "Descripcion objeto tecnico",
]

ALIAS_EQUIPO = [
    "Equipo",
    "Nº equipo",
    "N° equipo",
    "Número de equipo",
    "Numero de equipo",
]

ALIAS_UBICACION = [
    "Ubicación técnica",
    "Ubicacion tecnica",
    "Ubic. técnica",
    "Ubicacion",
]

ALIAS_FECHA_AVISO = [
    "Fecha de aviso",
    "Fecha aviso",
    "Fecha de notificación",
    "Fecha de notificacion",
    "Fecha",
]

ALIAS_FECHA_CREACION = [
    "Fecha y hora de creación (América, Lima)",
    "Fecha y hora de creacion (America, Lima)",
    "Fecha de creación",
    "Fecha de creacion",
    "Creado el",
]

ALIAS_INICIO_DESEADO = [
    "Inicio deseado",
    "Inicio requerido",
    "Inicio planificado",
]

ALIAS_FIN_DESEADO = [
    "Fin deseado",
    "Fin requerido",
    "Fin programado",
]

ALIAS_PUESTO = [
    "Puesto de trabajo principal",
    "Puesto trabajo principal",
    "Puesto de trabajo",
]

ALIAS_RESPONSABLE = [
    "Responsable",
    "Autor",
    "Creado por",
    "Notificador",
]

ALIAS_PLAN_PM = [
    "Plan de mantenimiento",
    "Plan mantenimiento",
    "Plan PM",
    "Código plan",
    "Codigo plan",
    "COD PM",
    "Código PM",
    "Codigo PM",
]


def elegir_hoja_excel(path: Path) -> str:
    """
    Elige la primera hoja con más filas útiles.
    Evita depender de que se llame exactamente 'Exportación SAPUI5'.
    """
    xls = pd.ExcelFile(path)
    mejor_hoja = xls.sheet_names[0]
    mejor_score = -1

    for hoja in xls.sheet_names:
        try:
            df_tmp = pd.read_excel(path, sheet_name=hoja, dtype=object, nrows=30)
            score = int(df_tmp.dropna(how="all").shape[0]) * max(1, int(df_tmp.shape[1]))
            if score > mejor_score:
                mejor_score = score
                mejor_hoja = hoja
        except Exception:
            continue

    return mejor_hoja


def parsear_avisos_sap(ruta_excel: str, archivo_fuente: Optional[str] = None) -> Dict[str, Any]:
    path = Path(ruta_excel)
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo SAP de avisos: {ruta_excel}")

    hoja = elegir_hoja_excel(path)
    df = pd.read_excel(path, sheet_name=hoja, dtype=object)
    df = df.dropna(how="all")

    registros: List[Dict[str, Any]] = []
    estados: Dict[str, int] = {}
    centrales: Dict[str, int] = {}

    for idx, row_series in df.iterrows():
        row = row_series.to_dict()

        aviso_raw = buscar_por_alias(row, ALIAS_AVISO)
        numero_aviso, descripcion_desde_aviso = extraer_numero_y_descripcion(aviso_raw)

        descripcion_aviso = limpiar_texto(buscar_por_alias(row, ALIAS_DESCRIPCION_AVISO))
        if not descripcion_aviso:
            descripcion_aviso = descripcion_desde_aviso

        ot_raw = buscar_por_alias(row, ALIAS_OT_ASOCIADA)
        numero_ot_asociada, descripcion_ot_asociada = extraer_numero_y_descripcion(ot_raw)

        # Si el aviso vino como texto puro en descripción y el campo Aviso solo contiene número.
        if not numero_aviso:
            numero_aviso = extraer_primer_numero(aviso_raw)

        clase_aviso = limpiar_texto(buscar_por_alias(row, ALIAS_CLASE_AVISO))
        tipo_aviso = clase_aviso
        prioridad = limpiar_texto(buscar_por_alias(row, ALIAS_PRIORIDAD))

        centro_puesto = limpiar_texto(buscar_por_alias(row, ALIAS_CENTRO))
        objeto_tecnico = limpiar_texto(buscar_por_alias(row, ALIAS_OBJETO))
        descripcion_objeto_tecnico = limpiar_texto(buscar_por_alias(row, ALIAS_DESC_OBJETO))
        equipo = limpiar_texto(buscar_por_alias(row, ALIAS_EQUIPO))
        ubicacion_tecnica = limpiar_texto(buscar_por_alias(row, ALIAS_UBICACION))

        central = normalizar_central(
            centro_puesto
            or objeto_tecnico
            or ubicacion_tecnica
            or descripcion_objeto_tecnico
        )

        estado_aviso = limpiar_texto(buscar_por_alias(row, ALIAS_ESTADO_AVISO))
        estado_usuario = limpiar_texto(buscar_por_alias(row, ALIAS_ESTADO_USUARIO))
        estado_sistema = limpiar_texto(buscar_por_alias(row, ALIAS_ESTADO_SISTEMA))
        estado_control = clasificar_estado_aviso(estado_aviso, estado_sistema, estado_usuario)

        plan_pm = limpiar_texto(buscar_por_alias(row, ALIAS_PLAN_PM))
        if not plan_pm:
            plan_pm = extraer_plan_pm_de_texto(
                descripcion_aviso,
                descripcion_ot_asociada,
                objeto_tecnico,
                descripcion_objeto_tecnico,
                *[v for v in row.values() if isinstance(v, str)],
            )

        registro = {
            "numero_aviso": numero_aviso or None,
            "descripcion_aviso": descripcion_aviso or None,
            "numero_ot_asociada": numero_ot_asociada or None,
            "descripcion_ot_asociada": descripcion_ot_asociada or None,
            "clase_aviso": clase_aviso or None,
            "tipo_aviso": tipo_aviso or None,
            "prioridad": prioridad or None,
            "central": central or None,
            "centro_puesto": centro_puesto or None,
            "objeto_tecnico": objeto_tecnico or None,
            "descripcion_objeto_tecnico": descripcion_objeto_tecnico or None,
            "equipo": equipo or None,
            "ubicacion_tecnica": ubicacion_tecnica or None,
            "estado_aviso": estado_aviso or None,
            "estado_usuario": estado_usuario or None,
            "estado_sistema": estado_sistema or None,
            "estado_control": estado_control,
            "fecha_aviso": fecha_a_texto(buscar_por_alias(row, ALIAS_FECHA_AVISO)),
            "fecha_creacion": fecha_a_texto(buscar_por_alias(row, ALIAS_FECHA_CREACION)),
            "inicio_deseado": fecha_a_texto(buscar_por_alias(row, ALIAS_INICIO_DESEADO)),
            "fin_deseado": fecha_a_texto(buscar_por_alias(row, ALIAS_FIN_DESEADO)),
            "puesto_trabajo_principal": limpiar_texto(buscar_por_alias(row, ALIAS_PUESTO)) or None,
            "responsable": limpiar_texto(buscar_por_alias(row, ALIAS_RESPONSABLE)) or None,
            "plan_pm": plan_pm or None,
            "archivo_fuente": archivo_fuente or path.name,
            "hoja_fuente": hoja,
            "fila_sap": int(idx) + 2,
            "raw_data": {str(k): serializable(v) for k, v in row.items()},
        }

        # Saltar filas sin aviso y sin OT asociada.
        if not registro["numero_aviso"] and not registro["numero_ot_asociada"]:
            continue

        registros.append(registro)
        estados[estado_control] = estados.get(estado_control, 0) + 1
        central_key = central or "SIN CENTRAL"
        centrales[central_key] = centrales.get(central_key, 0) + 1

    avisos_unicos = len({r["numero_aviso"] for r in registros if r.get("numero_aviso")})
    ots_asociadas_unicas = len({r["numero_ot_asociada"] for r in registros if r.get("numero_ot_asociada")})
    planes_unicos = len({r["plan_pm"] for r in registros if r.get("plan_pm")})

    return {
        "ok": True,
        "archivo_fuente": archivo_fuente or path.name,
        "hoja_fuente": hoja,
        "filas_excel": int(len(df)),
        "registros": registros,
        "resumen": {
            "filas_procesadas": int(len(df)),
            "registros_validos": len(registros),
            "avisos_extraidos": avisos_unicos,
            "ots_asociadas_extraidas": ots_asociadas_unicas,
            "planes_extraidos": planes_unicos,
            "estados": estados,
            "centrales": centrales,
        },
    }


# Prueba local opcional:
# python parser_sap_avisos.py archivo.xlsx
if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) < 2:
        print("Uso: python parser_sap_avisos.py <archivo_avisos.xlsx>")
        raise SystemExit(1)

    res = parsear_avisos_sap(sys.argv[1])
    print(json.dumps({k: v for k, v in res.items() if k != "registros"}, ensure_ascii=False, indent=2))
    print("Primeros registros:")
    print(json.dumps(res["registros"][:5], ensure_ascii=False, indent=2))
