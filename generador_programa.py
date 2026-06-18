import os
import re
import json
import tempfile
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


PLANTILLA_NOMBRE = "PROGRAMA SEMANAL PLANTILLA.xlsx"

CENTRALES_VALIDAS = {
    "C. T. SANTA ROSA": "SANTA ROSA",
    "CT SANTA ROSA": "SANTA ROSA",
    "C.T. SANTA ROSA": "SANTA ROSA",
    "C T SANTA ROSA": "SANTA ROSA",
    "SANTA ROSA": "SANTA ROSA",
    "STA ROSA": "SANTA ROSA",

    "C.C. VENTANILLA": "VENTANILLA",
    "CC VENTANILLA": "VENTANILLA",
    "C.C VENTANILLA": "VENTANILLA",
    "C C VENTANILLA": "VENTANILLA",
    "VENTANILLA": "VENTANILLA",
}

CENTRAL_LABEL = {
    "SANTA ROSA": "C. T. Santa Rosa",
    "VENTANILLA": "C.C. Ventanilla",
}

MESES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

DIAS_CORTOS = ["SÁB", "DOM", "LUN", "MAR", "MIÉ", "JUE", "VIE"]

# Rango de columnas que se copiará desde datos_originales.columnas_excel.
# Según tu plantilla, el cuerpo relevante empieza en C y llega hasta AU.
COL_INICIO_DATOS = "C"
COL_FIN_DATOS = "AU"

FILA_INICIO_DATOS = 10
FILA_MODELO_ESTILO = 10

# Columnas fallback si datos_originales no trae alguna columna.
COLS_FALLBACK = {
    "ot_grafo": "C",
    "central": "D",
    "unidad": "E",
    "sistema": "F",
    "equipo": "G",
    "cod_pm_aviso": "H",
    "pedido": "I",
    "tipo_mant": "K",
    "condicion": "L",
    "riesgo": "M",
    "area_solicitante": "N",
    "inspector": "O",
    "rt_terceros": "P",
    "recursos": "Q",
    "hora_inicio": "R",
    "hora_fin": "S",
    "empresa": "AD",
    "actividad": "AF",
}


def normalizar_texto(valor):
    if valor is None:
        return ""

    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_central(valor):
    texto = normalizar_texto(valor).upper()
    texto = texto.replace(".", "")
    texto = re.sub(r"\s+", " ", texto)

    if "SANTA ROSA" in texto:
        return "SANTA ROSA"

    if "VENTANILLA" in texto:
        return "VENTANILLA"

    return CENTRALES_VALIDAS.get(texto, texto)


def parse_semana_inicio(semana):
    """
    Espera semana tipo: 2026-06-13.
    Ese día representa sábado.
    """
    try:
        return datetime.strptime(str(semana), "%Y-%m-%d").date()
    except Exception:
        raise ValueError("La semana debe tener formato YYYY-MM-DD, por ejemplo 2026-06-13.")


def fecha_larga(fecha):
    return f"{fecha.day:02d} DE {MESES[fecha.month]} DEL {fecha.year}"


def fecha_corta(fecha):
    return f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year}"


def convertir_valor_excel(valor):
    """
    Limpia valores antes de escribirlos en Excel.
    Mantiene números y fechas cuando sea posible.
    """
    if valor is None:
        return None

    if isinstance(valor, str):
        v = valor.strip()
        if v.upper() in ["", "NULL", "NONE", "NAN"]:
            return None
        return valor

    return valor


def copiar_estilo_celda(origen, destino):
    if origen.has_style:
        destino.font = copy(origen.font)
        destino.fill = copy(origen.fill)
        destino.border = copy(origen.border)
        destino.alignment = copy(origen.alignment)
        destino.number_format = origen.number_format
        destino.protection = copy(origen.protection)

    if origen.hyperlink:
        destino._hyperlink = copy(origen.hyperlink)

    if origen.comment:
        destino.comment = copy(origen.comment)


def copiar_estilo_fila(ws, fila_origen, fila_destino, max_col):
    ws.row_dimensions[fila_destino].height = ws.row_dimensions[fila_origen].height

    for col in range(1, max_col + 1):
        copiar_estilo_celda(ws.cell(fila_origen, col), ws.cell(fila_destino, col))


def limpiar_area_datos(ws, fila_inicio=FILA_INICIO_DATOS):
    """
    Limpia valores del cuerpo de datos sin destruir formato.
    """
    for row in ws.iter_rows(min_row=fila_inicio, max_row=ws.max_row):
        for cell in row:
            cell.value = None


def obtener_valor(row, *keys, default=""):
    for key in keys:
        if key in row and row.get(key) not in [None, "", [], {}]:
            return row.get(key)
    return default


def parse_jsonb(valor):
    """
    Supabase puede devolver jsonb como dict o string.
    """
    if valor is None:
        return {}

    if isinstance(valor, dict):
        return valor

    if isinstance(valor, str):
        try:
            return json.loads(valor)
        except Exception:
            return {}

    return {}


def obtener_columnas_originales(actividad):
    """
    Lee datos_originales.columnas_excel.
    Estructura esperada:
    {
      "hoja": "...",
      "fila_excel": 10,
      "columnas_excel": {
        "A": "...",
        "B": "...",
        "C": "...",
        ...
      },
      "campos_detectados": {...}
    }
    """
    datos = parse_jsonb(actividad.get("datos_originales"))

    columnas = datos.get("columnas_excel")
    if isinstance(columnas, dict):
        return columnas

    return {}


def obtener_campos_detectados(actividad):
    datos = parse_jsonb(actividad.get("datos_originales"))

    campos = datos.get("campos_detectados")
    if isinstance(campos, dict):
        return campos

    return {}


def valor_original_o_fallback(actividad, columnas_originales, letra_columna, *fallback_keys):
    """
    Primero intenta usar la columna original del Excel.
    Si no existe, usa los campos planos de pms_actividades.
    """
    if letra_columna in columnas_originales:
        valor = columnas_originales.get(letra_columna)
        if valor not in [None, "", "NULL", "None", "nan"]:
            return convertir_valor_excel(valor)

    for key in fallback_keys:
        valor = actividad.get(key)
        if valor not in [None, "", "NULL", "None", "nan", [], {}]:
            return convertir_valor_excel(valor)

    return None


def actualizar_encabezado_semana(ws, semana_inicio):
    """
    Actualiza de forma robusta textos de fechas del encabezado.
    No depende de una sola celda exacta.
    """
    semana_fin = semana_inicio + timedelta(days=6)

    titulo = f"PROGRAMA SEMANAL DEL {fecha_larga(semana_inicio)} AL {fecha_larga(semana_fin)}"
    semana_texto = f"SEMANA DEL {fecha_corta(semana_inicio)} AL {fecha_corta(semana_fin)}"

    for row in ws.iter_rows(min_row=1, max_row=9):
        for cell in row:
            if isinstance(cell.value, str):
                texto = cell.value.upper()

                if "PROGRAMA SEMANAL" in texto and ("DEL" in texto or "SEMANA" in texto):
                    cell.value = titulo

                if "SEMANA" in texto and ("AL" in texto or "DEL" in texto):
                    cell.value = semana_texto

    # Buscar encabezados de días en el rango W:AC, pero sin asumir demasiado.
    # Si la plantilla tiene esos campos, los actualizamos.
    day_cols = ["W", "X", "Y", "Z", "AA", "AB", "AC"]

    for i, col in enumerate(day_cols):
        fecha = semana_inicio + timedelta(days=i)

        try:
            ws[f"{col}8"] = fecha
            ws[f"{col}8"].number_format = "dd/mm/yyyy"
        except Exception:
            pass

        try:
            ws[f"{col}9"] = f"{DIAS_CORTOS[i]}\n{fecha.day:02d}/{fecha.month:02d}"
        except Exception:
            pass


def consultar_archivos_semana(supabase, semana):
    resp = (
        supabase.table("pms_archivos")
        .select(
            "id,semana,proveedor,expositor,central_presentada,central_presentada_norm,"
            "archivo_nombre,estado_validacion,errores,advertencias,actividades,"
            "observaciones,dias"
        )
        .eq("semana", semana)
        .execute()
    )

    data = resp.data or []
    return {x["id"]: x for x in data if x.get("id")}


def consultar_actividades(supabase, semana, central_norm):
    """
    Consulta actividades de la semana y central.
    Se filtra por pms_actividades.central para consolidar únicamente
    Santa Rosa o Ventanilla, aunque el archivo original tenga más hojas.
    """
    resp = (
        supabase.table("pms_actividades")
        .select("*")
        .eq("semana", semana)
        .eq("central", central_norm)
        .order("proveedor")
        .order("fila_excel")
        .limit(5000)
        .execute()
    )

    return resp.data or []


def seleccionar_hoja(wb, central_norm):
    """
    Intenta seleccionar una hoja con nombre de la central.
    Si no existe, usa la hoja activa.
    """
    posibles = []

    if central_norm == "SANTA ROSA":
        posibles = ["SANTA ROSA", "STA ROSA", "CT SANTA ROSA", "C.T. SANTA ROSA"]

    elif central_norm == "VENTANILLA":
        posibles = ["VENTANILLA", "CC VENTANILLA", "C.C. VENTANILLA"]

    nombres = {ws.title.upper().strip(): ws.title for ws in wb.worksheets}

    for p in posibles:
        if p.upper() in nombres:
            return wb[nombres[p.upper()]]

    for ws in wb.worksheets:
        nombre = ws.title.upper()
        if central_norm in nombre:
            return ws

    return wb.active


def dejar_solo_hoja_objetivo(wb, ws_objetivo, central_norm):
    """
    Deja una sola hoja para evitar duplicidades.
    """
    ws_objetivo.title = central_norm

    for ws in list(wb.worksheets):
        if ws.title != ws_objetivo.title:
            wb.remove(ws)


def escribir_fila_desde_original(ws, fila_destino, actividad, archivo_info):
    """
    Escribe una fila consolidada copiando el rango C:AU desde datos_originales.columnas_excel.

    Si alguna columna no viene en el JSON, usa fallback desde pms_actividades.
    """
    columnas_originales = obtener_columnas_originales(actividad)

    col_inicio = column_index_from_string(COL_INICIO_DATOS)
    col_fin = column_index_from_string(COL_FIN_DATOS)

    for col_idx in range(col_inicio, col_fin + 1):
        letra = get_column_letter(col_idx)
        valor = columnas_originales.get(letra)
        ws.cell(row=fila_destino, column=col_idx).value = convertir_valor_excel(valor)

    # Fallbacks y normalizaciones clave.
    proveedor = obtener_valor(
        actividad,
        "proveedor",
        default=archivo_info.get("proveedor", "") if archivo_info else "",
    )

    # Si AD está vacío, ponemos proveedor.
    col_empresa = COLS_FALLBACK["empresa"]
    if ws[f"{col_empresa}{fila_destino}"].value in [None, ""]:
        ws[f"{col_empresa}{fila_destino}"] = proveedor

    # Forzar central normalizada en D para que el consolidado quede limpio.
    ws[f"{COLS_FALLBACK['central']}{fila_destino}"] = obtener_valor(
        actividad,
        "central",
        default=normalizar_central(archivo_info.get("central_presentada", "")) if archivo_info else "",
    )

    # Si alguna columna principal vino vacía, rellenar con campos planos.
    ws[f"{COLS_FALLBACK['ot_grafo']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["ot_grafo"],
        "ot_grafo",
    )

    ws[f"{COLS_FALLBACK['unidad']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["unidad"],
        "unidad",
    )

    ws[f"{COLS_FALLBACK['sistema']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["sistema"],
        "sistema",
    )

    ws[f"{COLS_FALLBACK['equipo']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["equipo"],
        "equipo",
    )

    ws[f"{COLS_FALLBACK['tipo_mant']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["tipo_mant"],
        "tipo_mant",
    )

    ws[f"{COLS_FALLBACK['condicion']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["condicion"],
        "condicion",
    )

    ws[f"{COLS_FALLBACK['riesgo']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["riesgo"],
        "riesgo",
    )

    ws[f"{COLS_FALLBACK['inspector']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["inspector"],
        "inspector",
        "inspector_responsable",
    )

    ws[f"{COLS_FALLBACK['rt_terceros']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["rt_terceros"],
        "rt_terceros",
    )

    ws[f"{COLS_FALLBACK['actividad']}{fila_destino}"] = valor_original_o_fallback(
        actividad,
        columnas_originales,
        COLS_FALLBACK["actividad"],
        "actividad",
        "motivo",
    )


def preparar_filas_destino(ws, total_actividades):
    """
    Prepara suficientes filas copiando el estilo de la fila modelo.
    No borra encabezados.
    """
    max_col = max(ws.max_column, column_index_from_string(COL_FIN_DATOS))

    if total_actividades <= 0:
        total_actividades = 1

    # Limpiar valores existentes.
    limpiar_area_datos(ws, fila_inicio=FILA_INICIO_DATOS)

    # Insertar filas si la plantilla no tiene suficientes.
    filas_actuales_desde_inicio = max(ws.max_row - FILA_INICIO_DATOS + 1, 1)

    if total_actividades > filas_actuales_desde_inicio:
        filas_extra = total_actividades - filas_actuales_desde_inicio
        ws.insert_rows(ws.max_row + 1, amount=filas_extra)

    # Copiar estilo de fila modelo a todas las filas destino.
    for idx in range(total_actividades):
        fila = FILA_INICIO_DATOS + idx
        copiar_estilo_fila(ws, FILA_MODELO_ESTILO, fila, max_col)


def nombre_archivo_salida(semana, central_norm):
    central_limpia = central_norm.replace(" ", "_")
    return f"PROGRAMA_UNICO_{central_limpia}_{semana}.xlsx"


def generar_programa_unico(
    supabase,
    semana,
    central,
    salida_path=None,
    plantilla_path=None,
):
    """
    Genera el Excel consolidado manteniendo la plantilla.

    Nuevo comportamiento:
    - Usa datos_originales.columnas_excel como fuente principal.
    - Copia C:AU desde cada fila original al Excel consolidado.
    - Mantiene formato de la plantilla destino.
    - Filtra por semana y central.
    """

    semana_inicio = parse_semana_inicio(semana)
    central_norm = normalizar_central(central)

    if central_norm not in ["SANTA ROSA", "VENTANILLA"]:
        raise ValueError("Central no válida. Usa SANTA ROSA o VENTANILLA.")

    if plantilla_path is None:
        plantilla_path = Path(__file__).resolve().parent / PLANTILLA_NOMBRE
    else:
        plantilla_path = Path(plantilla_path)

    if not plantilla_path.exists():
        raise FileNotFoundError(
            f"No se encontró la plantilla: {plantilla_path}. "
            f"Verifica que exista el archivo '{PLANTILLA_NOMBRE}' en el repo."
        )

    archivos_map = consultar_archivos_semana(supabase, semana)
    actividades = consultar_actividades(supabase, semana, central_norm)

    wb = load_workbook(plantilla_path)
    ws = seleccionar_hoja(wb, central_norm)

    dejar_solo_hoja_objetivo(wb, ws, central_norm)
    actualizar_encabezado_semana(ws, semana_inicio)

    preparar_filas_destino(ws, len(actividades))

    for idx, actividad in enumerate(actividades):
        fila = FILA_INICIO_DATOS + idx
        archivo_id = actividad.get("pms_archivo_id")
        archivo_info = archivos_map.get(archivo_id, {})
        escribir_fila_desde_original(ws, fila, actividad, archivo_info)

    if not actividades:
        ws[f"{COLS_FALLBACK['actividad']}{FILA_INICIO_DATOS}"] = (
            f"No hay actividades registradas para {CENTRAL_LABEL[central_norm]} en la semana {semana}."
        )
        ws[f"{COLS_FALLBACK['central']}{FILA_INICIO_DATOS}"] = central_norm

    if salida_path is None:
        tmp_dir = tempfile.mkdtemp(prefix="programa_unico_")
        salida_path = Path(tmp_dir) / nombre_archivo_salida(semana, central_norm)
    else:
        salida_path = Path(salida_path)

    wb.save(salida_path)

    return {
        "ok": True,
        "archivo_generado": str(salida_path),
        "nombre_archivo": salida_path.name,
        "total_actividades": len(actividades),
        "central": central_norm,
        "central_label": CENTRAL_LABEL[central_norm],
        "semana": semana,
    }
