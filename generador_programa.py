import os
import re
import tempfile
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook


PLANTILLA_NOMBRE = "PROGRAMA SEMANAL PLANTILLA.xlsx"

CENTRALES_VALIDAS = {
    "C. T. SANTA ROSA": "SANTA ROSA",
    "CT SANTA ROSA": "SANTA ROSA",
    "C.T. SANTA ROSA": "SANTA ROSA",
    "SANTA ROSA": "SANTA ROSA",
    "C.C. VENTANILLA": "VENTANILLA",
    "CC VENTANILLA": "VENTANILLA",
    "C.C VENTANILLA": "VENTANILLA",
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

# Columnas principales según la plantilla compartida.
# Ajustables si luego identificamos una columna exacta distinta.
COLS = {
    "ot_grafo": "C",
    "central": "D",
    "unidad": "E",
    "sistema": "F",
    "equipo": "G",
    "condicion": "L",
    "riesgo": "M",
    "inspector": "O",
    "rt_terceros": "P",
    "empresa": "AD",
    "actividad": "AF",
}

# Columnas de días en la plantilla.
DAY_COLS = {
    0: "W",   # Sábado
    1: "X",   # Domingo
    2: "Y",   # Lunes
    3: "Z",   # Martes
    4: "AA",  # Miércoles
    5: "AB",  # Jueves
    6: "AC",  # Viernes
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


def limpiar_area_datos(ws, fila_inicio=10):
    """
    Limpia valores de las filas de datos sin destruir formato.
    """
    for row in ws.iter_rows(min_row=fila_inicio, max_row=ws.max_row):
        for cell in row:
            # No tocar encabezados, solo contenido.
            cell.value = None


def obtener_valor(row, *keys, default=""):
    for key in keys:
        if key in row and row.get(key) not in [None, ""]:
            return row.get(key)
    return default


def obtener_dias_actividad(row, archivo_info):
    """
    Prioridad:
    1. Si la actividad trae días propios, usar esos.
    2. Si no, usar días del archivo/subida.
    """
    dias = row.get("dias")

    if dias is None:
        dias = row.get("dias_programados")

    if dias is None and archivo_info:
        dias = archivo_info.get("dias")

    if dias is None:
        return []

    if isinstance(dias, list):
        salida = []
        for d in dias:
            try:
                salida.append(int(d))
            except Exception:
                pass
        return sorted(set([d for d in salida if 0 <= d <= 6]))

    if isinstance(dias, str):
        # Puede venir como "[2,3]" o "2,3" o "Lun, Mar".
        texto = dias.strip()

        if texto.startswith("[") and texto.endswith("]"):
            texto = texto[1:-1]

        mapa = {
            "SAB": 0,
            "SÁB": 0,
            "DOM": 1,
            "LUN": 2,
            "MAR": 3,
            "MIE": 4,
            "MIÉ": 4,
            "JUE": 5,
            "VIE": 6,
        }

        salida = []

        for parte in re.split(r"[,\s;/]+", texto.upper()):
            parte = parte.strip()
            if not parte:
                continue

            if parte.isdigit():
                n = int(parte)
                if 0 <= n <= 6:
                    salida.append(n)

            elif parte in mapa:
                salida.append(mapa[parte])

        return sorted(set(salida))

    return []


def actualizar_encabezado_semana(ws, semana_inicio):
    """
    Actualiza de forma robusta los textos de fechas del encabezado.
    No depende de una sola celda exacta.
    """
    semana_fin = semana_inicio + timedelta(days=6)

    titulo = f"PROGRAMA SEMANAL DEL {fecha_larga(semana_inicio)} AL {fecha_larga(semana_fin)}"
    semana_texto = f"SEMANA DEL {fecha_corta(semana_inicio)} AL {fecha_corta(semana_fin)}"

    # Buscar celdas de encabezado donde pueda estar el título.
    for row in ws.iter_rows(min_row=1, max_row=9):
        for cell in row:
            if isinstance(cell.value, str):
                texto = cell.value.upper()

                if "PROGRAMA SEMANAL" in texto and ("DEL" in texto or "SEMANA" in texto):
                    cell.value = titulo

                if "SEMANA" in texto and ("AL" in texto or "DEL" in texto):
                    cell.value = semana_texto

    # Además, escribir fechas en la zona de días W:AC.
    # En muchas plantillas la fila 8 o 9 contiene los días/fechas.
    for i in range(7):
        fecha = semana_inicio + timedelta(days=i)
        col = DAY_COLS[i]

        # Fila 8: fecha.
        ws[f"{col}8"] = fecha
        ws[f"{col}8"].number_format = "dd/mm/yyyy"

        # Fila 9: día corto.
        ws[f"{col}9"] = f"{DIAS_CORTOS[i]}\n{fecha.day:02d}/{fecha.month:02d}"


def consultar_archivos_semana(supabase, semana):
    resp = (
        supabase.table("pms_archivos")
        .select(
            "id,semana,proveedor,expositor,central_presentada,archivo_nombre,"
            "estado_validacion,errores,advertencias,actividades,observaciones,dias"
        )
        .eq("semana", semana)
        .execute()
    )

    data = resp.data or []
    return {x["id"]: x for x in data if x.get("id")}


def consultar_actividades(supabase, semana, central_norm):
    """
    Consulta actividades de la semana y central.
    Se consulta por pms_actividades.central porque el archivo puede traer varias centrales.
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


def escribir_actividad(ws, fila, actividad, archivo_info):
    proveedor = obtener_valor(
        actividad,
        "proveedor",
        default=archivo_info.get("proveedor", "") if archivo_info else "",
    )

    central = obtener_valor(actividad, "central", default="")
    unidad = obtener_valor(actividad, "unidad", "grupo", default="")
    sistema = obtener_valor(actividad, "sistema", default="")
    equipo = obtener_valor(actividad, "equipo", "sub_sistema", "subsistema", default="")
    actividad_txt = obtener_valor(actividad, "actividad", "descripcion", default="")
    ot_grafo = obtener_valor(actividad, "ot_grafo", "ot", "grafo", default="")
    condicion = obtener_valor(actividad, "condicion", "condición", default="")
    riesgo = obtener_valor(actividad, "riesgo", default="")
    inspector = obtener_valor(actividad, "inspector", "inspector_responsable", default="")
    rt_terceros = obtener_valor(actividad, "rt_terceros", "rt", default="")

    ws[f"{COLS['ot_grafo']}{fila}"] = ot_grafo
    ws[f"{COLS['central']}{fila}"] = central
    ws[f"{COLS['unidad']}{fila}"] = unidad
    ws[f"{COLS['sistema']}{fila}"] = sistema
    ws[f"{COLS['equipo']}{fila}"] = equipo
    ws[f"{COLS['condicion']}{fila}"] = condicion
    ws[f"{COLS['riesgo']}{fila}"] = riesgo
    ws[f"{COLS['inspector']}{fila}"] = inspector
    ws[f"{COLS['rt_terceros']}{fila}"] = rt_terceros
    ws[f"{COLS['empresa']}{fila}"] = proveedor
    ws[f"{COLS['actividad']}{fila}"] = actividad_txt

    dias = obtener_dias_actividad(actividad, archivo_info)

    for i, col in DAY_COLS.items():
        ws[f"{col}{fila}"] = "X" if i in dias else ""


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

    # Buscar coincidencia parcial.
    for ws in wb.worksheets:
        nombre = ws.title.upper()
        if central_norm in nombre:
            return ws

    return wb.active


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

    Parámetros:
    - supabase: cliente Supabase ya inicializado.
    - semana: texto YYYY-MM-DD. Ejemplo: 2026-06-13.
    - central: "SANTA ROSA", "VENTANILLA", "C. T. Santa Rosa" o "C.C. Ventanilla".
    - salida_path: ruta opcional donde guardar el archivo.
    - plantilla_path: ruta opcional de la plantilla.

    Retorna:
    {
      "ok": True,
      "archivo_generado": "...",
      "total_actividades": 38,
      "central": "SANTA ROSA",
      "semana": "2026-06-13"
    }
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

    actualizar_encabezado_semana(ws, semana_inicio)

    fila_inicio = 10
    fila_base_estilo = 10
    max_col = ws.max_column

    limpiar_area_datos(ws, fila_inicio=fila_inicio)

    # Si hay más actividades que filas disponibles en la plantilla, insertar filas.
    filas_disponibles = max(ws.max_row - fila_inicio + 1, 1)
    if len(actividades) > filas_disponibles:
        filas_extra = len(actividades) - filas_disponibles
        ws.insert_rows(ws.max_row + 1, amount=filas_extra)

    # Copiar estilo de la fila base a las filas que se usarán.
    for idx in range(len(actividades)):
        fila = fila_inicio + idx
        copiar_estilo_fila(ws, fila_base_estilo, fila, max_col)

    # Escribir actividades consolidadas.
    for idx, actividad in enumerate(actividades):
        fila = fila_inicio + idx
        archivo_id = actividad.get("pms_archivo_id")
        archivo_info = archivos_map.get(archivo_id, {})
        escribir_actividad(ws, fila, actividad, archivo_info)

    # Si no hay actividades, dejar una nota en la primera fila.
    if not actividades:
        ws[f"{COLS['actividad']}{fila_inicio}"] = (
            f"No hay actividades registradas para {CENTRAL_LABEL[central_norm]} en la semana {semana}."
        )

    # Guardar archivo.
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
