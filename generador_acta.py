import os
import re
import tempfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document


TEMPLATE_NAME = "PLANTILLA_ACTA_INTERFERENCIAS_BASE.docx"


def _root_dir() -> Path:
    return Path(__file__).resolve().parent


def _template_path() -> Path:
    path = _root_dir() / TEMPLATE_NAME
    if not path.exists():
        raise FileNotFoundError(
            f"No se encontró la plantilla {TEMPLATE_NAME}. "
            "Debe estar en la misma carpeta que main.py."
        )
    return path


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalizar_espacios(texto: str) -> str:
    return re.sub(r"\s+", " ", _safe_text(texto)).strip()


def _replace_in_paragraph(paragraph, replacements: Dict[str, str]) -> None:
    """
    Reemplazo conservador de placeholders dentro de párrafos.
    Si el placeholder está partido en varios runs, reconstruye el párrafo.
    """
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    new_text = full_text
    for key, value in replacements.items():
        new_text = new_text.replace(key, value)

    if new_text == full_text:
        return

    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)


def _replace_in_cell(cell, replacements: Dict[str, str]) -> None:
    for p in cell.paragraphs:
        _replace_in_paragraph(p, replacements)

    for table in cell.tables:
        for row in table.rows:
            for c in row.cells:
                _replace_in_cell(c, replacements)


def _replace_all(doc: Document, replacements: Dict[str, str]) -> None:
    for p in doc.paragraphs:
        _replace_in_paragraph(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _replace_in_cell(cell, replacements)

    for section in doc.sections:
        for container in [section.header, section.footer]:
            for p in container.paragraphs:
                _replace_in_paragraph(p, replacements)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        _replace_in_cell(cell, replacements)


def _set_cell_text(cell, text: str) -> None:
    cell.text = _safe_text(text)


def _copiar_estilo_fila(row_origen, row_destino) -> None:
    """
    Copia propiedades de fila y celdas para que las filas agregadas mantengan
    el formato visual de la plantilla.
    """
    row_destino._tr.get_or_add_trPr().clear()
    for child in row_origen._tr.get_or_add_trPr():
        row_destino._tr.get_or_add_trPr().append(deepcopy(child))

    for src_cell, dst_cell in zip(row_origen.cells, row_destino.cells):
        src_tcpr = src_cell._tc.get_or_add_tcPr()
        dst_tcpr = dst_cell._tc.get_or_add_tcPr()
        dst_tcpr.clear()
        for child in src_tcpr:
            dst_tcpr.append(deepcopy(child))


def _asegurar_filas(tabla, cantidad_total: int, fila_modelo_idx: int = 1) -> None:
    while len(tabla.rows) < cantidad_total:
        new_row = tabla.add_row()
        try:
            _copiar_estilo_fila(tabla.rows[fila_modelo_idx], new_row)
        except Exception:
            pass


def _limpiar_filas_participantes(tabla) -> None:
    # Mantener encabezado. Limpiar filas de datos.
    for idx in range(1, len(tabla.rows)):
        row = tabla.rows[idx]
        for cidx, cell in enumerate(row.cells):
            _set_cell_text(cell, str(idx) if cidx == 0 else "")


def _extraer_participantes_adicionales(notas: str) -> List[Dict[str, str]]:
    """
    Parser simple de notas. Detecta líneas como:
    - Hector Tinoco (Orygen - HSE&Q)
    - Manuel Castillo (Orygen - Mantenimiento)
    - Carolina Mostajo - ORYGEN
    """
    participantes = []
    vistos = set()

    for raw in _safe_text(notas).splitlines():
        line = raw.strip(" -*\t")
        if not line:
            continue

        m = re.match(r"^([A-Za-zÁÉÍÓÚÜÑáéíóúüñ .']{3,})\s*\(([^)]+)\)\s*$", line)
        if m:
            nombre = _normalizar_espacios(m.group(1))
            info = _normalizar_espacios(m.group(2))
            empresa = "ORYGEN" if "ORYGEN" in info.upper() else info.upper()
            key = (nombre.upper(), empresa.upper())
            if key not in vistos:
                participantes.append({"nombre": nombre, "contrato": "", "empresa": empresa})
                vistos.add(key)
            continue

        m = re.match(r"^([A-Za-zÁÉÍÓÚÜÑáéíóúüñ .']{3,})\s*[-–]\s*([A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ &.]+)\s*$", line)
        if m:
            nombre = _normalizar_espacios(m.group(1))
            empresa = _normalizar_espacios(m.group(2)).upper()
            if "ORYGEN" in empresa or "HSE" in empresa or "MANT" in empresa:
                empresa_final = "ORYGEN" if "ORYGEN" in empresa else empresa
                key = (nombre.upper(), empresa_final.upper())
                if key not in vistos:
                    participantes.append({"nombre": nombre, "contrato": "", "empresa": empresa_final})
                    vistos.add(key)

    return participantes


def _estado_simple(estado: str) -> str:
    e = _safe_text(estado).upper()
    if not e:
        return "pendiente de validación"
    if "CONFORME" in e and "OBS" not in e:
        return "conforme"
    if "SIN ARCHIVO" in e:
        return "sin archivo"
    if "OBSERVADO" in e or "CORRECCIÓN" in e or "CORRECCION" in e:
        return "con observaciones"
    if "ERROR" in e:
        return "con error de validación"
    return e.lower()


def _armar_observaciones(payload: Dict[str, Any]) -> str:
    notas = _safe_text(payload.get("notas"))
    empresas = payload.get("empresas") or []
    faltantes = payload.get("faltantes") or []

    total = len(empresas)
    observadas = [
        e for e in empresas
        if any(k in _safe_text(e.get("estado_validacion")).upper() for k in ["OBSERVADO", "CORRECCIÓN", "CORRECCION", "ERROR", "SIN ARCHIVO"])
    ]

    lineas = []
    lineas.append(
        f"Se efectuó la reunión de coordinación e identificación de interferencias correspondiente al {payload.get('pms') or 'PMS'}, "
        f"para la semana {payload.get('rango_semana') or payload.get('semana') or ''}, en la instalación {payload.get('central') or ''}."
    )

    if total:
        lineas.append(
            f"Se revisó la programación semanal registrada en la plataforma, con participación de {total} empresa(s) proveedora(s)."
        )

    if observadas:
        resumen_obs = ", ".join(
            f"{_safe_text(e.get('empresa'))} ({_estado_simple(e.get('estado_validacion'))})"
            for e in observadas[:10]
        )
        lineas.append(f"Se identificaron programas que requieren seguimiento o revisión: {resumen_obs}.")
    else:
        lineas.append("No se identificaron observaciones relevantes en la documentación revisada.")

    if faltantes:
        lineas.append(f"Empresas pendientes de registro o carga de programa: {', '.join(map(str, faltantes))}.")

    if notas:
        lineas.append("")
        lineas.append("Comentarios relevantes registrados:")
        lineas.append(notas)

    # Mantener una frase estándar útil para el formato.
    if "interferencia" not in notas.lower():
        lineas.append("")
        lineas.append("No se identificaron interferencias relevantes adicionales durante la reunión, salvo las que pudieran derivarse de la coordinación operativa de las actividades programadas.")

    return "\n".join([l for l in lineas if l is not None]).strip()


def _armar_acciones(payload: Dict[str, Any]) -> List[Dict[str, str]]:
    acciones = list(payload.get("acciones") or [])
    empresas = payload.get("empresas") or []

    if acciones:
        return [
            {
                "accion": _safe_text(a.get("accion") or a.get("descripcion")),
                "responsable": _safe_text(a.get("responsable")),
            }
            for a in acciones
            if _safe_text(a.get("accion") or a.get("descripcion"))
        ]

    generadas = []

    observadas = [
        e for e in empresas
        if any(k in _safe_text(e.get("estado_validacion")).upper() for k in ["OBSERVADO", "CORRECCIÓN", "CORRECCION", "ERROR"])
    ]
    sin_archivo = [
        e for e in empresas
        if "SIN ARCHIVO" in _safe_text(e.get("estado_validacion")).upper()
    ]

    for e in observadas[:5]:
        empresa = _safe_text(e.get("empresa"))
        generadas.append({
            "accion": f"Regularizar las observaciones identificadas en el programa PMS presentado por {empresa}.",
            "responsable": empresa,
        })

    for e in sin_archivo[:3]:
        empresa = _safe_text(e.get("empresa"))
        generadas.append({
            "accion": f"Completar la carga del archivo de programa semanal correspondiente a {empresa}.",
            "responsable": empresa,
        })

    faltantes = payload.get("faltantes") or []
    for empresa in faltantes[:5]:
        empresa = _safe_text(empresa)
        if empresa and not any(a.get("responsable", "").upper() == empresa.upper() for a in generadas):
            generadas.append({
                "accion": f"Coordinar con {empresa} la carga o regularización de su programa semanal.",
                "responsable": empresa,
            })

    if not generadas:
        generadas.append({
            "accion": "Mantener la coordinación semanal de actividades y comunicar oportunamente cualquier interferencia, restricción o cambio en la programación.",
            "responsable": "PROVEEDORES / ORYGEN",
        })

    return generadas


def generar_acta_interferencias(payload: Dict[str, Any]) -> Dict[str, str]:
    """
    Genera un DOCX usando la plantilla oficial RG02-P.HS.PE.013.
    No usa IA: rellena el formato con reglas y datos estructurados.
    """
    doc = Document(str(_template_path()))

    fecha = _safe_text(payload.get("fecha_reunion"))
    if not fecha:
        fecha = datetime.now().strftime("%d.%m.%Y")

    central = _safe_text(payload.get("central") or payload.get("central_label"))
    observaciones = _armar_observaciones(payload)
    acciones = _armar_acciones(payload)

    replacements = {
        "[[FECHA_REUNION]]": fecha,
        "[[CENTRAL]]": central,
        "[[OBSERVACIONES]]": observaciones,
        "[[ACCION_01]]": acciones[0]["accion"] if len(acciones) > 0 else "",
        "[[RESPONSABLE_01]]": acciones[0]["responsable"] if len(acciones) > 0 else "",
        "[[ACCION_02]]": acciones[1]["accion"] if len(acciones) > 1 else "",
        "[[RESPONSABLE_02]]": acciones[1]["responsable"] if len(acciones) > 1 else "",
    }
    _replace_all(doc, replacements)

    # Participantes: tabla 1
    participantes = []
    vistos = set()

    for e in payload.get("empresas") or []:
        nombre = _safe_text(e.get("expositor"))
        empresa = _safe_text(e.get("empresa"))
        contrato = _safe_text(e.get("contrato"))
        if not nombre and not empresa:
            continue
        key = (nombre.upper(), empresa.upper())
        if key in vistos:
            continue
        participantes.append({"nombre": nombre, "contrato": contrato, "empresa": empresa})
        vistos.add(key)

    for p in payload.get("participantes_adicionales") or []:
        nombre = _safe_text(p.get("nombre"))
        empresa = _safe_text(p.get("empresa") or "ORYGEN")
        contrato = _safe_text(p.get("contrato"))
        if not nombre:
            continue
        key = (nombre.upper(), empresa.upper())
        if key in vistos:
            continue
        participantes.append({"nombre": nombre, "contrato": contrato, "empresa": empresa})
        vistos.add(key)

    for p in _extraer_participantes_adicionales(payload.get("notas") or ""):
        key = (p["nombre"].upper(), p["empresa"].upper())
        if key in vistos:
            continue
        participantes.append(p)
        vistos.add(key)

    if len(doc.tables) >= 2:
        tabla_part = doc.tables[1]
        _asegurar_filas(tabla_part, max(15, len(participantes) + 1), fila_modelo_idx=1)
        _limpiar_filas_participantes(tabla_part)

        for idx, p in enumerate(participantes, start=1):
            row = tabla_part.rows[idx]
            _set_cell_text(row.cells[0], str(idx))
            _set_cell_text(row.cells[1], p.get("nombre", ""))
            _set_cell_text(row.cells[2], p.get("contrato", ""))
            _set_cell_text(row.cells[3], p.get("empresa", ""))
            _set_cell_text(row.cells[4], "")

    # Acciones extra: tabla 2, filas 3 en adelante.
    if len(doc.tables) >= 3:
        tabla_acc = doc.tables[2]
        primera_fila_accion = 3
        filas_necesarias = primera_fila_accion + max(len(acciones), 2)
        _asegurar_filas(tabla_acc, filas_necesarias, fila_modelo_idx=3)

        for i in range(primera_fila_accion, len(tabla_acc.rows)):
            for c in tabla_acc.rows[i].cells:
                _set_cell_text(c, "")

        for i, accion in enumerate(acciones):
            row = tabla_acc.rows[primera_fila_accion + i]
            _set_cell_text(row.cells[0], accion.get("accion", ""))
            _set_cell_text(row.cells[1], accion.get("responsable", ""))

    out = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    out.close()

    doc.save(out.name)

    pms = _safe_text(payload.get("pms")).replace(" ", "_") or "PMS"
    central_file = re.sub(r"[^A-Z0-9]+", "_", central.upper()).strip("_") or "CENTRAL"
    fecha_file = fecha.replace("/", ".").replace("-", ".")
    nombre_archivo = f"ACTA_INTERFERENCIAS_{pms}_{central_file}_{fecha_file}.docx"

    return {
        "archivo_generado": out.name,
        "nombre_archivo": nombre_archivo,
    }
