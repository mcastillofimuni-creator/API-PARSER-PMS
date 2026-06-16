import pandas as pd
import openpyxl


def preparar_datos_parser(ruta_excel):
    """
    Parser mínimo temporal.
    Sirve para probar que la API descarga el Excel, lo abre y guarda una validación básica.
    Luego reemplazaremos este archivo por el parser completo.
    """

    wb = openpyxl.load_workbook(ruta_excel, data_only=True, read_only=True)

    hojas_visibles = []
    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        estado = getattr(ws, "sheet_state", "visible")

        if estado == "visible":
            hojas_visibles.append({
                "hoja": nombre_hoja,
                "estado": estado,
                "max_row": ws.max_row,
                "max_column": ws.max_column
            })

    wb.close()

    actividades = []
    observaciones = []

    if not hojas_visibles:
        observaciones.append({
            "nivel": "ERROR",
            "tipo_observacion": "No se encontraron hojas visibles en el archivo.",
            "unidad": "",
            "actividad": "",
            "inspector_responsable": "",
            "fila_excel": 0,
            "campo": "hoja",
            "valor_detectado": "",
            "sugerencia": "Verificar que el PMS tenga al menos una hoja visible con información."
        })

        return {
            "actividades": actividades,
            "observaciones": observaciones,
            "hojas": hojas_visibles,
            "errores": 1,
            "advertencias": 0,
            "estado": "ERROR - SIN HOJAS VISIBLES"
        }

    # Observación temporal para confirmar que el parser leyó el archivo.
    observaciones.append({
        "nivel": "ADVERTENCIA",
        "tipo_observacion": f"Parser temporal ejecutado correctamente. Hojas visibles detectadas: {len(hojas_visibles)}.",
        "unidad": "",
        "actividad": "",
        "inspector_responsable": "",
        "fila_excel": 0,
        "campo": "archivo",
        "valor_detectado": str([h["hoja"] for h in hojas_visibles]),
        "sugerencia": "Luego se reemplazará este parser temporal por el parser completo de validación PMS."
    })

    return {
        "actividades": actividades,
        "observaciones": observaciones,
        "hojas": hojas_visibles,
        "errores": 0,
        "advertencias": 1,
        "estado": "VALIDADO - PARSER TEMPORAL"
    }
