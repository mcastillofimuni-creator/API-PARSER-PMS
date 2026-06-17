import os
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from supabase import create_client, Client

from parser_pms import preparar_datos_parser


# ============================================================
# CARGA DE VARIABLES DE ENTORNO
# ============================================================

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "").strip()
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "pms-archivos").strip()
FRONTEND_ORIGIN = os.getenv("FRONTEND_ORIGIN", "*").strip()

if not SUPABASE_URL:
    raise RuntimeError("Falta variable de entorno SUPABASE_URL")

if not SUPABASE_SERVICE_ROLE_KEY:
    raise RuntimeError("Falta variable de entorno SUPABASE_SERVICE_ROLE_KEY")


supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)


# ============================================================
# APP FASTAPI
# ============================================================

app = FastAPI(
    title="PMS Parser API",
    version="1.0.0",
    description="API para validar PMS semanales y guardar actividades/observaciones en Supabase.",
)

allowed_origins = ["*"] if FRONTEND_ORIGIN == "*" else [
    origin.strip()
    for origin in FRONTEND_ORIGIN.split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# MODELOS
# ============================================================

class ValidarRequest(BaseModel):
    pms_archivo_id: str


# ============================================================
# UTILIDADES
# ============================================================

def ejecutar_select_archivo(pms_archivo_id: str) -> Dict[str, Any]:
    """
    Obtiene la fila del archivo PMS desde pms_archivos.
    """

    try:
        resp = (
            supabase
            .table("pms_archivos")
            .select("*")
            .eq("id", pms_archivo_id)
            .single()
            .execute()
        )

        if not resp.data:
            raise HTTPException(
                status_code=404,
                detail=f"No se encontró pms_archivo_id={pms_archivo_id}",
            )

        return resp.data

    except HTTPException:
        raise

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Error consultando pms_archivos: {str(exc)}",
        )


def descargar_archivo_storage(archivo_path: str, nombre_archivo: Optional[str] = None) -> str:
    """
    Descarga el Excel desde Supabase Storage y lo guarda temporalmente.
    Retorna la ruta local.
    """

    if not archivo_path:
        raise HTTPException(
            status_code=400,
            detail="El registro no tiene archivo_path. No hay archivo para validar.",
        )

    try:
        contenido = supabase.storage.from_(SUPABASE_BUCKET).download(archivo_path)

        suffix = ".xlsx"

        if nombre_archivo:
            nombre_lower = nombre_archivo.lower()
            if nombre_lower.endswith(".xlsm"):
                suffix = ".xlsm"
            elif nombre_lower.endswith(".xls"):
                suffix = ".xls"
            elif nombre_lower.endswith(".csv"):
                suffix = ".csv"
            elif nombre_lower.endswith(".xlsx"):
                suffix = ".xlsx"

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(contenido)
        tmp.flush()
        tmp.close()

        return tmp.name

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo descargar el archivo desde Storage: {str(exc)}",
        )


def limpiar_resultados_previos(pms_archivo_id: str) -> None:
    """
    Borra actividades y observaciones previas para evitar duplicados
    cuando se valida nuevamente el mismo PMS.
    """

    try:
        (
            supabase
            .table("pms_actividades")
            .delete()
            .eq("pms_archivo_id", pms_archivo_id)
            .execute()
        )

        (
            supabase
            .table("pms_observaciones")
            .delete()
            .eq("pms_archivo_id", pms_archivo_id)
            .execute()
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudieron limpiar resultados previos: {str(exc)}",
        )


def insertar_actividades(
    pms_archivo_id: str,
    semana: str,
    proveedor: str,
    actividades: List[Dict[str, Any]],
) -> None:
    """
    Inserta actividades parseadas en pms_actividades.
    """

    if not actividades:
        return

    filas = []

    for act in actividades:
        filas.append({
            "pms_archivo_id": pms_archivo_id,
            "semana": semana,
            "proveedor": proveedor,
            "fila_excel": act.get("fila_excel", 0),
            "central": act.get("central", ""),
            "unidad": act.get("unidad", ""),
            "sistema": act.get("sistema", ""),
            "equipo": act.get("equipo", ""),
            "actividad": act.get("actividad", ""),
            "ot_grafo": act.get("ot_grafo", ""),
            "riesgo": act.get("riesgo", ""),
            "inspector": act.get("inspector", ""),
            "rt_terceros": act.get("rt_terceros", ""),
        })

    try:
        (
            supabase
            .table("pms_actividades")
            .insert(filas)
            .execute()
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudieron insertar actividades: {str(exc)}",
        )


def insertar_observaciones(
    pms_archivo_id: str,
    semana: str,
    proveedor: str,
    observaciones: List[Dict[str, Any]],
) -> None:
    """
    Inserta observaciones parseadas en pms_observaciones.
    """

    if not observaciones:
        return

    filas = []

    for obs in observaciones:
        filas.append({
            "pms_archivo_id": pms_archivo_id,
            "semana": semana,
            "proveedor": proveedor,
            "nivel": obs.get("nivel", ""),
            "central": obs.get("central", ""),
            "tipo_observacion": obs.get("tipo_observacion", ""),
            "unidad": obs.get("unidad", ""),
            "actividad": obs.get("actividad", ""),
            "inspector_responsable": obs.get("inspector_responsable", ""),
            "fila_excel": obs.get("fila_excel", 0),
            "campo": obs.get("campo", ""),
            "valor_detectado": obs.get("valor_detectado", ""),
            "sugerencia": obs.get("sugerencia", ""),
        })

    try:
        (
            supabase
            .table("pms_observaciones")
            .insert(filas)
            .execute()
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudieron insertar observaciones: {str(exc)}",
        )


def actualizar_archivo_validado(
    pms_archivo_id: str,
    resultado: Dict[str, Any],
) -> None:
    """
    Actualiza resumen de validación en pms_archivos.
    """

    payload = {
        "estado_validacion": resultado.get("estado", ""),
        "errores": int(resultado.get("errores", 0) or 0),
        "advertencias": int(resultado.get("advertencias", 0) or 0),
        "actividades": len(resultado.get("actividades", []) or []),
        "observaciones": len(resultado.get("observaciones", []) or []),
        "centrales_detectadas": resultado.get("centrales_detectadas", []) or [],
    }

    try:
        (
            supabase
            .table("pms_archivos")
            .update(payload)
            .eq("id", pms_archivo_id)
            .execute()
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo actualizar pms_archivos: {str(exc)}",
        )


# ============================================================
# ENDPOINTS
# ============================================================

@app.get("/")
def root():
    return {
        "ok": True,
        "service": "PMS Parser API",
        "message": "API activa",
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "status": "healthy",
    }


@app.post("/validar-pms")
def validar_pms(req: ValidarRequest):
    """
    Valida un PMS ya registrado en Supabase.

    Flujo:
    1. Busca el registro en pms_archivos.
    2. Descarga el Excel desde Storage.
    3. Ejecuta parser_pms.py.
    4. Borra resultados anteriores.
    5. Inserta actividades.
    6. Inserta observaciones.
    7. Actualiza resumen en pms_archivos.
    """

    pms_archivo_id = req.pms_archivo_id.strip()

    if not pms_archivo_id:
        raise HTTPException(
            status_code=400,
            detail="pms_archivo_id es obligatorio.",
        )

    archivo = ejecutar_select_archivo(pms_archivo_id)

    semana = archivo.get("semana", "") or ""
    proveedor = archivo.get("proveedor", "") or ""
    archivo_path = archivo.get("archivo_path", "") or ""
    archivo_nombre = archivo.get("archivo_nombre", "") or ""
    central_presentada = archivo.get("central_presentada", "") or ""

    ruta_temporal = None

    try:
        ruta_temporal = descargar_archivo_storage(
            archivo_path=archivo_path,
            nombre_archivo=archivo_nombre,
        )

        resultado = preparar_datos_parser(
            ruta_temporal,
            central_presentada=central_presentada,
        )

        limpiar_resultados_previos(pms_archivo_id)

        insertar_actividades(
            pms_archivo_id=pms_archivo_id,
            semana=semana,
            proveedor=proveedor,
            actividades=resultado.get("actividades", []),
        )

        insertar_observaciones(
            pms_archivo_id=pms_archivo_id,
            semana=semana,
            proveedor=proveedor,
            observaciones=resultado.get("observaciones", []),
        )

        actualizar_archivo_validado(
            pms_archivo_id=pms_archivo_id,
            resultado=resultado,
        )

        return {
            "ok": True,
            "pms_archivo_id": pms_archivo_id,
            "proveedor": proveedor,
            "semana": semana,
            "central_presentada": resultado.get("central_presentada", central_presentada),
            "central_presentada_norm": resultado.get("central_presentada_norm", ""),
            "centrales_detectadas": resultado.get("centrales_detectadas", []),
            "archivo_path": archivo_path,
            "estado": resultado.get("estado", ""),
            "errores": resultado.get("errores", 0),
            "advertencias": resultado.get("advertencias", 0),
            "actividades": len(resultado.get("actividades", []) or []),
            "observaciones": len(resultado.get("observaciones", []) or []),
            "detalle_observaciones": resultado.get("observaciones", []),
        }

    finally:
        if ruta_temporal:
            try:
                Path(ruta_temporal).unlink(missing_ok=True)
            except Exception:
                pass
