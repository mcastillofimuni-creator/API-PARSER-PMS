import os
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from supabase import create_client, Client

from parser_pms import preparar_datos_parser
from generador_programa import generar_programa_unico


SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "").strip()
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "pms-archivos").strip()
FRONTEND_ORIGIN = os.getenv("FRONTEND_ORIGIN", "*").strip()

if not SUPABASE_URL:
    raise RuntimeError("Falta la variable de entorno SUPABASE_URL.")

if not SUPABASE_SERVICE_ROLE_KEY:
    raise RuntimeError("Falta la variable de entorno SUPABASE_SERVICE_ROLE_KEY.")


supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

app = FastAPI(
    title="PMS Parser API",
    version="1.1.0",
    description="API para validar PMS semanales y generar programa único consolidado.",
)

origins = ["*"] if FRONTEND_ORIGIN == "*" else [
    FRONTEND_ORIGIN,
    "https://pms-orygen.vercel.app",
    "https://pms-orygen-git-main-mancastle-s-projects.vercel.app",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)


class ValidarRequest(BaseModel):
    pms_archivo_id: str


class GenerarProgramaRequest(BaseModel):
    semana: str
    central: str


@app.get("/")
def root():
    return {
        "ok": True,
        "service": "PMS Parser API",
        "message": "API activa",
        "endpoints": [
            "/health",
            "/validar-pms",
            "/generar-programa-unico",
        ],
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "status": "healthy",
    }


def descargar_archivo_storage(archivo_path: str) -> str:
    """
    Descarga desde Supabase Storage el Excel asociado al PMS.
    Retorna la ruta temporal local.
    """
    if not archivo_path:
        raise HTTPException(status_code=400, detail="El registro no tiene archivo_path.")

    try:
        contenido = supabase.storage.from_(SUPABASE_BUCKET).download(archivo_path)
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo descargar el archivo desde Storage: {exc}",
        )

    suffix = Path(archivo_path).suffix or ".xlsx"

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(contenido)
    tmp.flush()
    tmp.close()

    return tmp.name


def obtener_registro_archivo(pms_archivo_id: str) -> Dict[str, Any]:
    """
    Busca el registro principal en pms_archivos.
    """
    resp = (
        supabase.table("pms_archivos")
        .select("*")
        .eq("id", pms_archivo_id)
        .single()
        .execute()
    )

    if not resp.data:
        raise HTTPException(
            status_code=404,
            detail=f"No se encontró pms_archivo_id: {pms_archivo_id}",
        )

    return resp.data


def borrar_resultados_previos(pms_archivo_id: str):
    """
    Borra actividades y observaciones previas del mismo archivo.
    Esto permite revalidar o reemplazar archivo sin duplicar resultados.
    """
    try:
        supabase.table("pms_observaciones").delete().eq("pms_archivo_id", pms_archivo_id).execute()
        supabase.table("pms_actividades").delete().eq("pms_archivo_id", pms_archivo_id).execute()
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudieron borrar resultados anteriores: {exc}",
        )


def normalizar_lista(valor):
    if valor is None:
        return []

    if isinstance(valor, list):
        return valor

    return []


def insertar_actividades(pms_archivo_id: str, resultado: Dict[str, Any]):
    """
    Inserta actividades devueltas por parser_pms.py.
    Soporta varios nombres por compatibilidad:
    - actividades_detalle
    - detalle_actividades
    - actividades_data
    """
    actividades = (
        resultado.get("actividades_detalle")
        or resultado.get("detalle_actividades")
        or resultado.get("actividades_data")
        or []
    )

    if not isinstance(actividades, list) or len(actividades) == 0:
        return

    filas = []

    for a in actividades:
        if not isinstance(a, dict):
            continue

        fila = {
            "pms_archivo_id": pms_archivo_id,
            "semana": resultado.get("semana"),
            "proveedor": resultado.get("proveedor"),
            "fila_excel": a.get("fila_excel"),
            "central": a.get("central"),
            "unidad": a.get("unidad"),
            "sistema": a.get("sistema"),
            "equipo": a.get("equipo"),
            "actividad": a.get("actividad"),
            "ot_grafo": a.get("ot_grafo"),
            "riesgo": a.get("riesgo"),
            "inspector": a.get("inspector") or a.get("inspector_responsable"),
            "rt_terceros": a.get("rt_terceros"),
        }

        # Campos opcionales si existen en tu tabla.
        if "condicion" in a:
            fila["condicion"] = a.get("condicion")

        if "dias" in a:
            fila["dias"] = normalizar_lista(a.get("dias"))

        filas.append(fila)

    if filas:
        try:
            supabase.table("pms_actividades").insert(filas).execute()
        except Exception as exc:
            raise HTTPException(
                status_code=500,
                detail=f"No se pudieron insertar actividades: {exc}",
            )


def insertar_observaciones(pms_archivo_id: str, resultado: Dict[str, Any]):
    """
    Inserta observaciones devueltas por parser_pms.py.
    """
    observaciones = (
        resultado.get("detalle_observaciones")
        or resultado.get("observaciones_detalle")
        or resultado.get("observaciones_data")
        or []
    )

    if not isinstance(observaciones, list) or len(observaciones) == 0:
        return

    filas = []

    for o in observaciones:
        if not isinstance(o, dict):
            continue

        filas.append(
            {
                "pms_archivo_id": pms_archivo_id,
                "semana": resultado.get("semana"),
                "proveedor": resultado.get("proveedor"),
                "nivel": o.get("nivel"),
                "tipo_observacion": o.get("tipo_observacion") or o.get("observacion"),
                "central": o.get("central"),
                "unidad": o.get("unidad"),
                "actividad": o.get("actividad"),
                "inspector_responsable": o.get("inspector_responsable") or o.get("inspector"),
                "fila_excel": o.get("fila_excel"),
                "campo": o.get("campo"),
                "valor_detectado": o.get("valor_detectado") or o.get("valor"),
                "sugerencia": o.get("sugerencia"),
            }
        )

    if filas:
        try:
            supabase.table("pms_observaciones").insert(filas).execute()
        except Exception as exc:
            raise HTTPException(
                status_code=500,
                detail=f"No se pudieron insertar observaciones: {exc}",
            )


def actualizar_resumen_archivo(pms_archivo_id: str, resultado: Dict[str, Any]):
    """
    Actualiza pms_archivos con el resultado final.
    """
    payload = {
        "estado_validacion": resultado.get("estado") or resultado.get("estado_validacion") or "VALIDADO",
        "errores": int(resultado.get("errores") or 0),
        "advertencias": int(resultado.get("advertencias") or 0),
        "actividades": int(resultado.get("actividades") or 0),
        "observaciones": int(resultado.get("observaciones") or 0),
    }

    if "central_presentada" in resultado:
        payload["central_presentada"] = resultado.get("central_presentada")

    if "central_presentada_norm" in resultado:
        payload["central_presentada_norm"] = resultado.get("central_presentada_norm")

    if "centrales_detectadas" in resultado:
        payload["centrales_detectadas"] = normalizar_lista(resultado.get("centrales_detectadas"))

    try:
        supabase.table("pms_archivos").update(payload).eq("id", pms_archivo_id).execute()
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo actualizar pms_archivos: {exc}",
        )


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
    pms_archivo_id = req.pms_archivo_id

    registro = obtener_registro_archivo(pms_archivo_id)
    archivo_path = registro.get("archivo_path")

    try:
        supabase.table("pms_archivos").update(
            {
                "estado_validacion": "VALIDANDO...",
                "errores": 0,
                "advertencias": 0,
                "actividades": 0,
                "observaciones": 0,
            }
        ).eq("id", pms_archivo_id).execute()
    except Exception:
        pass

    ruta_local = None

    try:
        ruta_local = descargar_archivo_storage(archivo_path)

        resultado = preparar_datos_parser(
            ruta_local,
            pms_archivo_id=pms_archivo_id,
            archivo_info=registro,
        )

        if not isinstance(resultado, dict):
            raise HTTPException(
                status_code=500,
                detail="El parser no devolvió un diccionario válido.",
            )

        resultado["pms_archivo_id"] = pms_archivo_id
        resultado.setdefault("proveedor", registro.get("proveedor"))
        resultado.setdefault("semana", registro.get("semana"))
        resultado.setdefault("archivo_path", registro.get("archivo_path"))

        borrar_resultados_previos(pms_archivo_id)
        insertar_actividades(pms_archivo_id, resultado)
        insertar_observaciones(pms_archivo_id, resultado)
        actualizar_resumen_archivo(pms_archivo_id, resultado)

        return {
            "ok": True,
            **resultado,
        }

    except HTTPException:
        raise

    except Exception as exc:
        try:
            supabase.table("pms_archivos").update(
                {
                    "estado_validacion": "ERROR EN VALIDACIÓN",
                    "errores": 1,
                }
            ).eq("id", pms_archivo_id).execute()
        except Exception:
            pass

        raise HTTPException(
            status_code=500,
            detail=f"Error validando PMS: {exc}",
        )

    finally:
        if ruta_local:
            try:
                os.remove(ruta_local)
            except Exception:
                pass


@app.post("/generar-programa-unico")
def generar_programa_unico_endpoint(req: GenerarProgramaRequest):
    """
    Genera un Excel consolidado usando la plantilla:
    PROGRAMA SEMANAL PLANTILLA.xlsx

    Recibe:
    {
      "semana": "2026-06-13",
      "central": "SANTA ROSA"
    }

    o:
    {
      "semana": "2026-06-13",
      "central": "VENTANILLA"
    }
    """
    try:
        resultado = generar_programa_unico(
            supabase=supabase,
            semana=req.semana,
            central=req.central,
        )

        archivo_generado = resultado["archivo_generado"]
        nombre_archivo = resultado["nombre_archivo"]

        return FileResponse(
            path=archivo_generado,
            filename=nombre_archivo,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo generar el programa único: {exc}",
        )
