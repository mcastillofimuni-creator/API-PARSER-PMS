import os
import tempfile
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from supabase import create_client, Client

from parser_pms import preparar_datos_parser

# ============================================================
# VARIABLES DE ENTORNO
# Estas se configurarán luego en Render, NO en GitHub.
# ============================================================

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "pms-archivos")
FRONTEND_ORIGIN = os.getenv("FRONTEND_ORIGIN", "*")

if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
    raise RuntimeError("Faltan variables SUPABASE_URL o SUPABASE_SERVICE_ROLE_KEY")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

# ============================================================
# API
# ============================================================

app = FastAPI(
    title="PMS Parser API",
    version="1.0.0",
    description="API para validar PMS semanales y guardar actividades/observaciones en Supabase."
)

origins = ["*"] if FRONTEND_ORIGIN == "*" else [FRONTEND_ORIGIN]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ValidarRequest(BaseModel):
    pms_archivo_id: str


@app.get("/")
def root():
    return {
        "ok": True,
        "service": "PMS Parser API",
        "message": "API activa"
    }


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/validar-pms")
def validar_pms(req: ValidarRequest):
    pms_id = req.pms_archivo_id

    try:
        # ------------------------------------------------------------
        # 1. Buscar el registro del PMS en Supabase
        # ------------------------------------------------------------
        registro_resp = (
            supabase
            .table("pms_archivos")
            .select("*")
            .eq("id", pms_id)
            .single()
            .execute()
        )

        registro = registro_resp.data

        if not registro:
            raise HTTPException(
                status_code=404,
                detail="No se encontró el PMS solicitado en pms_archivos"
            )

        archivo_path = registro.get("archivo_path")
        semana = registro.get("semana")
        proveedor = registro.get("proveedor")

        if not archivo_path:
            raise HTTPException(
                status_code=400,
                detail="El PMS no tiene archivo_path asociado"
            )

        # ------------------------------------------------------------
        # 2. Marcar estado como EN PROCESO
        # ------------------------------------------------------------
        supabase.table("pms_archivos").update({
            "estado_validacion": "EN PROCESO",
            "errores": 0,
            "advertencias": 0
        }).eq("id", pms_id).execute()

        # ------------------------------------------------------------
        # 3. Descargar archivo desde Supabase Storage
        # ------------------------------------------------------------
        contenido = supabase.storage.from_(SUPABASE_BUCKET).download(archivo_path)

        if not contenido:
            raise HTTPException(
                status_code=500,
                detail="No se pudo descargar el archivo desde Supabase Storage"
            )

        nombre_archivo = os.path.basename(archivo_path)

        # ------------------------------------------------------------
        # 4. Guardar temporalmente y ejecutar parser
        # ------------------------------------------------------------
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_local = os.path.join(tmpdir, nombre_archivo)

            with open(ruta_local, "wb") as f:
                f.write(contenido)

            resultado = preparar_datos_parser(ruta_local)

        actividades = resultado.get("actividades", [])
        observaciones = resultado.get("observaciones", [])
        errores = int(resultado.get("errores", 0) or 0)
        advertencias = int(resultado.get("advertencias", 0) or 0)
        estado = resultado.get("estado", "VALIDADO")

        # ------------------------------------------------------------
        # 5. Limpiar resultados anteriores del mismo PMS
        # ------------------------------------------------------------
        supabase.table("pms_actividades").delete().eq("pms_archivo_id", pms_id).execute()
        supabase.table("pms_observaciones").delete().eq("pms_archivo_id", pms_id).execute()

        # ------------------------------------------------------------
        # 6. Insertar actividades
        # ------------------------------------------------------------
        actividades_insert = []

        for a in actividades:
            actividades_insert.append({
                "pms_archivo_id": pms_id,
                "semana": semana,
                "proveedor": proveedor,
                "fila_excel": a.get("fila_excel"),
                "unidad": a.get("unidad"),
                "sistema": a.get("sistema"),
                "equipo": a.get("equipo"),
                "actividad": a.get("actividad"),
                "ot_grafo": a.get("ot_grafo"),
                "tipo_mant": a.get("tipo_mant"),
                "riesgo": a.get("riesgo"),
                "inspector": a.get("inspector"),
                "rt_terceros": a.get("rt_terceros"),
            })

        if actividades_insert:
            supabase.table("pms_actividades").insert(actividades_insert).execute()

        # ------------------------------------------------------------
        # 7. Insertar observaciones
        # ------------------------------------------------------------
        observaciones_insert = []

        for o in observaciones:
            observaciones_insert.append({
                "pms_archivo_id": pms_id,
                "semana": semana,
                "proveedor": proveedor,
                "nivel": o.get("nivel"),
                "tipo_observacion": o.get("tipo_observacion"),
                "unidad": o.get("unidad"),
                "actividad": o.get("actividad"),
                "inspector_responsable": o.get("inspector_responsable"),
                "fila_excel": o.get("fila_excel"),
                "campo": o.get("campo"),
                "valor_detectado": o.get("valor_detectado"),
                "sugerencia": o.get("sugerencia"),
            })

        if observaciones_insert:
            supabase.table("pms_observaciones").insert(observaciones_insert).execute()

        # ------------------------------------------------------------
        # 8. Actualizar cabecera del PMS
        # ------------------------------------------------------------
        supabase.table("pms_archivos").update({
            "estado_validacion": estado,
            "errores": errores,
            "advertencias": advertencias
        }).eq("id", pms_id).execute()

        # ------------------------------------------------------------
        # 9. Respuesta de API
        # ------------------------------------------------------------
        return {
            "ok": True,
            "pms_archivo_id": pms_id,
            "proveedor": proveedor,
            "semana": semana,
            "archivo_path": archivo_path,
            "estado": estado,
            "errores": errores,
            "advertencias": advertencias,
            "actividades": len(actividades_insert),
            "observaciones": len(observaciones_insert),
            "detalle_observaciones": observaciones[:50]
        }

    except HTTPException:
        raise

    except Exception as e:
        # Si algo falla, marcar el PMS como ERROR DE VALIDACION
        try:
            supabase.table("pms_archivos").update({
                "estado_validacion": "ERROR DE VALIDACION",
                "errores": 1,
                "advertencias": 0
            }).eq("id", pms_id).execute()
        except Exception:
            pass

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
