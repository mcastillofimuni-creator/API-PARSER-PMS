import os
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from supabase import create_client, Client

from parser_pms import preparar_datos_parser
from generador_programa import generar_programa_unico
from generador_acta import generar_acta_interferencias
from parser_sap_ordenes import parsear_ordenes_sap
from rapidfuzz import fuzz


SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "").strip()
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "pms-archivos").strip()
FRONTEND_ORIGIN = os.getenv("FRONTEND_ORIGIN", "*").strip()
SAP_PANEL_PASSWORD = os.getenv("SAP_PANEL_PASSWORD", "2110").strip()

if not SUPABASE_URL:
    raise RuntimeError("Falta la variable de entorno SUPABASE_URL.")

if not SUPABASE_SERVICE_ROLE_KEY:
    raise RuntimeError("Falta la variable de entorno SUPABASE_SERVICE_ROLE_KEY.")


supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

app = FastAPI(
    title="PMS Parser API",
    version="1.2.0",
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


class ActaEmpresa(BaseModel):
    empresa: str = ""
    expositor: str = ""
    contrato: str = ""
    estado_validacion: str = ""
    programo: List[str] = []
    presento: List[str] = []
    archivo: str = ""


class ActaParticipante(BaseModel):
    nombre: str = ""
    contrato: str = ""
    empresa: str = ""


class ActaAccion(BaseModel):
    accion: str = ""
    responsable: str = ""


class GenerarActaRequest(BaseModel):
    semana: str
    pms: str = ""
    rango_semana: str = ""
    fecha_reunion: str = ""
    central: str
    notas: str = ""
    empresas: List[ActaEmpresa] = []
    faltantes: List[str] = []
    participantes_adicionales: List[ActaParticipante] = []
    acciones: List[ActaAccion] = []


class ValidarSapRequest(BaseModel):
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
            "/generar-acta-interferencias",
            "/cargar-maestro-sap",
            "/validar-pms-contra-sap",
            "/control-sap/validar-ots",
            "/control-sap/actualizar-ot",
        ],
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "status": "healthy",
    }


def hacer_json_serializable(valor):
    """
    Convierte valores raros de Excel/Pandas/OpenPyXL en algo que Supabase JSONB acepte.
    Evita que datos_originales reviente por fechas, objetos o NaN.
    """
    if valor is None:
        return None

    if isinstance(valor, (str, int, float, bool)):
        # Evitar NaN: JSON no lo quiere como invitado en la fiesta.
        if isinstance(valor, float) and valor != valor:
            return None
        return valor

    if isinstance(valor, dict):
        return {
            str(k): hacer_json_serializable(v)
            for k, v in valor.items()
        }

    if isinstance(valor, list):
        return [hacer_json_serializable(v) for v in valor]

    if isinstance(valor, tuple):
        return [hacer_json_serializable(v) for v in valor]

    # Fechas, decimales, objetos de Excel, etc.
    return str(valor)


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

    Ahora también guarda:
    - datos_originales jsonb

    Esto permitirá generar el programa único con más columnas de la plantilla.
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

        datos_originales = (
            a.get("datos_originales")
            or a.get("fila_original")
            or a.get("raw")
            or a.get("row_original")
            or {}
        )

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
            "datos_originales": hacer_json_serializable(datos_originales),
        }

        # Campos opcionales si existen en tu tabla.
        if "condicion" in a:
            fila["condicion"] = a.get("condicion")

        # Campo nuevo: pedido SAP detectado en columna propia o movido desde OT/Grafo
        # cuando el proveedor colocó un número tipo 3500xxxx / 4500xxxx.
        if "numero_pedido" in a:
            fila["numero_pedido"] = a.get("numero_pedido")

        if "pedido" in a and a.get("pedido") and not fila.get("numero_pedido"):
            fila["numero_pedido"] = a.get("pedido")

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


@app.post("/generar-acta-interferencias")
def generar_acta_interferencias_endpoint(req: GenerarActaRequest):
    """
    Genera un Word oficial de acta de reunión de interferencias usando:
    PLANTILLA_ACTA_INTERFERENCIAS_BASE.docx

    No usa IA externa. Rellena la plantilla con:
    - fecha
    - central
    - participantes
    - observaciones
    - acciones acordadas
    """
    try:
        resultado = generar_acta_interferencias(req.model_dump())

        archivo_generado = resultado["archivo_generado"]
        nombre_archivo = resultado["nombre_archivo"]

        return FileResponse(
            path=archivo_generado,
            filename=nombre_archivo,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo generar el acta de interferencias: {exc}",
        )



# ─── Maestro SAP OT/Avisos ───

def extraer_numeros_sap(valor: Any) -> List[str]:
    """Extrae posibles números SAP desde OT/Grafo, aviso o texto libre."""
    import re

    if valor is None:
        return []
    texto = str(valor).strip()
    if not texto:
        return []
    numeros = re.findall(r"\d{6,12}", texto)
    vistos = []
    for n in numeros:
        if n not in vistos:
            vistos.append(n)
    return vistos


def normalizar_central_backend(valor: Any) -> str:
    txt = str(valor or "").upper()
    if "SANTA ROSA" in txt:
        return "SANTA ROSA"
    if "VENTANILLA" in txt:
        return "VENTANILLA"
    return txt.strip()


def buscar_valor_en_raw(raw: Any, posibles_claves: List[str]) -> str:
    if not isinstance(raw, dict):
        return ""

    def norm(x: Any) -> str:
        import unicodedata
        t = str(x or "").upper()
        t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
        return "".join(ch if ch.isalnum() else "_" for ch in t).strip("_")

    mapa = {norm(k): v for k, v in raw.items()}
    for k in posibles_claves:
        nk = norm(k)
        if nk in mapa and mapa[nk] not in (None, ""):
            return str(mapa[nk])
    return ""


def insertar_en_lotes(tabla: str, filas: List[Dict[str, Any]], lote: int = 400):
    for i in range(0, len(filas), lote):
        supabase.table(tabla).insert(filas[i:i + lote]).execute()


@app.post("/cargar-maestro-sap")
async def cargar_maestro_sap(file: UploadFile = File(...)):
    """
    Carga el Excel SAP de Órdenes de mantenimiento y reemplaza el maestro vigente.

    Guarda datos en sap_ordenes_avisos:
    - numero_ot / descripcion_ot
    - numero_aviso / descripcion_aviso
    - central
    - estado_orden / estado_sistema / estado_control
    - objeto técnico, equipo, ubicación técnica y fechas principales
    """
    nombre = file.filename or "ordenes_sap.xlsx"
    suffix = Path(nombre).suffix or ".xlsx"
    ruta_local = None

    try:
        contenido = await file.read()
        if not contenido:
            raise HTTPException(status_code=400, detail="El archivo está vacío.")

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(contenido)
        tmp.flush()
        tmp.close()
        ruta_local = tmp.name

        resultado = parsear_ordenes_sap(ruta_local, archivo_fuente=nombre)
        registros = resultado.get("registros") or []

        if not registros:
            raise HTTPException(status_code=400, detail="No se encontraron OT/Avisos válidos en el Excel SAP.")

        # Reemplazo total del maestro. Simple y limpio para la carga semanal.
        try:
            supabase.table("sap_ordenes_avisos").delete().gte(
                "fecha_carga", "1900-01-01T00:00:00+00:00"
            ).execute()
        except Exception:
            # Fallback por si alguna fila antigua no tuviera fecha_carga.
            supabase.table("sap_ordenes_avisos").delete().neq(
                "numero_ot", "__NO_EXISTE__"
            ).execute()

        insertar_en_lotes("sap_ordenes_avisos", registros, lote=300)

        return {
            "ok": True,
            "archivo_fuente": nombre,
            **(resultado.get("resumen") or {}),
        }

    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo cargar el maestro SAP: {exc}",
        )
    finally:
        if ruta_local:
            try:
                os.remove(ruta_local)
            except Exception:
                pass


def buscar_sap_por_campo(campo: str, numero: str) -> Optional[Dict[str, Any]]:
    if not numero:
        return None
    resp = (
        supabase.table("sap_ordenes_avisos")
        .select("*")
        .eq(campo, numero)
        .limit(1)
        .execute()
    )
    data = resp.data or []
    return data[0] if data else None


def construir_observacion_sap(
    *,
    actividad: Dict[str, Any],
    req: ValidarSapRequest,
    numero: str,
    campo_origen: str,
    nivel: str,
    tipo: str,
    detalle: str,
    sugerencia: str,
    sap_row: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    return {
        "semana": req.semana,
        "central": normalizar_central_backend(req.central),
        "pms_archivo_id": actividad.get("pms_archivo_id"),
        "actividad_id": actividad.get("id"),
        "fila_excel": actividad.get("fila_excel"),
        "proveedor": actividad.get("proveedor"),
        "actividad": actividad.get("actividad"),
        "numero_informado": numero,
        "campo_origen": campo_origen,
        "nivel": nivel,
        "tipo_observacion": tipo,
        "detalle": detalle,
        "sugerencia": sugerencia,
        "numero_ot_sap": (sap_row or {}).get("numero_ot"),
        "numero_aviso_sap": (sap_row or {}).get("numero_aviso"),
        "estado_control": (sap_row or {}).get("estado_control"),
        "descripcion_sap": (sap_row or {}).get("descripcion_ot") or (sap_row or {}).get("descripcion_aviso"),
    }


@app.post("/validar-pms-contra-sap")
def validar_pms_contra_sap(req: ValidarSapRequest):
    """
    Valida las actividades PMS contra el maestro sap_ordenes_avisos ya cargado.

    Primera versión:
    - OT existe como OT.
    - Número puesto como OT pero existe como Aviso.
    - OT no encontrada.
    - OT cerrada/borrada.
    - OT de otra central.
    - Aviso informado sin OT/Grafo.
    """
    central_req = normalizar_central_backend(req.central)

    try:
        actividades_resp = (
            supabase.table("pms_actividades")
            .select("*")
            .eq("semana", req.semana)
            .execute()
        )
        actividades = actividades_resp.data or []

        actividades = [
            a for a in actividades
            if not central_req or normalizar_central_backend(a.get("central")) == central_req
        ]

        # Borra observaciones SAP previas de la misma semana/central.
        try:
            supabase.table("pms_observaciones_sap").delete().eq("semana", req.semana).eq("central", central_req).execute()
        except Exception:
            pass

        observaciones: List[Dict[str, Any]] = []
        ok_ot = 0
        ot_cerrada = 0
        ot_otra_central = 0
        aviso_como_ot = 0
        ot_no_encontrada = 0
        aviso_sin_ot = 0

        for act in actividades:
            raw = act.get("datos_originales") or {}
            ot_valor = act.get("ot_grafo") or buscar_valor_en_raw(raw, [
                "N°OT / GRAFO", "N°OT / PEDIDO", "N° OT", "OT", "ORDEN", "NºOT / GRAFO"
            ])
            aviso_valor = act.get("cod_pm_aviso") or buscar_valor_en_raw(raw, [
                "COD PM / AVISO", "COD PM / AVISO GEMA", "AVISO", "COD MP / AVISO GEMA", "COD MP / AVISO"
            ])

            numeros_ot = extraer_numeros_sap(ot_valor)
            numeros_aviso = extraer_numeros_sap(aviso_valor)

            if numeros_ot:
                for numero in numeros_ot:
                    sap_ot = buscar_sap_por_campo("numero_ot", numero)
                    if sap_ot:
                        estado = str(sap_ot.get("estado_control") or "")
                        central_sap = normalizar_central_backend(sap_ot.get("central"))

                        if estado in {"CERRADO_TEC", "BORRADO", "COMPLETADO_EMPRESA"}:
                            ot_cerrada += 1
                            observaciones.append(construir_observacion_sap(
                                actividad=act,
                                req=req,
                                numero=numero,
                                campo_origen="ot_grafo",
                                nivel="ADVERTENCIA" if estado != "BORRADO" else "ERROR",
                                tipo="OT con estado no operativo",
                                detalle=f"La OT existe en SAP, pero figura con estado {estado}.",
                                sugerencia="Verificar si corresponde programarla o si debe reemplazarse por una OT vigente/liberada.",
                                sap_row=sap_ot,
                            ))
                        elif central_req and central_sap and central_sap != central_req:
                            ot_otra_central += 1
                            observaciones.append(construir_observacion_sap(
                                actividad=act,
                                req=req,
                                numero=numero,
                                campo_origen="ot_grafo",
                                nivel="ADVERTENCIA",
                                tipo="OT pertenece a otra central",
                                detalle=f"La OT existe en SAP, pero pertenece a {central_sap}.",
                                sugerencia="Confirmar si la OT corresponde al PMS de la central filtrada.",
                                sap_row=sap_ot,
                            ))
                        else:
                            ok_ot += 1
                        continue

                    sap_aviso = buscar_sap_por_campo("numero_aviso", numero)
                    if sap_aviso:
                        aviso_como_ot += 1
                        sugerida = sap_aviso.get("numero_ot")
                        observaciones.append(construir_observacion_sap(
                            actividad=act,
                            req=req,
                            numero=numero,
                            campo_origen="ot_grafo",
                            nivel="ADVERTENCIA",
                            tipo="Aviso colocado como OT",
                            detalle="El número informado en OT/Grafo no existe como OT, pero sí existe como Aviso SAP.",
                            sugerencia=(
                                f"Revisar en SAP y reemplazar por la OT asociada {sugerida}."
                                if sugerida else
                                "Revisar en SAP la OT asociada al aviso."
                            ),
                            sap_row=sap_aviso,
                        ))
                        continue

                    ot_no_encontrada += 1
                    observaciones.append(construir_observacion_sap(
                        actividad=act,
                        req=req,
                        numero=numero,
                        campo_origen="ot_grafo",
                        nivel="ERROR",
                        tipo="Número no encontrado en SAP",
                        detalle="El número informado no fue encontrado como OT ni como Aviso en el maestro SAP cargado.",
                        sugerencia="Verificar si el número fue digitado correctamente o si falta actualizar el maestro SAP.",
                    ))

            elif numeros_aviso:
                for numero in numeros_aviso:
                    sap_aviso = buscar_sap_por_campo("numero_aviso", numero)
                    if sap_aviso:
                        aviso_sin_ot += 1
                        sugerida = sap_aviso.get("numero_ot")
                        observaciones.append(construir_observacion_sap(
                            actividad=act,
                            req=req,
                            numero=numero,
                            campo_origen="cod_pm_aviso",
                            nivel="ADVERTENCIA",
                            tipo="Aviso informado sin OT/Grafo",
                            detalle="La actividad tiene Aviso/COD PM informado, pero no tiene OT/Grafo.",
                            sugerencia=(
                                f"Completar la OT/Grafo asociada: {sugerida}."
                                if sugerida else
                                "Buscar en SAP la OT asociada al aviso."
                            ),
                            sap_row=sap_aviso,
                        ))
                    else:
                        observaciones.append(construir_observacion_sap(
                            actividad=act,
                            req=req,
                            numero=numero,
                            campo_origen="cod_pm_aviso",
                            nivel="ADVERTENCIA",
                            tipo="Aviso no encontrado en SAP",
                            detalle="El número informado como Aviso/COD PM no fue encontrado en el maestro SAP cargado.",
                            sugerencia="Verificar número o actualizar el maestro SAP.",
                        ))

        if observaciones:
            insertar_en_lotes("pms_observaciones_sap", observaciones, lote=300)

        resumen = {
            "actividades_revisadas": len(actividades),
            "ots_validas": ok_ot,
            "avisos_colocados_como_ot": aviso_como_ot,
            "ots_no_encontradas": ot_no_encontrada,
            "ots_estado_no_operativo": ot_cerrada,
            "ots_otra_central": ot_otra_central,
            "avisos_sin_ot": aviso_sin_ot,
            "observaciones_generadas": len(observaciones),
        }

        return {
            "ok": True,
            "semana": req.semana,
            "central": central_req,
            **resumen,
            "observaciones": observaciones[:50],
        }

    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo validar PMS contra SAP: {exc}",
        )


# ─── Control SAP privado ───

def _validar_password_sap(password: str):
    if str(password or "").strip() != SAP_PANEL_PASSWORD:
        raise HTTPException(status_code=401, detail="Clave incorrecta para Control SAP.")


def _indexar_registros_sap(registros: List[Dict[str, Any]]):
    por_ot: Dict[str, Dict[str, Any]] = {}
    por_aviso: Dict[str, Dict[str, Any]] = {}
    for r in registros:
        ot = str(r.get("numero_ot") or "").strip()
        av = str(r.get("numero_aviso") or "").strip()
        if ot and ot not in por_ot:
            por_ot[ot] = r
        if av and av not in por_aviso:
            por_aviso[av] = r
    return por_ot, por_aviso


def _texto_sap_para_score(r: Dict[str, Any]) -> str:
    partes = [
        r.get("descripcion_ot"),
        r.get("descripcion_aviso"),
        r.get("descripcion_objeto_tecnico"),
        r.get("equipo"),
        r.get("ubicacion_tecnica"),
        r.get("objeto_tecnico"),
    ]
    return " ".join(str(x or "") for x in partes).strip()


def _sugerir_ot_por_texto(actividad: Dict[str, Any], registros: List[Dict[str, Any]], central_req: str) -> Dict[str, Any]:
    texto_pms = " ".join(str(x or "") for x in [
        actividad.get("actividad"),
        actividad.get("unidad"),
        actividad.get("sistema"),
        actividad.get("equipo"),
    ]).strip()

    if not texto_pms:
        return {}

    mejor = None
    mejor_score = -1

    for r in registros:
        ot = str(r.get("numero_ot") or "").strip()
        if not ot:
            continue

        central_sap = normalizar_central_backend(r.get("central"))
        if central_req and central_sap and central_sap != central_req:
            # Penalizamos otra central, no la descartamos totalmente por si SAP vino con central rara.
            penalizacion = 15
        else:
            penalizacion = 0

        texto_sap = _texto_sap_para_score(r)
        if not texto_sap:
            continue

        score = int(fuzz.token_set_ratio(texto_pms.upper(), texto_sap.upper())) - penalizacion

        unidad = str(actividad.get("unidad") or "").upper().strip()
        if unidad and unidad in texto_sap.upper():
            score += 8

        if score > mejor_score:
            mejor_score = score
            mejor = r

    if not mejor:
        return {}

    return {
        "numero_ot": mejor.get("numero_ot"),
        "descripcion_ot": mejor.get("descripcion_ot") or mejor.get("descripcion_aviso") or _texto_sap_para_score(mejor),
        "numero_aviso": mejor.get("numero_aviso"),
        "estado_control": mejor.get("estado_control"),
        "central": mejor.get("central"),
        "score": max(0, min(100, mejor_score)),
    }


@app.post("/control-sap/validar-ots")
async def control_sap_validar_ots(
    password: str = Form(...),
    semana: str = Form(...),
    central: str = Form(...),
    file: UploadFile = File(...),
):
    """
    Panel privado: recibe Excel SAP, valida contra pms_actividades de la semana/central,
    y devuelve una tabla comparativa sin modificar el PMS.
    """
    _validar_password_sap(password)
    central_req = normalizar_central_backend(central)

    nombre = file.filename or "ordenes_sap.xlsx"
    suffix = Path(nombre).suffix or ".xlsx"
    ruta_local = None

    try:
        contenido = await file.read()
        if not contenido:
            raise HTTPException(status_code=400, detail="El archivo SAP está vacío.")

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(contenido)
        tmp.flush()
        tmp.close()
        ruta_local = tmp.name

        resultado_sap = parsear_ordenes_sap(ruta_local, archivo_fuente=nombre)
        registros_sap = resultado_sap.get("registros") or []
        por_ot, por_aviso = _indexar_registros_sap(registros_sap)

        actividades_resp = (
            supabase.table("pms_actividades")
            .select("*")
            .eq("semana", semana)
            .execute()
        )
        actividades = actividades_resp.data or []
        actividades = [
            a for a in actividades
            if not central_req or normalizar_central_backend(a.get("central")) == central_req
        ]

        filas = []
        resumen = {
            "actividades_revisadas": len(actividades),
            "ots_ok": 0,
            "avisos_como_ot": 0,
            "no_encontradas": 0,
            "sin_numero": 0,
            "sugeridas": 0,
            "estado_no_operativo": 0,
        }

        for act in actividades:
            raw = act.get("datos_originales") or {}
            ot_valor = act.get("ot_grafo") or buscar_valor_en_raw(raw, [
                "N°OT / GRAFO", "N°OT / PEDIDO", "N° OT", "OT", "ORDEN", "NºOT / GRAFO", "N°OT"
            ])
            aviso_valor = act.get("cod_pm_aviso") or buscar_valor_en_raw(raw, [
                "COD PM / AVISO", "COD PM / AVISO GEMA", "AVISO", "COD MP / AVISO GEMA", "COD MP / AVISO"
            ])
            aviso_detectado = ", ".join(extraer_numeros_sap(aviso_valor))

            numeros = extraer_numeros_sap(ot_valor)
            campo_origen = "OT/Grafo"
            if not numeros:
                numeros = extraer_numeros_sap(aviso_valor)
                campo_origen = "Aviso/COD PM"

            if not numeros:
                sugerida = _sugerir_ot_por_texto(act, registros_sap, central_req)
                if sugerida.get("numero_ot"):
                    resumen["sugeridas"] += 1
                else:
                    resumen["sin_numero"] += 1
                filas.append({
                    "actividad_id": act.get("id"),
                    "pms_archivo_id": act.get("pms_archivo_id"),
                    "empresa": act.get("proveedor"),
                    "fila_excel": act.get("fila_excel"),
                    "numero_pms": "",
                    "aviso_pms": aviso_detectado,
                    "campo_origen": "Sin número",
                    "actividad_pms": act.get("actividad"),
                    "unidad_pms": act.get("unidad"),
                    "sistema_pms": act.get("sistema"),
                    "equipo_pms": act.get("equipo"),
                    "estado": "SIN_NUMERO",
                    "observacion": "La actividad no tiene OT ni Aviso identificable.",
                    "ot_sap": "",
                    "aviso_sap": sugerida.get("numero_aviso") or "",
                    "descripcion_sap": "",
                    "ot_sugerida": sugerida.get("numero_ot") or "",
                    "descripcion_sugerida": sugerida.get("descripcion_ot") or "",
                    "score_sugerencia": sugerida.get("score") or 0,
                    "estado_sap": sugerida.get("estado_control") or "",
                })
                continue

            for numero in numeros:
                row_ot = por_ot.get(numero)
                row_aviso = por_aviso.get(numero)
                sugerida = {}

                if row_ot:
                    estado = str(row_ot.get("estado_control") or "")
                    if estado in {"CERRADO_TEC", "BORRADO", "COMPLETADO_EMPRESA"}:
                        resumen["estado_no_operativo"] += 1
                        estado_fila = "ESTADO_NO_OPERATIVO"
                        obs = f"La OT existe en SAP, pero figura con estado {estado}."
                    else:
                        resumen["ots_ok"] += 1
                        estado_fila = "OK"
                        obs = "La OT informada existe en SAP."

                    filas.append({
                        "actividad_id": act.get("id"),
                        "pms_archivo_id": act.get("pms_archivo_id"),
                        "empresa": act.get("proveedor"),
                        "fila_excel": act.get("fila_excel"),
                        "numero_pms": numero,
                        "aviso_pms": aviso_detectado,
                        "campo_origen": campo_origen,
                        "actividad_pms": act.get("actividad"),
                        "unidad_pms": act.get("unidad"),
                        "sistema_pms": act.get("sistema"),
                        "equipo_pms": act.get("equipo"),
                        "estado": estado_fila,
                        "observacion": obs,
                        "ot_sap": row_ot.get("numero_ot") or "",
                        "aviso_sap": row_ot.get("numero_aviso") or "",
                        "descripcion_sap": row_ot.get("descripcion_ot") or row_ot.get("descripcion_aviso") or "",
                        "ot_sugerida": row_ot.get("numero_ot") or "",
                        "descripcion_sugerida": row_ot.get("descripcion_ot") or row_ot.get("descripcion_aviso") or "",
                        "score_sugerencia": 100,
                        "estado_sap": estado,
                    })
                    continue

                if row_aviso:
                    resumen["avisos_como_ot"] += 1
                    sugerida_ot = row_aviso.get("numero_ot") or ""
                    filas.append({
                        "actividad_id": act.get("id"),
                        "pms_archivo_id": act.get("pms_archivo_id"),
                        "empresa": act.get("proveedor"),
                        "fila_excel": act.get("fila_excel"),
                        "numero_pms": numero,
                        "aviso_pms": aviso_detectado,
                        "campo_origen": campo_origen,
                        "actividad_pms": act.get("actividad"),
                        "unidad_pms": act.get("unidad"),
                        "sistema_pms": act.get("sistema"),
                        "equipo_pms": act.get("equipo"),
                        "estado": "AVISO_COMO_OT",
                        "observacion": "El número informado no existe como OT, pero sí existe como Aviso SAP.",
                        "ot_sap": "",
                        "aviso_sap": row_aviso.get("numero_aviso") or "",
                        "descripcion_sap": row_aviso.get("descripcion_aviso") or "",
                        "ot_sugerida": sugerida_ot,
                        "descripcion_sugerida": row_aviso.get("descripcion_ot") or row_aviso.get("descripcion_aviso") or "",
                        "score_sugerencia": 90 if sugerida_ot else 0,
                        "estado_sap": row_aviso.get("estado_control") or "",
                    })
                    continue

                sugerida = _sugerir_ot_por_texto(act, registros_sap, central_req)
                if sugerida.get("numero_ot"):
                    resumen["sugeridas"] += 1
                else:
                    resumen["no_encontradas"] += 1

                filas.append({
                    "actividad_id": act.get("id"),
                    "pms_archivo_id": act.get("pms_archivo_id"),
                    "empresa": act.get("proveedor"),
                    "fila_excel": act.get("fila_excel"),
                    "numero_pms": numero,
                    "aviso_pms": aviso_detectado,
                    "campo_origen": campo_origen,
                    "actividad_pms": act.get("actividad"),
                    "unidad_pms": act.get("unidad"),
                    "sistema_pms": act.get("sistema"),
                    "equipo_pms": act.get("equipo"),
                    "estado": "NO_ENCONTRADA",
                    "observacion": "El número informado no fue encontrado como OT ni como Aviso en el Excel SAP cargado.",
                    "ot_sap": "",
                    "aviso_sap": sugerida.get("numero_aviso") or "",
                    "descripcion_sap": "",
                    "ot_sugerida": sugerida.get("numero_ot") or "",
                    "descripcion_sugerida": sugerida.get("descripcion_ot") or "",
                    "score_sugerencia": sugerida.get("score") or 0,
                    "estado_sap": sugerida.get("estado_control") or "",
                })

        return {
            "ok": True,
            "semana": semana,
            "central": central_req,
            "archivo_fuente": nombre,
            "sap": resultado_sap.get("resumen") or {},
            "resumen": resumen,
            "filas": filas,
        }

    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"No se pudo validar OTs en Control SAP: {exc}")
    finally:
        if ruta_local:
            try:
                os.remove(ruta_local)
            except Exception:
                pass


class ActualizarOtControlSapRequest(BaseModel):
    password: str
    actividad_id: str
    ot_nueva: str


@app.post("/control-sap/actualizar-ot")
def control_sap_actualizar_ot(req: ActualizarOtControlSapRequest):
    """Actualiza manualmente la OT/Grafo de una actividad PMS."""
    _validar_password_sap(req.password)
    ot = str(req.ot_nueva or "").strip()
    if not ot:
        raise HTTPException(status_code=400, detail="Debes ingresar una OT válida.")
    if not req.actividad_id:
        raise HTTPException(status_code=400, detail="No se recibió actividad_id.")

    try:
        resp = (
            supabase.table("pms_actividades")
            .update({"ot_grafo": ot})
            .eq("id", req.actividad_id)
            .execute()
        )
        return {"ok": True, "actividad_id": req.actividad_id, "ot_nueva": ot, "data": resp.data}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"No se pudo actualizar la OT: {exc}")

