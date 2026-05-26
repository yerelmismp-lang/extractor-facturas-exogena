"""
Extractor de Facturas - API Web
Informacion Exogena 2025 - Formato 1001
"""
import io
import uuid
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

from extractor_facturas import generate_excel, iter_folder

app = FastAPI(title="Extractor de Facturas - Exogena 2025")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Cache en memoria: token -> (excel_bytes, expira_en)
_CACHE: Dict[str, tuple] = {}


def _limpiar_cache():
    now = datetime.now()
    viejos = [k for k, (_, exp) in _CACHE.items() if now > exp]
    for k in viejos:
        del _CACHE[k]


@app.get("/", response_class=HTMLResponse)
async def index():
    html = Path("static/index.html")
    if html.exists():
        return HTMLResponse(html.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>Error: static/index.html no encontrado</h1>", status_code=500)


@app.post("/procesar")
async def procesar(archivos: List[UploadFile] = File(...)):
    """Recibe facturas, extrae datos, genera Excel y devuelve estadisticas + token."""
    _limpiar_cache()

    if not archivos:
        raise HTTPException(status_code=400, detail="No se recibieron archivos.")

    EXTENSIONES_VALIDAS = {".pdf", ".xml", ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"}

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        # Guardar archivos subidos
        guardados = 0
        for f in archivos:
            if not f.filename:
                continue
            nombre = Path(f.filename).name  # Proteccion contra path traversal
            ext = Path(nombre).suffix.lower()
            if ext not in EXTENSIONES_VALIDAS:
                continue
            contenido = await f.read()
            (tmp / nombre).write_bytes(contenido)
            guardados += 1

        if guardados == 0:
            raise HTTPException(
                status_code=400,
                detail="Ninguno de los archivos es un formato valido. Use PDF, XML, JPG o PNG.",
            )

        # Procesar facturas
        records = []
        for rec, _idx, _total in iter_folder(tmp, year_filter=2025):
            records.append(rec)

        if not records:
            raise HTTPException(
                status_code=400,
                detail="No se pudieron extraer datos de los archivos enviados.",
            )

        # Generar Excel en memoria
        excel_path = tmp / "exogena_2025_F1001.xlsx"
        generate_excel(records, excel_path)
        excel_bytes = excel_path.read_bytes()

    # Calcular estadisticas
    n_ok      = sum(1 for r in records if r.get("confianza") in ("Alta", "Media"))
    n_revisar = sum(1 for r in records if r.get("confianza") in ("Baja", "Revisar"))
    n_error   = sum(1 for r in records if r.get("confianza") == "Error")
    base_tot  = sum((r.get("base")  or 0) for r in records if r.get("confianza") != "Error")
    iva_tot   = sum((r.get("iva")   or 0) for r in records if r.get("confianza") != "Error")
    total_tot = sum((r.get("total") or 0) for r in records if r.get("confianza") != "Error")

    # Guardar en cache (expira en 15 minutos)
    token = str(uuid.uuid4())
    _CACHE[token] = (excel_bytes, datetime.now() + timedelta(minutes=15))

    return JSONResponse({
        "ok": True,
        "token": token,
        "stats": {
            "total":   len(records),
            "ok":      n_ok,
            "revisar": n_revisar,
            "error":   n_error,
            "base":    round(base_tot, 0),
            "iva":     round(iva_tot, 0),
            "valor":   round(total_tot, 0),
        },
        "registros": [
            {
                "archivo":   r["archivo"],
                "proveedor": r.get("proveedor") or "",
                "nit":       r.get("nit") or "",
                "fecha":     r.get("fecha") or "",
                "total":     r.get("total") or 0,
                "confianza": r.get("confianza") or "",
                "error":     r.get("error") or "",
            }
            for r in records
        ],
    })


@app.get("/descargar/{token}")
async def descargar(token: str):
    """Descarga el Excel generado usando el token recibido en /procesar."""
    _limpiar_cache()
    entrada = _CACHE.get(token)
    if not entrada:
        raise HTTPException(
            status_code=404,
            detail="El archivo expiro o no existe. Vuelva a procesar las facturas.",
        )
    excel_bytes, _ = entrada
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="exogena_2025_F1001_{ts}.xlsx"'
        },
    )
