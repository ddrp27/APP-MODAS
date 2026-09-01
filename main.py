"""
API FASTAPI PARA DESPLIEGUE EN RENDER.COM
Azzorti / Dupree - Campaña C3-6 2026

Endpoints:
  - GET / : Sirve la aplicación web index.html
  - GET /api/health : Monitoreo de estado
  - POST /api/inyectar-unificado : Inyección quirúrgica multi-hoja y streaming de descarga
"""

import os
import json
import tempfile
import shutil
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from excel_engine import ExcelBocetacionEngine

app = FastAPI(title="Azzorti Bocetacion API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(BASE_DIR, "PLANTILLA BOCETACION_FC_FT NVA C3-6 2026.xlsm")

@app.get("/", response_class=HTMLResponse)
async def read_index():
    index_path = os.path.join(BASE_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Azzorti Bocetacion API Online</h1>"

@app.get("/api/health")
async def health_check():
    return {
        "status": "online",
        "default_template_exists": os.path.exists(DEFAULT_TEMPLATE)
    }

@app.post("/api/inyectar")
@app.post("/api/inyectar-unificado")
async def inyectar_unificado(
    payload_json: str = Form(...),
    excel_file: UploadFile = File(None)
):
    try:
        data = json.loads(payload_json)
        sheets_payload = {}

        if "sheets" in data:
            sheets_payload = data["sheets"]
        else:
            target_sheet = data.get('target_sheet', 'BT1')
            sheets_payload = {target_sheet: {"cells": data.get('green_cells_mapping', {})}}

        if not sheets_payload:
            raise HTTPException(status_code=400, detail="No se enviaron datos de bocetos para inyectar.")

        with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp_file:
            tmp_path = tmp_file.name

        if excel_file:
            content = await excel_file.read()
            with open(tmp_path, "wb") as f:
                f.write(content)
        else:
            if not os.path.exists(DEFAULT_TEMPLATE):
                raise HTTPException(status_code=404, detail="Plantilla base no encontrada en el servidor.")
            shutil.copyfile(DEFAULT_TEMPLATE, tmp_path)

        engine = ExcelBocetacionEngine(tmp_path)
        result = engine.inject_multiple_sheets(sheets_payload)

        if not result.get('success'):
            raise HTTPException(status_code=500, detail="Error en la inyección de celdas.")

        num_sheets = len(result.get('modified_sheets', []))
        filename_download = f"PLANTILLA_BOCETACION_UNIFICADA_{num_sheets}_BOCETOS.xlsm"

        return FileResponse(
            path=tmp_path,
            filename=filename_download,
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            headers={
                "X-Modified-Sheets": ",".join(result.get('modified_sheets', [])),
                "X-Modified-Fields": str(result.get('total_modified_fields', 0))
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
