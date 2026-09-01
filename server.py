"""
SERVIDOR WEB Y API DE INYECCIÓN UNIFICADA DE BOCETOS
Azzorti / Dupree - Campaña C3-6 2026

Zero dependencias externas.
Soporta:
  1. Carga inicial obligatoria de plantilla (.xlsm).
  2. Edición multi-hoja (BT1..BT20) con persistencia en sesión.
  3. Descarga unificada de la plantilla con todas las hojas modificadas.
"""

import http.server
import socketserver
import json
import os
import urllib.parse
import sys
import glob
import tempfile
import shutil
from excel_engine import ExcelBocetacionEngine

PORT = int(os.environ.get("PORT", 8000))
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(WORKSPACE_DIR, "PLANTILLA BOCETACION_FC_FT NVA C3-6 2026.xlsm")

class BocetacionAPIHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WORKSPACE_DIR, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        
        if parsed.path in ['/api/health', '/api/status']:
            self.send_json_response(200, {
                "status": "online",
                "default_template_exists": os.path.exists(DEFAULT_TEMPLATE),
                "template_name": os.path.basename(DEFAULT_TEMPLATE) if os.path.exists(DEFAULT_TEMPLATE) else None
            })
            return

        # Servir index.html por defecto
        if parsed.path == '/' or parsed.path == '':
            self.path = '/index.html'

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        
        if parsed.path == '/api/inyectar' or parsed.path == '/api/inyectar-unificado':
            try:
                content_type = self.headers.get('Content-Type', '')
                content_length = int(self.headers.get('Content-Length', 0))
                raw_body = self.rfile.read(content_length)

                sheets_payload = {}
                uploaded_file_bytes = None

                # Caso 1: multipart/form-data
                if 'multipart/form-data' in content_type:
                    boundary = content_type.split('boundary=')[1].encode()
                    parts = raw_body.split(b'--' + boundary)
                    
                    for part in parts:
                        if not part or part == b'--\r\n' or part == b'--':
                            continue
                        
                        header_and_data = part.split(b'\r\n\r\n', 1)
                        if len(header_and_data) < 2:
                            continue
                        
                        head, data = header_and_data
                        head_str = head.decode('utf-8', errors='ignore')
                        data = data.rstrip(b'\r\n')

                        if 'name="payload_json"' in head_str:
                            payload = json.loads(data.decode('utf-8'))
                            # Soporta formato multi-hoja { "sheets": { "BT1": {...}, "BT2": {...} } } o individual
                            if "sheets" in payload:
                                sheets_payload = payload["sheets"]
                            else:
                                target_sheet = payload.get('target_sheet', 'BT1')
                                sheets_payload = {target_sheet: {"cells": payload.get('green_cells_mapping', {})}}

                        elif 'name="excel_file"' in head_str and b'filename="' in head:
                            if len(data) > 100:
                                uploaded_file_bytes = data

                # Caso 2: application/json
                elif 'application/json' in content_type:
                    payload = json.loads(raw_body.decode('utf-8'))
                    if "sheets" in payload:
                        sheets_payload = payload["sheets"]
                    else:
                        target_sheet = payload.get('target_sheet', 'BT1')
                        sheets_payload = {target_sheet: {"cells": payload.get('green_cells_mapping', {})}}

                if not sheets_payload:
                    raise ValueError("No se enviaron datos de bocetos para inyectar.")

                # Crear archivo temporal para procesar
                with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp_file:
                    tmp_path = tmp_file.name

                if uploaded_file_bytes:
                    with open(tmp_path, "wb") as f:
                        f.write(uploaded_file_bytes)
                else:
                    if not os.path.exists(DEFAULT_TEMPLATE):
                        raise FileNotFoundError("No se encontró la plantilla predeterminada en el servidor.")
                    shutil.copyfile(DEFAULT_TEMPLATE, tmp_path)

                # Inyectar de forma unificada todas las hojas modificadas
                engine = ExcelBocetacionEngine(tmp_path)
                result = engine.inject_multiple_sheets(sheets_payload)

                if not result.get('success'):
                    raise RuntimeError("Error en la inyección de celdas.")

                # Leer archivo modificado
                with open(tmp_path, "rb") as f:
                    output_bytes = f.read()

                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

                # Enviar archivo .xlsm unificado como descarga directa
                num_sheets = len(result.get('modified_sheets', []))
                filename_download = f"PLANTILLA_BOCETACION_UNIFICADA_{num_sheets}_BOCETOS.xlsm"
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12')
                self.send_header('Content-Disposition', f'attachment; filename="{filename_download}"')
                self.send_header('Content-Length', str(len(output_bytes)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Access-Control-Expose-Headers', 'Content-Disposition, X-Modified-Sheets, X-Modified-Fields')
                self.send_header('X-Modified-Sheets', ','.join(result.get('modified_sheets', [])))
                self.send_header('X-Modified-Fields', str(result.get('total_modified_fields', 0)))
                self.end_headers()
                self.wfile.write(output_bytes)
                return

            except Exception as e:
                err_msg = str(e)
                print(f"[!] Error procesando inyección unificada: {err_msg}")
                self.send_json_response(500, {
                    "success": False,
                    "error": err_msg,
                    "detail": err_msg
                })
            return

        self.send_json_response(404, {"error": "Endpoint no encontrado"})

    def send_json_response(self, status_code, data):
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

def start_server():
    server_address = ('', PORT)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(server_address, BocetacionAPIHandler) as httpd:
        print("="*65)
        print(f"  SERVIDOR AZZORTI LISTO EN: http://localhost:{PORT}")
        print(f"  Plantilla activa: {os.path.basename(DEFAULT_TEMPLATE)}")
        print("="*65)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")

if __name__ == '__main__':
    start_server()
