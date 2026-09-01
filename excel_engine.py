"""
MOTOR DE INYECCIÓN DE EXCEL MULTI-HOJA, RENOMBRADO & INSERCIÓN DE IMÁGENES OPENXML
Azzorti / Dupree - Campaña C3-6 2026

Zonas de Inyección de Imágenes:
  - F5:O21   : Dibujo Plano (Geometral Vectorial)
  - P5:Z21   : Dibujo con Color (Boceto / Render)
  - AI9:AO21 : Imagen de Referentes / Muestras
  - B37:I42  : Foto de Modelación (Frente)
"""

import os
import sys
import zipfile
import re
import base64
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
import shutil

IMAGE_ANCHORS = {
    'dibujo_plano': {
        'from_col': 5,  'from_row': 4,
        'to_col': 15,   'to_row': 21,
        'name': 'Dibujo Plano (F5:O21)'
    },
    'dibujo_color': {
        'from_col': 15, 'from_row': 4,
        'to_col': 26,   'to_row': 21,
        'name': 'Dibujo con Color (P5:Z21)'
    },
    'referente': {
        'from_col': 34, 'from_row': 8,
        'to_col': 41,   'to_row': 21,
        'name': 'Imagen Referente (AI9:AO21)'
    },
    'foto': {
        'from_col': 1,  'from_row': 36,
        'to_col': 9,    'to_row': 42,
        'name': 'Foto Modelacion (B37:I42)'
    }
}

class ExcelBocetacionEngine:
    def __init__(self, excel_path):
        self.excel_path = os.path.abspath(excel_path)
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"No se encontró el archivo: {self.excel_path}")

    def inject_data(self, target_sheet_name, green_cells, images=None):
        return self.inject_multiple_sheets({target_sheet_name: {"cells": green_cells, "images": images}})

    def inject_multiple_sheets(self, sheets_dict):
        temp_zip_path = self.excel_path + ".temp.zip"
        total_modified = 0
        total_images = 0
        modified_sheets = []
        renamed_sheets = {}

        with zipfile.ZipFile(self.excel_path, 'r') as z_in:
            wb_xml_str = z_in.read('xl/workbook.xml').decode('utf-8')
            app_xml_str = z_in.read('docProps/app.xml').decode('utf-8') if 'docProps/app.xml' in z_in.namelist() else None
            content_types_str = z_in.read('[Content_Types].xml').decode('utf-8')
            
            wb_root = ET.fromstring(wb_xml_str.encode('utf-8'))
            wb_rels_root = ET.fromstring(z_in.read('xl/_rels/workbook.xml.rels'))
            
            ns_main = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            ns_rels = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            rels_map = {r.attrib['Id']: r.attrib['Target'] for r in wb_rels_root.findall('.//rel:Relationship', ns_rels)}
            
            sheet_to_path = {}
            for s in wb_root.findall('.//main:sheet', ns_main):
                name = s.attrib.get('name')
                rId = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                sheet_to_path[name] = 'xl/' + rels_map.get(rId, '')

            updated_files = {} # { filename: bytes }
            new_drawings = {}   # { drawing_path: (drawing_xml_bytes, drawing_rels_bytes, { media_path: media_bytes }) }

            for original_sheet_name, sheet_data in sheets_dict.items():
                target_path = sheet_to_path.get(original_sheet_name)
                if not target_path or target_path not in z_in.namelist():
                    print(f"[!] Advertencia: Hoja '{original_sheet_name}' no encontrada. Omitiendo.")
                    continue

                cells = sheet_data.get('cells', {})
                images = sheet_data.get('images', {})
                
                # 1. AN1 es la Referencia de Producto (renombra la pestaña en Excel)
                ref_val = str(cells.get('AN1', cells.get('AK1', ''))).strip()
                if ref_val:
                    cells['AN1'] = ref_val
                    clean_sheet_name = re.sub(r'[:\\/\?\*\[\]]', '_', ref_val)[:31].strip()
                    if clean_sheet_name:
                        renamed_sheets[original_sheet_name] = clean_sheet_name

                sheet_bytes = z_in.read(target_path)
                sheet_count = 0

                # 2. Inyección de Celdas
                for coord, val in cells.items():
                    if val is not None and str(val).strip() != "":
                        sheet_bytes = self._update_single_cell(sheet_bytes, coord, str(val))
                        sheet_count += 1

                # 3. Inyección de Imágenes OpenXML
                sheet_images_count = 0
                active_images = {k: v for k, v in images.items() if v and isinstance(v, str) and v.startswith('data:image')}
                
                if active_images:
                    sheet_basename = os.path.basename(target_path) # e.g. sheet10.xml
                    sheet_rels_path = f"xl/worksheets/_rels/{sheet_basename}.rels"
                    
                    sheet_rels_str = z_in.read(sheet_rels_path).decode('utf-8') if sheet_rels_path in z_in.namelist() else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
                    sheet_xml_str = sheet_bytes.decode('utf-8')

                    drawing_filename = f"drawing_{original_sheet_name}.xml"
                    drawing_full_path = f"xl/drawings/{drawing_filename}"
                    drawing_rels_full_path = f"xl/drawings/_rels/{drawing_filename}.rels"

                    anchors_xml = []
                    drawing_rels_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
                    sheet_media = {}

                    img_idx = 1
                    for img_key, b64_str in active_images.items():
                        if img_key not in IMAGE_ANCHORS:
                            continue
                        
                        anchor_def = IMAGE_ANCHORS[img_key]
                        raw_bytes, ext = self._decode_base64_image(b64_str)
                        if not raw_bytes:
                            continue

                        media_rel_path = f"media/img_{original_sheet_name}_{img_key}.{ext}"
                        media_full_path = f"xl/{media_rel_path}"
                        sheet_media[media_full_path] = raw_bytes

                        rel_id = f"rId{img_idx}"
                        drawing_rels_xml.append(f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../{media_rel_path}"/>')

                        anchor_snippet = f'''<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from>
    <xdr:col>{anchor_def['from_col']}</xdr:col>
    <xdr:colOff>25400</xdr:colOff>
    <xdr:row>{anchor_def['from_row']}</xdr:row>
    <xdr:rowOff>25400</xdr:rowOff>
  </xdr:from>
  <xdr:to>
    <xdr:col>{anchor_def['to_col']}</xdr:col>
    <xdr:colOff>0</xdr:colOff>
    <xdr:row>{anchor_def['to_row']}</xdr:row>
    <xdr:rowOff>0</xdr:rowOff>
  </xdr:to>
  <xdr:pic>
    <xdr:nvPicPr>
      <xdr:cNvPr id="{img_idx + 20}" name="{anchor_def['name']}"/>
      <xdr:cNvPicPr><a:picLocks noChangeAspect="0"/></xdr:cNvPicPr>
    </xdr:nvPicPr>
    <xdr:blipFill>
      <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rel_id}"/>
      <a:stretch><a:fillRect/></a:stretch>
    </xdr:blipFill>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    </xdr:spPr>
  </xdr:pic>
  <xdr:clientData/>
</xdr:twoCellAnchor>'''
                        anchors_xml.append(anchor_snippet)
                        img_idx += 1
                        sheet_images_count += 1

                    drawing_rels_xml.append('</Relationships>')
                    
                    drawing_xml_content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
{''.join(anchors_xml)}
</xdr:wsDr>'''

                    # Actualizar worksheet rels
                    if 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"' not in sheet_rels_str:
                        drawing_rel_tag = f'<Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/{drawing_filename}"/>'
                        sheet_rels_str = sheet_rels_str.replace('</Relationships>', drawing_rel_tag + '</Relationships>')
                        updated_files[sheet_rels_path] = sheet_rels_str.encode('utf-8')

                    # Actualizar worksheet XML para enlazar el drawing
                    if '<drawing ' not in sheet_xml_str:
                        if '<legacyDrawing' in sheet_xml_str:
                            sheet_xml_str = sheet_xml_str.replace('<legacyDrawing', '<drawing r:id="rIdDrawing"/><legacyDrawing')
                        elif '<extLst' in sheet_xml_str:
                            sheet_xml_str = sheet_xml_str.replace('<extLst', '<drawing r:id="rIdDrawing"/><extLst')
                        else:
                            sheet_xml_str = sheet_xml_str.replace('</worksheet>', '<drawing r:id="rIdDrawing"/></worksheet>')
                        sheet_bytes = sheet_xml_str.encode('utf-8')

                    # Actualizar Content_Types
                    if f'PartName="/xl/drawings/{drawing_filename}"' not in content_types_str:
                        override_tag = f'<Override PartName="/xl/drawings/{drawing_filename}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                        content_types_str = content_types_str.replace('</Types>', override_tag + '</Types>')

                    new_drawings[drawing_full_path] = (
                        drawing_xml_content.encode('utf-8'),
                        ''.join(drawing_rels_xml).encode('utf-8'),
                        sheet_media
                    )

                if sheet_count > 0 or sheet_images_count > 0:
                    updated_files[target_path] = sheet_bytes
                    total_modified += sheet_count
                    total_images += sheet_images_count
                    display_name = renamed_sheets.get(original_sheet_name, original_sheet_name)
                    modified_sheets.append(f"{original_sheet_name} -> {display_name}" if original_sheet_name != display_name else original_sheet_name)

            # Asegurar extensiones de imagen en [Content_Types].xml
            if '<Default Extension="png"' not in content_types_str:
                content_types_str = content_types_str.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/>')
            if '<Default Extension="jpeg"' not in content_types_str and '<Default Extension="jpg"' not in content_types_str:
                content_types_str = content_types_str.replace('</Types>', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/></Types>')

            # Renombrar en workbook.xml y app.xml
            for orig_name, new_name in renamed_sheets.items():
                wb_xml_str = re.sub(rf'name="{orig_name}"', f'name="{new_name}"', wb_xml_str)
                wb_xml_str = re.sub(rf"'{orig_name}'!", f"'{new_name}'!", wb_xml_str)
                wb_xml_str = re.sub(rf"\b{orig_name}!", f"'{new_name}'!", wb_xml_str)
                if app_xml_str:
                    app_xml_str = re.sub(rf'<vt:lpstr>{orig_name}</vt:lpstr>', f'<vt:lpstr>{new_name}</vt:lpstr>', app_xml_str)

            # Escribir el nuevo archivo ZIP
            with zipfile.ZipFile(temp_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
                for item in z_in.infolist():
                    if item.filename in updated_files:
                        z_out.writestr(item, updated_files[item.filename])
                    elif item.filename == 'xl/workbook.xml':
                        z_out.writestr(item, wb_xml_str.encode('utf-8'))
                    elif item.filename == 'docProps/app.xml' and app_xml_str:
                        z_out.writestr(item, app_xml_str.encode('utf-8'))
                    elif item.filename == '[Content_Types].xml':
                        z_out.writestr(item, content_types_str.encode('utf-8'))
                    else:
                        z_out.writestr(item, z_in.read(item.filename))

                # Escribir nuevos drawings y archivos multimedia
                for drawing_path, (drw_bytes, drw_rels_bytes, media_dict) in new_drawings.items():
                    z_out.writestr(drawing_path, drw_bytes)
                    drawing_rels_path = drawing_path.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels'
                    z_out.writestr(drawing_rels_path, drw_rels_bytes)
                    for m_path, m_bytes in media_dict.items():
                        z_out.writestr(m_path, m_bytes)

        shutil.move(temp_zip_path, self.excel_path)

        return {
            "success": True,
            "engine": "Lossless OpenXML Multi-Sheet & Image Engine",
            "modified_sheets": modified_sheets,
            "renamed_sheets": renamed_sheets,
            "total_modified_fields": total_modified,
            "total_embedded_images": total_images,
            "message": f"Inyección unificada completada con éxito. Se modificaron {len(modified_sheets)} hojas y se incrustaron {total_images} imágenes."
        }

    def _decode_base64_image(self, b64_str):
        try:
            ext = 'png'
            if 'image/jpeg' in b64_str or 'image/jpg' in b64_str:
                ext = 'jpeg'
            
            if ',' in b64_str:
                b64_str = b64_str.split(',', 1)[1]
            
            raw = base64.b64decode(b64_str)
            return raw, ext
        except Exception as e:
            print(f"[!] Error decodificando imagen base64: {e}")
            return None, None

    def _update_single_cell(self, xml_bytes, cell_coord, new_val):
        xml_str = xml_bytes.decode('utf-8')
        raw_val = str(new_val).strip()
        
        is_numeric = False
        if re.match(r'^-?\d+(\.\d+)?$', raw_val):
            if cell_coord in ['AD1', 'Q24', 'S24', 'U24', 'AK24', 'AM24', 'AO24', 'AW24', 'AS24', 'AS25', 'AS26', 'AS27', 'AS28', 'AS29', 'AS30', 'AS31', 'AS32', 'AS33']:
                is_numeric = True

        cell_pattern = rf'(<c\s+r="{cell_coord}"(?:\s+[^/>]*)?)(?:/>|>(?:.*?)</c>)'
        match = re.search(cell_pattern, xml_str, flags=re.DOTALL)
        
        if is_numeric:
            if match:
                tag_start = match.group(1).rstrip('/')
                tag_clean = re.sub(r'\s+t="[^"]*"', '', tag_start)
                new_cell_tag = f'{tag_clean}><v>{raw_val}</v></c>'
                xml_str = xml_str[:match.start()] + new_cell_tag + xml_str[match.end():]
            else:
                col, row = re.match(r'([A-Z]+)([0-9]+)', cell_coord).groups()
                row_pattern = rf'(<row\s+r="{row}"(?:\s+[^>]*)?>)(.*?)(</row>)'
                row_match = re.search(row_pattern, xml_str, flags=re.DOTALL)
                if row_match:
                    row_start, row_content, row_end = row_match.groups()
                    new_cell = f'<c r="{cell_coord}"><v>{raw_val}</v></c>'
                    xml_str = xml_str[:row_match.start()] + row_start + row_content + new_cell + row_end + xml_str[row_match.end():]
        else:
            escaped_val = saxutils.escape(raw_val)
            if match:
                tag_start = match.group(1).rstrip('/')
                tag_clean = re.sub(r'\s+t="[^"]*"', '', tag_start)
                new_cell_tag = f'{tag_clean} t="inlineStr"><is><t>{escaped_val}</t></is></c>'
                xml_str = xml_str[:match.start()] + new_cell_tag + xml_str[match.end():]
            else:
                col, row = re.match(r'([A-Z]+)([0-9]+)', cell_coord).groups()
                row_pattern = rf'(<row\s+r="{row}"(?:\s+[^>]*)?>)(.*?)(</row>)'
                row_match = re.search(row_pattern, xml_str, flags=re.DOTALL)
                if row_match:
                    row_start, row_content, row_end = row_match.groups()
                    new_cell = f'<c r="{cell_coord}" t="inlineStr"><is><t>{escaped_val}</t></is></c>'
                    xml_str = xml_str[:row_match.start()] + row_start + row_content + new_cell + row_end + xml_str[row_match.end():]

        return xml_str.encode('utf-8')
