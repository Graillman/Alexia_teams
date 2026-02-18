from http.server import BaseHTTPRequestHandler
import json, io, copy, cgi, os, hashlib, hmac, time

from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml

# ── Sécurité ────────────────────────────────────────────────────────────────
MAX_FILE_SIZE = 20 * 1024 * 1024   # 20 Mo max
ALLOWED_MIME  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

def is_valid_docx(data: bytes) -> bool:
    """Vérifie la signature ZIP et la présence de word/document.xml."""
    if len(data) < 4:
        return False
    # Signature ZIP : PK\x03\x04
    if data[:4] != b'PK\x03\x04':
        return False
    # Vérifie que c'est bien un .docx (contient word/)
    try:
        from zipfile import ZipFile
        with ZipFile(io.BytesIO(data)) as z:
            names = z.namelist()
            return 'word/document.xml' in names
    except Exception:
        return False

def sanitize_filename(name: str) -> str:
    """Retire les caractères dangereux du nom de fichier."""
    import re
    name = os.path.basename(name)
    name = re.sub(r'[^\w\-. ]', '_', name)
    return name[:100] or 'document.docx'


# ── Correcteur principal ─────────────────────────────────────────────────────
def fix_for_teams(input_bytes: bytes) -> bytes:
    """
    Corrige un .docx pour un affichage parfait sur Teams :
    1. Convertit les images flottantes (anchor) en inline
    2. Réduit les espacements excessifs
    3. Supprime les modes de compatibilité anciens
    """
    buf = io.BytesIO(input_bytes)
    doc = Document(buf)
    body = doc.element.body

    # ── 1. Images flottantes → inline ────────────────────────────────────────
    anchors = body.findall('.//' + qn('wp:anchor'))
    for anchor in anchors:
        drawing = anchor.getparent()
        if drawing is None or drawing.tag != qn('w:drawing'):
            continue

        graphic = anchor.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
        if graphic is None:
            continue

        extent  = anchor.find(qn('wp:extent'))
        cx = extent.get('cx', '2000000') if extent is not None else '2000000'
        cy = extent.get('cy', '1500000') if extent is not None else '1500000'
        docPr  = anchor.find(qn('wp:docPr'))

        inline = parse_xml(
            f'<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            f' distT="0" distB="0" distL="114240" distR="114240">'
            f'<wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'</wp:inline>'
        )
        if docPr is not None:
            inline.append(copy.deepcopy(docPr))
        inline.append(parse_xml(
            '<wp:cNvGraphicFramePr xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
            '</wp:cNvGraphicFramePr>'
        ))
        inline.append(copy.deepcopy(graphic))

        drawing.remove(anchor)
        drawing.append(inline)

    # ── 2. Espacements excessifs ──────────────────────────────────────────────
    for p in body.iter(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            continue
        for attr in [qn('w:before'), qn('w:after')]:
            val = spacing.get(attr)
            if val and val.lstrip('-').isdigit() and int(val) > 400:
                spacing.set(attr, '200')
        line     = spacing.get(qn('w:line'))
        lineRule = spacing.get(qn('w:lineRule'))
        if line and line.lstrip('-').isdigit():
            if lineRule in ('exact', 'atLeast') and int(line) > 500:
                spacing.set(qn('w:line'), '276')
                spacing.set(qn('w:lineRule'), 'auto')

    # ── 3. Compatibilité ─────────────────────────────────────────────────────
    try:
        compat = doc.settings.element.find(qn('w:compat'))
        if compat is not None:
            for bad in ['useWord2002TableStyleRules', 'useWord97LineBreakRules', 'growAutofit']:
                el = compat.find(qn(f'w:{bad}'))
                if el is not None:
                    compat.remove(el)
    except Exception:
        pass

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ── Handler Vercel ───────────────────────────────────────────────────────────
class handler(BaseHTTPRequestHandler):

    def do_POST(self):
        if self.path != '/api/fix':
            self._send(404, {'error': 'Not found'})
            return

        # ── Vérif taille avant de lire ────────────────────────────────────
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length > MAX_FILE_SIZE:
            self._send(413, {'error': f'Fichier trop lourd (max 20 Mo).'})
            return

        content_type = self.headers.get('Content-Type', '')
        body_bytes   = self.rfile.read(content_length)

        # ── Parse multipart ───────────────────────────────────────────────
        fs = cgi.FieldStorage(
            fp=io.BytesIO(body_bytes),
            environ={
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE':   content_type,
                'CONTENT_LENGTH': str(content_length),
            },
            keep_blank_values=True
        )

        if 'file' not in fs:
            self._send(400, {'error': 'Aucun fichier reçu.'})
            return

        file_item  = fs['file']
        if not hasattr(file_item, 'file'):
            self._send(400, {'error': 'Fichier invalide.'})
            return

        file_bytes = file_item.file.read()

        # ── Vérifications de sécurité ─────────────────────────────────────
        if len(file_bytes) > MAX_FILE_SIZE:
            self._send(413, {'error': 'Fichier trop lourd (max 20 Mo).'})
            return

        if not is_valid_docx(file_bytes):
            self._send(400, {'error': 'Le fichier n\'est pas un .docx valide.'})
            return

        # ── Traitement ────────────────────────────────────────────────────
        try:
            fixed_bytes = fix_for_teams(file_bytes)
        except Exception as e:
            self._send(500, {'error': f'Erreur lors du traitement : {str(e)}'})
            return

        # ── Réponse sécurisée ─────────────────────────────────────────────
        raw_name   = getattr(file_item, 'filename', 'document.docx') or 'document.docx'
        clean_name = sanitize_filename(raw_name).replace('.docx', '_teams.docx')

        self.send_response(200)
        self.send_header('Content-Type', ALLOWED_MIME)
        self.send_header('Content-Disposition', f'attachment; filename="{clean_name}"')
        self.send_header('Content-Length', str(len(fixed_bytes)))
        # Headers de sécurité
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('X-Frame-Options', 'DENY')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(fixed_bytes)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def _send(self, code, data):
        b = json.dumps(data).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(b)))
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(b)

    def log_message(self, *args):
        pass
