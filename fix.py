from http.server import BaseHTTPRequestHandler
import json
import io
import copy
from zipfile import ZipFile
import re
import cgi

try:
    from lxml import etree
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False


# ── XML namespaces ──────────────────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'v':   'urn:schemas-microsoft-com:vml',
}

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

def tag(ns, local): return f'{{{ns}}}{local}'


# ── Core fixer ──────────────────────────────────────────────────────────────

def fix_docx(input_bytes: bytes, options: dict) -> bytes:
    """
    Read a .docx (zip), patch the XML, return the fixed .docx bytes.
    """
    in_buf  = io.BytesIO(input_bytes)
    out_buf = io.BytesIO()

    with ZipFile(in_buf, 'r') as zin, ZipFile(out_buf, 'w', compression=8) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == 'word/document.xml':
                data = _patch_document(data, options)
            elif item.filename == 'word/settings.xml':
                data = _patch_settings(data)
            elif item.filename == 'word/styles.xml' and options.get('fix_fonts'):
                data = _patch_styles(data)

            zout.writestr(item, data)

    return out_buf.getvalue()


def _patch_document(xml_bytes: bytes, options: dict) -> bytes:
    parser = etree.XMLParser(recover=True, remove_blank_text=False) if LXML_AVAILABLE else None
    root   = etree.fromstring(xml_bytes, parser) if LXML_AVAILABLE else etree.fromstring(xml_bytes)

    body = root.find(f'.//{tag(W,"body")}')
    if body is None:
        return xml_bytes

    paragraphs = body.findall(f'.//{tag(W,"p")}')

    for para in paragraphs:
        pPr = para.find(tag(W,'pPr'))
        if pPr is None:
            pPr = etree.SubElement(para, tag(W,'pPr'))
            para.insert(0, pPr)

        if options.get('fix_spacing'):
            _fix_paragraph_spacing(pPr)

        if options.get('fix_margins'):
            _fix_indentation(pPr)

        if options.get('fix_fonts'):
            _fix_run_fonts(para)

    if options.get('fix_images'):
        _fix_floating_images(body)

    if options.get('fix_tables'):
        _fix_tables(body)

    if LXML_AVAILABLE:
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return etree.tostring(root, encoding='unicode').encode('utf-8')


def _fix_paragraph_spacing(pPr):
    """
    Standardise spacing: remove excessive space-before/after, cap line spacing.
    Teams / Word Online renders best with:
      - spacingAfter  ≤ 160 twips  (about 8pt)
      - spacingBefore ≤ 160 twips
      - line spacing  = auto (no exact / atLeast > 480)
    """
    W_ns = W
    spacing = pPr.find(tag(W_ns, 'spacing'))
    if spacing is None:
        spacing = etree.SubElement(pPr, tag(W_ns, 'spacing'))

    def _clamp(attr, max_val):
        val = spacing.get(tag(W_ns, attr))
        if val and val.lstrip('-').isdigit():
            if int(val) > max_val:
                spacing.set(tag(W_ns, attr), str(max_val))

    _clamp('before', 160)
    _clamp('after',  160)

    line     = spacing.get(tag(W_ns, 'line'))
    lineRule = spacing.get(tag(W_ns, 'lineRule'))

    if line and line.lstrip('-').isdigit():
        if lineRule in ('exact', 'atLeast') and int(line) > 480:
            spacing.set(tag(W_ns, 'line'),     '276')   # 1.15×
            spacing.set(tag(W_ns, 'lineRule'), 'auto')
        elif lineRule == 'auto' and int(line) > 720:    # > 3×
            spacing.set(tag(W_ns, 'line'), '480')       # 2×


def _fix_indentation(pPr):
    """Remove unusually large left/right indents that break Teams layout."""
    ind = pPr.find(tag(W, 'ind'))
    if ind is None:
        return
    for attr in ('left', 'right', 'hanging'):
        val = ind.get(tag(W, attr))
        if val and val.lstrip('-').isdigit():
            if abs(int(val)) > 2880:   # > 5 cm
                ind.set(tag(W, attr), '0')


def _fix_run_fonts(para):
    """
    Replace fonts that are not available in Word Online with safe equivalents.
    """
    UNSAFE = {
        'Calibri Light': 'Calibri',
        'Cambria Math':  'Cambria',
        'Symbol':        'Arial',
        'Wingdings':     'Arial',
        'Wingdings 2':   'Arial',
        'Wingdings 3':   'Arial',
        'Webdings':      'Arial',
    }
    for rPr in para.findall(f'.//{tag(W,"rPr")}'):
        rFonts = rPr.find(tag(W, 'rFonts'))
        if rFonts is None:
            continue
        for attr in list(rFonts.attrib):
            val = rFonts.get(attr)
            if val in UNSAFE:
                rFonts.set(attr, UNSAFE[val])


def _fix_floating_images(body):
    """
    Convert floating (anchor) drawings to inline so they don't jump around
    in Word Online / Teams.
    Strategy: wrap anchor content in an inline element.
    """
    for drawing in body.findall(f'.//{tag(W,"drawing")}'):
        anchor = drawing.find(tag(WP, 'anchor'))
        if anchor is None:
            continue

        # Build a replacement <wp:inline> from the anchor's graphic
        graphic = anchor.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
        if graphic is None:
            continue

        # Try to get size (extent)
        extent = anchor.find(tag(WP, 'extent'))
        cx = extent.get('cx', '3000000') if extent is not None else '3000000'
        cy = extent.get('cy', '2000000') if extent is not None else '2000000'

        inline = etree.Element(tag(WP, 'inline'))
        inline.set('distT', '0'); inline.set('distB', '0')
        inline.set('distL', '114240'); inline.set('distR', '114240')

        ext_el = etree.SubElement(inline, tag(WP, 'extent'))
        ext_el.set('cx', cx); ext_el.set('cy', cy)

        effExtent = etree.SubElement(inline, tag(WP, 'effectExtent'))
        for side in ('l','t','r','b'): effExtent.set(side, '0')

        docPr = anchor.find(tag(WP, 'docPr'))
        if docPr is not None:
            inline.append(copy.deepcopy(docPr))

        inline.append(copy.deepcopy(graphic))

        # Replace anchor with inline inside the drawing element
        drawing.remove(anchor)
        drawing.append(inline)


def _fix_tables(body):
    """Ensure every table cell has at least minimal padding for Teams rendering."""
    for tbl in body.findall(f'.//{tag(W,"tbl")}'):
        tblPr = tbl.find(tag(W, 'tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, tag(W, 'tblPr'))
            tbl.insert(0, tblPr)

        # Set default cell margins if not set
        tblCellMar = tblPr.find(tag(W, 'tblCellMar'))
        if tblCellMar is None:
            tblCellMar = etree.SubElement(tblPr, tag(W, 'tblCellMar'))
            for side in ('top', 'left', 'bottom', 'right'):
                el = etree.SubElement(tblCellMar, tag(W, side))
                el.set(tag(W, 'w'), '80')
                el.set(tag(W, 'type'), 'dxa')


def _patch_settings(xml_bytes: bytes) -> bytes:
    """Disable compatibility modes that break Word Online rendering."""
    try:
        root = etree.fromstring(xml_bytes)
        # Remove compat settings that force old rendering
        W_ns = W
        compat = root.find(tag(W_ns, 'compat'))
        if compat is not None:
            # Keep compat but override problematic flags
            for bad in ('useWord2002TableStyleRules', 'useWord97LineBreakRules',
                        'useWord2003FootnoteLineBreakRules'):
                el = compat.find(tag(W_ns, bad))
                if el is not None:
                    compat.remove(el)
        if LXML_AVAILABLE:
            return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        return etree.tostring(root, encoding='unicode').encode('utf-8')
    except Exception:
        return xml_bytes


def _patch_styles(xml_bytes: bytes) -> bytes:
    """Replace unsafe fonts in style definitions."""
    try:
        xml_str = xml_bytes.decode('utf-8', errors='replace')
        replacements = {
            'Calibri Light': 'Calibri',
            'Cambria Math':  'Cambria',
        }
        for old, new in replacements.items():
            xml_str = xml_str.replace(old, new)
        return xml_str.encode('utf-8')
    except Exception:
        return xml_bytes


# ── Vercel handler ──────────────────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):

    def do_POST(self):
        if self.path != '/api/fix':
            self._send(404, {'error': 'Not found'})
            return

        content_type = self.headers.get('Content-Type', '')
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)

        # Parse multipart form
        environ = {
            'REQUEST_METHOD': 'POST',
            'CONTENT_TYPE': content_type,
            'CONTENT_LENGTH': str(content_length),
        }
        fs = cgi.FieldStorage(
            fp=io.BytesIO(body),
            environ=environ,
            keep_blank_values=True
        )

        # Get file
        if 'file' not in fs:
            self._send(400, {'error': 'Aucun fichier reçu.'})
            return

        file_item = fs['file']
        if not hasattr(file_item, 'file'):
            self._send(400, {'error': 'Fichier invalide.'})
            return

        file_bytes = file_item.file.read()

        # Get options
        options = {}
        if 'options' in fs:
            try:
                options = json.loads(fs['options'].value)
            except Exception:
                options = {}

        # Default all options to True if missing
        for k in ('fix_spacing', 'fix_margins', 'fix_fonts', 'fix_images', 'fix_tables'):
            options.setdefault(k, True)

        # Process
        try:
            fixed_bytes = fix_docx(file_bytes, options)
        except Exception as e:
            self._send(500, {'error': f'Erreur lors du traitement : {str(e)}'})
            return

        # Return fixed file
        original_name = getattr(file_item, 'filename', 'document.docx') or 'document.docx'
        fixed_name = original_name.replace('.docx', '_teams.docx')

        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        self.send_header('Content-Disposition', f'attachment; filename="{fixed_name}"')
        self.send_header('Content-Length', str(len(fixed_bytes)))
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
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass  # silence logs
