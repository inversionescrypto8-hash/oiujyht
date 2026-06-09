# -*- coding: utf-8 -*-
"""
Genera 'Sistema_Gestion_Inventario.xlsx' usando SOLO la librería estándar.
Un .xlsx es un ZIP de XML (OOXML/SpreadsheetML); aquí se construye a mano.

Compatibilidad: se usan fórmulas estándar (VLOOKUP, SUMIF, SUMIFS, IF,
EOMONTH, COUNTIF, INDEX, MATCH, LARGE) que funcionan igual en Excel y en
Google Sheets. NO se usan ARRAYFORMULA ni QUERY (exclusivas de Google y que
no sobreviven la importación de un .xlsx).
"""
import zipfile
import datetime
from xml.sax.saxutils import escape

# ----------------------------------------------------------------------
# Parámetros de tamaño (filas con fórmulas precargadas)
# ----------------------------------------------------------------------
CAT_ROWS = 200      # productos: filas 2..201
MOV_ROWS = 300      # movimientos: filas 2..301
CLI_ROWS = 100      # clientes:   filas 2..101
CAT_LAST = CAT_ROWS + 1   # 201
MOV_LAST = MOV_ROWS + 1   # 301
CLI_LAST = CLI_ROWS + 1   # 101

CASA = "Casa"
ML = "Mercado Libre"

# Colores (ARGB)
C_PRIMARIO = "FF1F3A5F"
C_SECUND = "FF2E6DA4"
C_ACENTO = "FF16A085"
C_CLARO = "FFEEF3F8"
C_ALERTA = "FFC0392B"
C_ADVERT = "FFE67E22"
C_TEXTO = "FF2C3E50"
C_BLANCO = "FFFFFFFF"
C_BORDE = "FFCFD8E3"
C_DARK = "FF34495E"
C_INPUT = "FFFFFBE6"
C_AUTOCALC = "FFF4F8FC"
C_GRAY = "FF95A5A6"

# ----------------------------------------------------------------------
# Registros de estilos (deduplican y devuelven índices)
# ----------------------------------------------------------------------
class Reg:
    def __init__(self, start=0):
        self.items = []
        self.index = {}
        self.start = start

    def add(self, key, xml):
        if key in self.index:
            return self.index[key]
        idx = len(self.items) + self.start
        self.index[key] = idx
        self.items.append(xml)
        return idx


numfmts = Reg(start=164)   # ids personalizados >= 164
fonts = Reg()
fills = Reg()
borders = Reg()
xfs = Reg()

# numFmts
FMT_MONEY = numfmts.add('money', '<numFmt numFmtId="164" formatCode="&quot;$ &quot;#,##0"/>')
FMT_INT = numfmts.add('int', '<numFmt numFmtId="165" formatCode="#,##0"/>')
FMT_PCT = numfmts.add('pct', '<numFmt numFmtId="166" formatCode="0.0%"/>')
FMT_DATE = numfmts.add('date', '<numFmt numFmtId="167" formatCode="dd/mm/yyyy"/>')
FMT_SIGNED = numfmts.add('signed', '<numFmt numFmtId="168" formatCode="+#,##0;\\-#,##0;0"/>')

# fills: reservar 0 (none) y 1 (gray125)
fills.add('none', '<fill><patternFill patternType="none"/></fill>')
fills.add('gray125', '<fill><patternFill patternType="gray125"/></fill>')

def fill_solid(color):
    return fills.add('solid' + color,
                     '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % color)

FILL_PRIM = fill_solid(C_PRIMARIO)
FILL_SEC = fill_solid(C_SECUND)
FILL_ACC = fill_solid(C_ACENTO)
FILL_LIGHT = fill_solid(C_CLARO)
FILL_INPUT = fill_solid(C_INPUT)
FILL_AUTO = fill_solid(C_AUTOCALC)
FILL_ALERT = fill_solid(C_ALERTA)
FILL_ADVERT = fill_solid(C_ADVERT)
FILL_DARK = fill_solid(C_DARK)

# fonts
def font(key, sz=11, bold=False, color=C_TEXTO, name="Arial"):
    b = '<b/>' if bold else ''
    return fonts.add(key, '<font>%s<sz val="%d"/><color rgb="%s"/><name val="%s"/></font>' % (b, sz, color, name))

F_DEFAULT = font('default')
F_HDR = font('hdr', sz=10, bold=True, color=C_BLANCO)
F_BOLD = font('bold', bold=True)
F_TITLE = font('title', sz=20, bold=True, color=C_BLANCO)
F_SUBHDR = font('subhdr', sz=11, bold=True, color=C_BLANCO)
F_ALERT = font('alert', bold=True, color=C_ALERTA)
F_SECOND = font('second', color=C_SECUND)
F_GRAYSMALL = font('graysmall', sz=8, color=C_GRAY)
F_TOTAL = font('total', sz=13, bold=True, color=C_BLANCO)
F_KPI = font('kpi', sz=18, bold=True, color=C_PRIMARIO)
F_KPIW = font('kpiw', sz=18, bold=True, color=C_BLANCO)
F_COMPANY = font('company', sz=16, bold=True, color=C_PRIMARIO)
F_FACTITLE = font('factitle', sz=13, bold=True, color=C_BLANCO)
F_FOOTER = font('footer', sz=12, bold=True, color=C_PRIMARIO)

# borders
BORDER_NONE = borders.add('none', '<border><left/><right/><top/><bottom/><diagonal/></border>')
def thin():
    e = '<left style="thin"><color rgb="%s"/></left><right style="thin"><color rgb="%s"/></right><top style="thin"><color rgb="%s"/></top><bottom style="thin"><color rgb="%s"/></bottom><diagonal/>' % (C_BORDE, C_BORDE, C_BORDE, C_BORDE)
    return borders.add('thin', '<border>' + e + '</border>')
BORDER_THIN = thin()

# cellXfs
def xf(numFmtId=0, fontId=0, fillId=0, borderId=0, halign=None, valign=None, wrap=False):
    al = ''
    if halign or valign or wrap:
        parts = []
        if halign: parts.append('horizontal="%s"' % halign)
        if valign: parts.append('vertical="%s"' % valign)
        if wrap: parts.append('wrapText="1"')
        al = '<alignment %s/>' % ' '.join(parts)
    applyFont = ' applyFont="1"' if fontId else ''
    applyFill = ' applyFill="1"' if fillId else ''
    applyBorder = ' applyBorder="1"' if borderId else ''
    applyNum = ' applyNumberFormat="1"' if numFmtId else ''
    applyAl = ' applyAlignment="1"' if al else ''
    key = (numFmtId, fontId, fillId, borderId, halign, valign, wrap)
    xml = '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"%s%s%s%s%s>%s</xf>' % (
        numFmtId, fontId, fillId, borderId, applyNum, applyFont, applyFill, applyBorder, applyAl, al)
    return xfs.add(key, xml)

# Estilos base (índice 0 = default obligatorio)
S_DEFAULT = xf()
S_HDR = xf(fontId=F_HDR, fillId=FILL_PRIM, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_TITLE = xf(fontId=F_TITLE, fillId=FILL_PRIM, halign="center", valign="center")
S_BANNER_SEC = xf(fontId=F_SUBHDR, fillId=FILL_SEC, halign="left", valign="center")
S_BANNER_ALERT = xf(fontId=F_SUBHDR, fillId=FILL_ALERT, halign="left", valign="center")
S_BANNER_PRIM = xf(fontId=F_SUBHDR, fillId=FILL_PRIM, halign="left", valign="center")
S_SUBTITLE = xf(fontId=F_BOLD, fillId=FILL_LIGHT, halign="center", valign="center")
S_LABEL = xf(fontId=F_BOLD, fillId=FILL_LIGHT, borderId=BORDER_THIN, halign="left", valign="center")
S_LABEL_RIGHT = xf(fontId=F_BOLD, halign="right", valign="center")
S_TEXT = xf(borderId=BORDER_THIN, halign="left", valign="center")
S_TEXT_WRAP = xf(borderId=BORDER_THIN, halign="left", valign="top", wrap=True)
S_MONEY = xf(numFmtId=FMT_MONEY, borderId=BORDER_THIN, halign="right", valign="center")
S_INT = xf(numFmtId=FMT_INT, borderId=BORDER_THIN, halign="center", valign="center")
S_PCT = xf(numFmtId=FMT_PCT, borderId=BORDER_THIN, halign="center", valign="center")
S_DATE = xf(numFmtId=FMT_DATE, borderId=BORDER_THIN, halign="center", valign="center")
S_SIGNED = xf(numFmtId=FMT_SIGNED, borderId=BORDER_THIN, halign="center", valign="center")
# automáticas (fondo azul claro)
S_MONEY_A = xf(numFmtId=FMT_MONEY, fillId=FILL_AUTO, borderId=BORDER_THIN, halign="right", valign="center")
S_INT_A = xf(numFmtId=FMT_INT, fillId=FILL_AUTO, borderId=BORDER_THIN, halign="center", valign="center")
S_PCT_A = xf(numFmtId=FMT_PCT, fillId=FILL_AUTO, borderId=BORDER_THIN, halign="center", valign="center")
S_TEXT_A = xf(fillId=FILL_AUTO, borderId=BORDER_THIN, halign="left", valign="center")
# entrada (fondo amarillo)
S_INPUT = xf(fillId=FILL_INPUT, borderId=BORDER_THIN, halign="left", valign="center")
S_INPUT_INT = xf(numFmtId=FMT_INT, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="center", valign="center")
S_INPUT_MONEY = xf(numFmtId=FMT_MONEY, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="right", valign="center")
S_INPUT_DATE = xf(numFmtId=FMT_DATE, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="center", valign="center")
# KPIs
S_KPI_LAB_SEC = xf(fontId=F_SUBHDR, fillId=FILL_SEC, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_ACC = xf(fontId=F_SUBHDR, fillId=FILL_ACC, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_ADV = xf(fontId=F_SUBHDR, fillId=FILL_ADVERT, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_PRI = xf(fontId=F_SUBHDR, fillId=FILL_PRIM, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_DARK = xf(fontId=F_SUBHDR, fillId=FILL_DARK, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_ALE = xf(fontId=F_SUBHDR, fillId=FILL_ALERT, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_MONEY = xf(numFmtId=FMT_MONEY, fontId=F_KPI, borderId=BORDER_THIN, halign="center", valign="center")
S_KPI_INT = xf(numFmtId=FMT_INT, fontId=F_KPI, borderId=BORDER_THIN, halign="center", valign="center")
# factura
S_FACTITLE = xf(fontId=F_FACTITLE, fillId=FILL_PRIM, halign="center", valign="center")
S_COMPANY = xf(fontId=F_COMPANY, halign="left", valign="center")
S_TOTLAB = xf(fontId=F_TOTAL, fillId=FILL_ACC, borderId=BORDER_THIN, halign="right", valign="center")
S_TOTVAL = xf(numFmtId=FMT_MONEY, fontId=F_TOTAL, fillId=FILL_ACC, borderId=BORDER_THIN, halign="right", valign="center")
S_FOOTER = xf(fontId=F_FOOTER, halign="center", valign="center")
S_FOOTER_SM = xf(fontId=F_GRAYSMALL, halign="center", valign="center")
S_SECOND = xf(fontId=F_SECOND, halign="left", valign="center")

# ----------------------------------------------------------------------
# Utilidades de celda
# ----------------------------------------------------------------------
def col_letter(c):
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s

def ref(r, c):
    return col_letter(c) + str(r)

EPOCH = datetime.date(1899, 12, 30)
def serial(d):
    return (d - EPOCH).days


class Sheet:
    def __init__(self, name, hide_gridlines=False, freeze_row=0, freeze_col=0):
        self.name = name
        self.cells = {}          # (r,c) -> dict
        self.merges = []         # "A1:B2"
        self.cols = {}           # c -> width
        self.hidden_cols = set() # columnas ocultas (auxiliares)
        self.rows = {}           # r -> height
        self.validations = []    # (sqref, formula1)
        self.hide_gridlines = hide_gridlines
        self.freeze_row = freeze_row
        self.freeze_col = freeze_col
        self.max_r = 1
        self.max_c = 1

    def _touch(self, r, c):
        self.max_r = max(self.max_r, r)
        self.max_c = max(self.max_c, c)

    def text(self, r, c, value, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 's', 'v': value, 's': style}
        self._touch(r, c)

    def num(self, r, c, value, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 'n', 'v': value, 's': style}
        self._touch(r, c)

    def date(self, r, c, d, style=S_DATE):
        self.cells[(r, c)] = {'t': 'n', 'v': serial(d), 's': style}
        self._touch(r, c)

    def formula(self, r, c, f, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 'f', 'v': f, 's': style}
        self._touch(r, c)

    def blank(self, r, c, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 'b', 'v': None, 's': style}
        self._touch(r, c)

    def merge(self, r1, c1, r2, c2):
        self.merges.append(ref(r1, c1) + ":" + ref(r2, c2))

    def colw(self, c, w):
        self.cols[c] = w

    def hide_col(self, c, w=8):
        self.cols[c] = w
        self.hidden_cols.add(c)

    def rowh(self, r, h):
        self.rows[r] = h

    def validate_list(self, r1, c1, r2, c2, formula1):
        self.validations.append((ref(r1, c1) + ":" + ref(r2, c2), formula1))

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        out.append('<dimension ref="A1:%s"/>' % ref(self.max_r, self.max_c))
        # sheetViews
        sv = '<sheetView workbookViewId="0"'
        if self.hide_gridlines:
            sv += ' showGridLines="0"'
        sv += '>'
        if self.freeze_row or self.freeze_col:
            tl = ref(self.freeze_row + 1, self.freeze_col + 1)
            sv += ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                   % (self.freeze_col, self.freeze_row, tl))
            sv += '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>' % (tl, tl)
        sv += '</sheetView>'
        out.append('<sheetViews>%s</sheetViews>' % sv)
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        # cols
        if self.cols:
            out.append('<cols>')
            for c in sorted(self.cols):
                hid = ' hidden="1"' if c in self.hidden_cols else ''
                out.append('<col min="%d" max="%d" width="%.2f" customWidth="1"%s/>' % (c, c, self.cols[c], hid))
            out.append('</cols>')
        # sheetData
        out.append('<sheetData>')
        rows = {}
        for (r, c), cell in self.cells.items():
            rows.setdefault(r, {})[c] = cell
        for r in sorted(rows):
            ht = ''
            if r in self.rows:
                ht = ' ht="%.2f" customHeight="1"' % self.rows[r]
            out.append('<row r="%d"%s>' % (r, ht))
            for c in sorted(rows[r]):
                cell = rows[r][c]
                rr = ref(r, c)
                s = cell['s']
                t = cell['t']
                if t == 's':
                    out.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                               % (rr, s, escape(str(cell['v']))))
                elif t == 'n':
                    v = cell['v']
                    vs = repr(v) if isinstance(v, float) else str(v)
                    out.append('<c r="%s" s="%d"><v>%s</v></c>' % (rr, s, vs))
                elif t == 'f':
                    out.append('<c r="%s" s="%d"><f>%s</f></c>' % (rr, s, escape(cell['v'])))
                else:
                    out.append('<c r="%s" s="%d"/>' % (rr, s))
            out.append('</row>')
        out.append('</sheetData>')
        # mergeCells
        if self.merges:
            out.append('<mergeCells count="%d">' % len(self.merges))
            for m in self.merges:
                out.append('<mergeCell ref="%s"/>' % m)
            out.append('</mergeCells>')
        # dataValidations
        if self.validations:
            out.append('<dataValidations count="%d">' % len(self.validations))
            for sqref, f1 in self.validations:
                out.append('<dataValidation type="list" allowBlank="1" showInputMessage="1" '
                           'showErrorMessage="0" sqref="%s"><formula1>%s</formula1></dataValidation>'
                           % (sqref, escape(f1)))
            out.append('</dataValidations>')
        out.append('<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        out.append('</worksheet>')
        return ''.join(out)


# ----------------------------------------------------------------------
# Construcción de las hojas
# ----------------------------------------------------------------------
def q(name):
    """Nombre de hoja entre comillas para fórmulas."""
    return "'" + name + "'"

CAT = "Catálogo"
COMP = "Compras"
VEN = "Ventas"
TRA = "Traslados"
AJU = "Ajustes"
INV = "Inventario"
CLI = "Clientes"
FAC = "Factura"
CFG = "Config"
DASH = "Dashboard"

today = datetime.date(2026, 6, 4)

# ---------- CONFIG ----------
def build_config():
    s = Sheet(CFG, freeze_row=1)
    s.colw(1, 34); s.colw(2, 40); s.colw(4, 22); s.colw(5, 10); s.colw(6, 22); s.colw(8, 16); s.colw(10, 26)
    s.text(1, 1, "Parámetro", S_HDR); s.text(1, 2, "Valor", S_HDR)
    params = [
        ("Nombre de la empresa", "Mi Negocio S.A.S."),
        ("NIT / Identificación", "000.000.000-0"),
        ("Dirección", "Calle 00 # 00-00, Ciudad"),
        ("Teléfono", "+57 300 000 0000"),
        ("Correo electrónico", "correo@minegocio.com"),
        ("Sitio web", "www.minegocio.com"),
        ("Prefijo de factura", "FAC-"),
        ("Próximo número de factura", 1),
        ("Mensaje pie de factura", "¡Gracias por su compra!"),
    ]
    r = 2
    for k, v in params:
        s.text(r, 1, k, S_LABEL)
        if isinstance(v, (int, float)):
            s.num(r, 2, v, S_TEXT)
        else:
            s.text(r, 2, v, S_TEXT)
        r += 1
    # Tabla IVA por categoría (columna D = categoría, columna E = IVA %).
    # Se deja el IVA EN BLANCO: escribe 19, 5, 0, etc. según cada categoría.
    categorias = ["General", "Electrónica", "Hogar", "Ropa y calzado", "Accesorios", "Papelería", "Otros"]
    s.text(1, 4, "Categorías", S_SUBTITLE)
    s.text(1, 5, "IVA %", S_SUBTITLE)
    for i, v in enumerate(categorias):
        s.text(2 + i, 4, v, S_TEXT)
        s.blank(2 + i, 5, S_INPUT_INT)   # IVA en blanco, editable por categoría
    # listas auxiliares
    def lista(col, titulo, valores):
        s.text(1, col, titulo, S_SUBTITLE)
        for i, v in enumerate(valores):
            s.text(2 + i, col, v, S_TEXT)
    lista(6, "Ubicaciones", [CASA, ML])
    lista(8, "Estados", ["Activo", "Inactivo"])
    lista(10, "Motivos de ajuste", ["Producto dañado", "Pérdida", "Corrección de conteo", "Ajuste administrativo", "Devolución"])
    return s

# ---------- CATÁLOGO ----------
def build_catalogo():
    s = Sheet(CAT, freeze_row=1, freeze_col=2)
    headers = ["Código", "Nombre", "Categoría", "Estado", "Stock Mínimo",
               "Costo Promedio", "Stock Casa", "Stock Mercado Libre", "Stock Total",
               "Valor Inventario", "Estado Stock", "IVA %"]
    widths = [12, 32, 18, 11, 12, 15, 12, 18, 12, 16, 14, 8]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [
        ("P001", "Audífonos Bluetooth", "Electrónica", "Activo", 5),
        ("P002", "Cargador USB-C 20W", "Electrónica", "Activo", 8),
        ("P003", "Camiseta básica", "Ropa y calzado", "Activo", 10),
        ("P004", "Termo 1L acero", "Hogar", "Activo", 4),
        ("P005", "Cuaderno argollado", "Papelería", "Activo", 15),
        ("P006", "Mouse inalámbrico", "Electrónica", "Activo", 6),
    ]
    for r in range(2, CAT_LAST + 1):
        i = r - 2
        if i < len(sample):
            cod, nom, cat, est, mn = sample[i]
            s.text(r, 1, cod, S_INPUT); s.text(r, 2, nom, S_INPUT); s.text(r, 3, cat, S_INPUT)
            s.text(r, 4, est, S_INPUT); s.num(r, 5, mn, S_INPUT_INT)
        else:
            s.blank(r, 1, S_INPUT); s.blank(r, 2, S_INPUT); s.blank(r, 3, S_INPUT)
            s.blank(r, 4, S_INPUT); s.blank(r, 5, S_INPUT_INT)
        A = "$A%d" % r
        # F costo promedio
        s.formula(r, 6, 'IF(%s="","",IFERROR(SUMIF(%s!$B:$B,%s,%s!$F:$F)/SUMIF(%s!$B:$B,%s,%s!$D:$D),0))'
                  % (A, q(COMP), A, q(COMP), q(COMP), A, q(COMP)), S_MONEY_A)
        # G stock casa
        s.formula(r, 7,
                  'IF(%s="","",SUMIF(%s!$B:$B,%s,%s!$D:$D)'
                  '+SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$G:$G,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$M:$M,%s&"|%s",%s!$F:$F)'
                  '+SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D))'
                  % (A, q(COMP), A, q(COMP),
                     q(TRA), A, CASA, q(TRA),
                     q(TRA), A, CASA, q(TRA),
                     q(VEN), A, CASA, q(VEN),
                     q(AJU), A, CASA, q(AJU)), S_INT_A)
        # H stock ML
        s.formula(r, 8,
                  'IF(%s="","",SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$G:$G,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$M:$M,%s&"|%s",%s!$F:$F)'
                  '+SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D))'
                  % (A, q(TRA), A, ML, q(TRA),
                     q(TRA), A, ML, q(TRA),
                     q(VEN), A, ML, q(VEN),
                     q(AJU), A, ML, q(AJU)), S_INT_A)
        # I total
        s.formula(r, 9, 'IF(%s="","",$G%d+$H%d)' % (A, r, r), S_INT_A)
        # J valor inventario
        s.formula(r, 10, 'IF(%s="","",$I%d*$F%d)' % (A, r, r), S_MONEY_A)
        # K estado stock
        s.formula(r, 11, 'IF(%s="","",IF($I%d<=0,"Agotado",IF($I%d<=$E%d,"Stock bajo","Disponible")))'
                  % (A, r, r, r), S_TEXT_A)
        # L IVA % (según categoría, tomado de la tabla Config!D:E). En blanco si la categoría no tiene IVA.
        s.formula(r, 12, 'IF(%s="","",IFERROR(VLOOKUP($C%d,%s!$D$2:$E$8,2,FALSE),0))'
                  % (A, r, q(CFG)), S_INT_A)
    # validaciones
    s.validate_list(2, 3, CAT_LAST, 3, "%s!$D$2:$D$8" % q(CFG))
    s.validate_list(2, 4, CAT_LAST, 4, "%s!$H$2:$H$3" % q(CFG))
    return s

# ---------- COMPRAS ----------
def build_compras():
    s = Sheet(COMP, freeze_row=1)
    headers = ["Fecha", "Código", "Producto", "Cantidad", "Costo Unitario", "Costo Total"]
    widths = [13, 13, 30, 11, 15, 15]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [
        (datetime.date(2026, 6, 1), "P001", 20, 35000),
        (datetime.date(2026, 6, 1), "P002", 30, 12000),
        (datetime.date(2026, 6, 2), "P003", 40, 9000),
        (datetime.date(2026, 6, 2), "P004", 15, 22000),
        (datetime.date(2026, 6, 3), "P006", 18, 28000),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, cod, qy, cu = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, cod, S_INPUT)
            s.num(r, 4, qy, S_INPUT_INT); s.num(r, 5, cu, S_INPUT_MONEY)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT)
            s.blank(r, 4, S_INPUT_INT); s.blank(r, 5, S_INPUT_MONEY)
        B = "$B%d" % r
        s.formula(r, 3, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (B, B, q(CAT)), S_TEXT_A)
        s.formula(r, 6, 'IF(%s="","",$D%d*$E%d)' % (B, r, r), S_MONEY_A)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    return s

# ---------- VENTAS ----------
def build_ventas():
    s = Sheet(VEN, freeze_row=1)
    headers = ["Fecha", "N° Factura", "Cliente", "Código", "Producto", "Cantidad",
               "Ubicación", "Valor Recibido", "Costo Prom. Unit.", "Costo Total",
               "Utilidad", "Margen %", "clave"]
    widths = [12, 12, 22, 12, 28, 10, 16, 15, 15, 14, 14, 10, 14]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [
        (datetime.date(2026, 6, 3), "FAC-0001", "Juan Pérez", "P001", 2, CASA, 110000),
        (datetime.date(2026, 6, 3), "FAC-0001", "Juan Pérez", "P002", 1, CASA, 22000),
        (datetime.date(2026, 6, 4), "FAC-0002", "María Gómez", "P003", 3, ML, 45000),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, fac, cli, cod, qy, ub, val = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, fac, S_INPUT); s.text(r, 3, cli, S_INPUT)
            s.text(r, 4, cod, S_INPUT); s.num(r, 6, qy, S_INPUT_INT); s.text(r, 7, ub, S_INPUT)
            s.num(r, 8, val, S_INPUT_MONEY)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 3, S_INPUT)
            s.blank(r, 4, S_INPUT); s.blank(r, 6, S_INPUT_INT); s.blank(r, 7, S_INPUT)
            s.blank(r, 8, S_INPUT_MONEY)
        D = "$D%d" % r
        s.formula(r, 5, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (D, D, q(CAT)), S_TEXT_A)
        s.formula(r, 9, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$F,6,FALSE),0))' % (D, D, q(CAT)), S_MONEY_A)
        s.formula(r, 10, 'IF(%s="","",$F%d*$I%d)' % (D, r, r), S_MONEY_A)
        s.formula(r, 11, 'IF(%s="","",$H%d-$J%d)' % (D, r, r), S_MONEY_A)
        s.formula(r, 12, 'IF(%s="","",IF($H%d=0,0,($H%d-$J%d)/$H%d))' % (D, r, r, r, r), S_PCT_A)
        s.formula(r, 13, 'IF(%s="","",%s&"|"&$G%d)' % (D, D, r), S_TEXT_A)
    s.validate_list(2, 3, MOV_LAST, 3, "%s!$B$2:$B$%d" % (q(CLI), CLI_LAST))
    s.validate_list(2, 4, MOV_LAST, 4, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 7, MOV_LAST, 7, "%s!$F$2:$F$3" % q(CFG))
    return s

# ---------- TRASLADOS ----------
def build_traslados():
    s = Sheet(TRA, freeze_row=1)
    headers = ["Fecha", "Código", "Producto", "Cantidad", "Origen", "Destino", "claveOrigen", "claveDestino"]
    widths = [13, 13, 30, 11, 16, 16, 16, 16]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [
        (datetime.date(2026, 6, 3), "P001", 5, CASA, ML),
        (datetime.date(2026, 6, 3), "P003", 10, CASA, ML),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, cod, qy, ori, des = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, cod, S_INPUT); s.num(r, 4, qy, S_INPUT_INT)
            s.text(r, 5, ori, S_INPUT); s.text(r, 6, des, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 4, S_INPUT_INT)
            s.blank(r, 5, S_INPUT); s.blank(r, 6, S_INPUT)
        B = "$B%d" % r
        s.formula(r, 3, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (B, B, q(CAT)), S_TEXT_A)
        s.formula(r, 7, 'IF(%s="","",%s&"|"&$E%d)' % (B, B, r), S_TEXT_A)
        s.formula(r, 8, 'IF(%s="","",%s&"|"&$F%d)' % (B, B, r), S_TEXT_A)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 5, MOV_LAST, 5, "%s!$F$2:$F$3" % q(CFG))
    s.validate_list(2, 6, MOV_LAST, 6, "%s!$F$2:$F$3" % q(CFG))
    return s

# ---------- AJUSTES ----------
def build_ajustes():
    s = Sheet(AJU, freeze_row=1)
    headers = ["Fecha", "Código", "Producto", "Cantidad (+/-)", "Ubicación", "Motivo", "Nota", "clave"]
    widths = [13, 13, 30, 14, 16, 22, 30, 14]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [
        (datetime.date(2026, 6, 3), "P004", -1, CASA, "Producto dañado", "Llegó abollado"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, cod, qy, ub, mot, nota = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, cod, S_INPUT); s.num(r, 4, qy, S_SIGNED)
            s.text(r, 5, ub, S_INPUT); s.text(r, 6, mot, S_INPUT); s.text(r, 7, nota, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 4, S_SIGNED)
            s.blank(r, 5, S_INPUT); s.blank(r, 6, S_INPUT); s.blank(r, 7, S_INPUT)
        B = "$B%d" % r
        s.formula(r, 3, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (B, B, q(CAT)), S_TEXT_A)
        s.formula(r, 8, 'IF(%s="","",%s&"|"&$E%d)' % (B, B, r), S_TEXT_A)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 5, MOV_LAST, 5, "%s!$F$2:$F$3" % q(CFG))
    s.validate_list(2, 6, MOV_LAST, 6, "%s!$J$2:$J$6" % q(CFG))
    return s

# ---------- CLIENTES ----------
def build_clientes():
    s = Sheet(CLI, freeze_row=1)
    headers = ["ID", "Nombre", "Documento / NIT", "Teléfono", "Correo", "Dirección", "Ciudad", "Notas"]
    widths = [10, 24, 18, 16, 26, 28, 14, 26]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [
        ("Juan Pérez", "1.111.111", "573001112233", "juan@correo.com", "Cra 1 #2-3", "Bogotá", ""),
        ("María Gómez", "2.222.222", "573004445566", "maria@correo.com", "Cll 4 #5-6", "Medellín", ""),
    ]
    for r in range(2, CLI_LAST + 1):
        i = r - 2
        s.formula(r, 1, 'IF($B%d="","","CLI-"&TEXT(ROW()-1,"000"))' % r, S_TEXT_A)
        if i < len(sample):
            nom, doc, tel, cor, dirx, ciu, notas = sample[i]
            s.text(r, 2, nom, S_INPUT); s.text(r, 3, doc, S_INPUT); s.text(r, 4, tel, S_INPUT)
            s.text(r, 5, cor, S_INPUT); s.text(r, 6, dirx, S_INPUT); s.text(r, 7, ciu, S_INPUT)
            s.blank(r, 8, S_INPUT)
        else:
            for c in range(2, 9):
                s.blank(r, c, S_INPUT)
    return s

# ---------- INVENTARIO (reporte) ----------
def build_inventario():
    s = Sheet(INV, freeze_row=1)
    headers = ["Código", "Producto", "Categoría", "Costo Promedio", "Stock Casa",
               "Stock Mercado Libre", "Stock Total", "Valor Inventario", "Estado"]
    widths = [12, 30, 18, 15, 12, 18, 12, 16, 14]
    srccols = [1, 2, 3, 6, 7, 8, 9, 10, 11]
    styles = [S_TEXT_A, S_TEXT_A, S_TEXT_A, S_MONEY_A, S_INT_A, S_INT_A, S_INT_A, S_MONEY_A, S_TEXT_A]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    for r in range(2, CAT_LAST + 1):
        for j, sc in enumerate(srccols):
            col = col_letter(sc)
            s.formula(r, j + 1, 'IF(%s!$A%d="","",%s!$%s%d)' % (q(CAT), r, q(CAT), col, r), styles[j])
    return s

# ---------- DASHBOARD ----------
def build_dashboard():
    s = Sheet(DASH, hide_gridlines=True, freeze_row=2)
    for c in range(1, 13):
        s.colw(c, 13)
    # Título
    s.text(1, 1, "DASHBOARD EJECUTIVO", S_TITLE); s.merge(1, 1, 1, 12); s.rowh(1, 40)
    s.formula(2, 1, '"Mes en curso: "&TEXT(TODAY(),"mmmm yyyy")&"   ·   Actualizado: "&TEXT(NOW(),"dd/mm/yyyy hh:mm")', S_SUBTITLE)
    s.merge(2, 1, 2, 12); s.rowh(2, 22)

    ini = 'EOMONTH(TODAY(),-1)+1'
    fin = 'EOMONTH(TODAY(),0)'
    df = 'SUMIFS(%s!$H:$H,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), ini, q(VEN), fin)
    gf = 'SUMIFS(%s!$K:$K,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), ini, q(VEN), fin)
    cf = 'SUMIFS(%s!$F:$F,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(COMP), q(COMP), ini, q(COMP), fin)

    def card(row, col, label, formula, lab_style, val_style):
        s.text(row, col, label, lab_style); s.merge(row, col, row, col + 2)
        s.formula(row + 1, col, formula, val_style); s.merge(row + 1, col, row + 1, col + 2)
        s.rowh(row, 22); s.rowh(row + 1, 38)

    card(4, 1, "VENTAS DEL MES", df, S_KPI_LAB_SEC, S_KPI_MONEY)
    card(4, 4, "GANANCIA DEL MES", gf, S_KPI_LAB_ACC, S_KPI_MONEY)
    card(4, 7, "COMPRAS DEL MES", cf, S_KPI_LAB_ADV, S_KPI_MONEY)
    card(4, 10, "VALOR DEL INVENTARIO", 'SUM(%s!$J$2:$J$%d)' % (q(CAT), CAT_LAST), S_KPI_LAB_PRI, S_KPI_MONEY)
    card(7, 1, "INVENTARIO EN CASA (UNID)", 'SUM(%s!$G$2:$G$%d)' % (q(CAT), CAT_LAST), S_KPI_LAB_DARK, S_KPI_INT)
    card(7, 4, "INVENTARIO EN ML (UNID)", 'SUM(%s!$H$2:$H$%d)' % (q(CAT), CAT_LAST), S_KPI_LAB_DARK, S_KPI_INT)
    card(7, 7, "PRODUCTOS CON STOCK BAJO", 'COUNTIF(%s!$K$2:$K$%d,"Stock bajo")' % (q(CAT), CAT_LAST), S_KPI_LAB_ADV, S_KPI_INT)
    card(7, 10, "PRODUCTOS AGOTADOS", 'COUNTIF(%s!$K$2:$K$%d,"Agotado")' % (q(CAT), CAT_LAST), S_KPI_LAB_ALE, S_KPI_INT)

    # Tabla productos más vendidos (top 8) usando columnas auxiliares ocultas (P..R)
    # P: nombre producto, Q: unidades vendidas, R: unidades ajustadas (desempate)
    auxcol = 16  # P
    for r in range(2, CAT_LAST + 1):
        s.formula(r, auxcol, 'IF(%s!$A%d="","",%s!$B%d)' % (q(CAT), r, q(CAT), r))
        s.formula(r, auxcol + 1, 'IF(%s!$A%d="","",SUMIF(%s!$D:$D,%s!$A%d,%s!$F:$F))'
                  % (q(CAT), r, q(VEN), q(CAT), r, q(VEN)))
        s.formula(r, auxcol + 2, 'IF(%s!$A%d="","",$Q%d+ROW()/100000)' % (q(CAT), r, r))
    s.colw(auxcol, 8); s.colw(auxcol + 1, 8); s.colw(auxcol + 2, 8)

    s.text(11, 1, "🏆 PRODUCTOS MÁS VENDIDOS", S_BANNER_PRIM); s.merge(11, 1, 11, 5)
    s.text(12, 1, "Producto", S_HDR); s.merge(12, 1, 12, 4)
    s.text(12, 5, "Unid.", S_HDR)
    rng_adj = "$R$2:$R$%d" % CAT_LAST
    rng_nom = "$P$2:$P$%d" % CAT_LAST
    for k in range(1, 9):
        rr = 12 + k
        large = 'LARGE(%s,%d)' % (rng_adj, k)
        s.formula(rr, 1, 'IFERROR(IF(INT(%s)=0,"",INDEX(%s,MATCH(%s,%s,0))),"")'
                  % (large, rng_nom, large, rng_adj), S_TEXT); s.merge(rr, 1, rr, 4)
        s.formula(rr, 5, 'IFERROR(IF(INT(%s)=0,"",INT(%s)),"")' % (large, large), S_INT)

    # Tabla alertas de inventario
    s.text(11, 7, "⚠️ ALERTAS DE INVENTARIO (bajo / agotado)", S_BANNER_ALERT); s.merge(11, 7, 11, 12)
    s.text(12, 7, "Producto", S_HDR); s.merge(12, 7, 12, 9)
    s.text(12, 10, "Casa", S_HDR); s.text(12, 11, "ML", S_HDR); s.text(12, 12, "Total", S_HDR)
    # listamos los primeros productos con estado de alerta (hasta 12)
    for k in range(1, 13):
        rr = 12 + k
        cat_row = k + 1  # mapea a filas 2.. del catálogo (vista simple de primeros productos)
        cond = '%s!$K%d="Agotado",%s!$K%d="Stock bajo"' % (q(CAT), cat_row, q(CAT), cat_row)
        s.formula(rr, 7, 'IF(OR(%s),%s!$B%d,"")' % (cond, q(CAT), cat_row), S_TEXT); s.merge(rr, 7, rr, 9)
        s.formula(rr, 10, 'IF(OR(%s),%s!$G%d,"")' % (cond, q(CAT), cat_row), S_INT)
        s.formula(rr, 11, 'IF(OR(%s),%s!$H%d,"")' % (cond, q(CAT), cat_row), S_INT)
        s.formula(rr, 12, 'IF(OR(%s),%s!$I%d,"")' % (cond, q(CAT), cat_row), S_INT)

    # Tabla de tendencias (6 meses) - sirve para que el usuario inserte un gráfico
    s.text(26, 1, "📈 TENDENCIAS (últimos 6 meses) — selecciona A27:D33 e inserta un gráfico", S_BANNER_SEC)
    s.merge(26, 1, 26, 12)
    s.text(27, 1, "Mes", S_HDR); s.text(27, 2, "Ventas", S_HDR); s.text(27, 3, "Utilidad", S_HDR); s.text(27, 4, "Compras", S_HDR)
    for i in range(6):
        rr = 28 + i
        off = -5 + i
        mini = 'DATE(YEAR(EOMONTH(TODAY(),%d)),MONTH(EOMONTH(TODAY(),%d)),1)' % (off, off)
        mfin = 'EOMONTH(TODAY(),%d)' % off
        s.formula(rr, 1, 'TEXT(%s,"mmm yy")' % mini, S_TEXT)
        s.formula(rr, 2, 'SUMIFS(%s!$H:$H,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), mini, q(VEN), mfin), S_MONEY)
        s.formula(rr, 3, 'SUMIFS(%s!$K:$K,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), mini, q(VEN), mfin), S_MONEY)
        s.formula(rr, 4, 'SUMIFS(%s!$F:$F,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(COMP), q(COMP), mini, q(COMP), mfin), S_MONEY)
    return s

# ---------- FACTURA ----------
def build_factura():
    s = Sheet(FAC, hide_gridlines=True)
    widths = [16, 30, 12, 13, 16, 17]
    for i, w in enumerate(widths):
        s.colw(i + 1, w)
    # Encabezado empresa
    s.formula(1, 2, "%s!$B$2" % q(CFG), S_COMPANY); s.merge(1, 2, 1, 3)
    s.formula(2, 2, '"NIT: "&%s!$B$3' % q(CFG), S_TEXT); s.merge(2, 2, 2, 3)
    s.formula(3, 2, "%s!$B$4" % q(CFG), S_TEXT); s.merge(3, 2, 3, 3)
    s.formula(4, 2, '"Tel: "&%s!$B$5&"   |   "&%s!$B$6' % (q(CFG), q(CFG)), S_TEXT); s.merge(4, 2, 4, 3)
    s.formula(5, 2, "%s!$B$7" % q(CFG), S_SECOND); s.merge(5, 2, 5, 3)
    s.text(1, 1, "LOGO", S_FOOTER_SM); s.merge(1, 1, 5, 1)
    # Metadatos
    s.text(1, 4, "FACTURA DE VENTA", S_FACTITLE); s.merge(1, 4, 1, 6)
    s.text(2, 4, "N°:", S_LABEL_RIGHT)
    s.formula(2, 5, '%s!$B$8&TEXT(%s!$B$9,"0000")' % (q(CFG), q(CFG)), xf(fontId=F_ALERT, halign="left", valign="center")); s.merge(2, 5, 2, 6)
    s.text(3, 4, "Fecha:", S_LABEL_RIGHT)
    s.formula(3, 5, "TODAY()", S_DATE); s.merge(3, 5, 3, 6)
    s.text(4, 4, "Pago:", S_LABEL_RIGHT); s.text(4, 5, "Efectivo", S_INPUT); s.merge(4, 5, 4, 6)
    s.text(5, 4, "Origen inv.:", S_LABEL_RIGHT); s.text(5, 5, CASA, S_INPUT); s.merge(5, 5, 5, 6)
    s.rowh(6, 8)
    # Datos cliente
    s.text(7, 1, "DATOS DEL CLIENTE", S_BANNER_SEC); s.merge(7, 1, 7, 6)
    s.text(8, 1, "Cliente:", S_LABEL_RIGHT); s.text(8, 2, "Juan Pérez", S_INPUT); s.merge(8, 2, 8, 3)
    s.text(8, 4, "Documento:", S_LABEL_RIGHT)
    s.formula(8, 5, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,2,FALSE),"")' % q(CLI), S_TEXT); s.merge(8, 5, 8, 6)
    s.text(9, 1, "Teléfono:", S_LABEL_RIGHT)
    s.formula(9, 2, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,3,FALSE),"")' % q(CLI), S_TEXT); s.merge(9, 2, 9, 3)
    s.text(9, 4, "Correo:", S_LABEL_RIGHT)
    s.formula(9, 5, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,4,FALSE),"")' % q(CLI), S_TEXT); s.merge(9, 5, 9, 6)
    s.text(10, 1, "Dirección:", S_LABEL_RIGHT)
    s.formula(10, 2, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,5,FALSE),"")' % q(CLI), S_TEXT); s.merge(10, 2, 10, 6)
    s.rowh(11, 8)
    # Tabla productos
    s.text(12, 1, "Código", S_HDR); s.text(12, 2, "Descripción", S_HDR); s.merge(12, 2, 12, 3)
    s.text(12, 4, "Cant.", S_HDR); s.text(12, 5, "Vr. Unitario", S_HDR); s.text(12, 6, "Total", S_HDR)
    # Columnas auxiliares ocultas: G=IVA% del producto (según su categoría), H=IVA $ del renglón.
    s.hide_col(7, 8); s.hide_col(8, 8)
    sample_items = [("P001", 2, 55000), ("P002", 1, 22000)]
    for idx, r in enumerate(range(13, 25)):
        if idx < len(sample_items):
            cod, qy, vu = sample_items[idx]
            s.text(r, 1, cod, S_INPUT)
            s.num(r, 4, qy, S_INPUT_INT); s.num(r, 5, vu, S_INPUT_MONEY)
        else:
            s.blank(r, 1, S_INPUT); s.blank(r, 4, S_INPUT_INT); s.blank(r, 5, S_INPUT_MONEY)
        s.formula(r, 2, 'IF($A%d="","",IFERROR(VLOOKUP($A%d,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (r, r, q(CAT)), S_TEXT_A); s.merge(r, 2, r, 3)
        s.formula(r, 6, 'IF(OR($A%d="",$D%d="",$E%d=""),"",$D%d*$E%d)' % (r, r, r, r, r), S_MONEY_A)
        # G: IVA % del producto (columna L del Catálogo, según categoría). H: IVA $ del renglón.
        s.formula(r, 7, 'IF($A%d="","",IFERROR(VLOOKUP($A%d,%s!$A:$L,12,FALSE),0))' % (r, r, q(CAT)), S_INT_A)
        s.formula(r, 8, 'IF($F%d="","",ROUND($F%d*$G%d/100,0))' % (r, r, r), S_MONEY_A)
    # Observaciones + totales
    s.text(25, 1, "Observaciones:", S_LABEL_RIGHT)
    s.blank(26, 1, S_TEXT_WRAP); s.merge(26, 1, 29, 3)
    s.text(26, 4, "Subtotal:", S_LABEL_RIGHT); s.merge(26, 4, 26, 5)
    s.formula(26, 6, "SUM($F$13:$F$24)", S_MONEY)
    s.text(27, 4, "IVA:", S_LABEL_RIGHT); s.merge(27, 4, 27, 5)
    s.formula(27, 6, "SUM($H$13:$H$24)", S_MONEY)
    s.text(28, 4, "TOTAL A PAGAR:", S_TOTLAB); s.merge(28, 4, 28, 5)
    s.formula(28, 6, "$F$26+$F$27", S_TOTVAL); s.rowh(28, 26)
    # Pie
    s.formula(31, 1, "%s!$B$10" % q(CFG), S_FOOTER); s.merge(31, 1, 31, 6)
    s.text(32, 1, "Documento generado electrónicamente · Sistema de Gestión", S_FOOTER_SM); s.merge(32, 1, 32, 6)
    # validaciones
    s.validate_list(8, 2, 8, 2, "%s!$B$2:$B$%d" % (q(CLI), CLI_LAST))
    s.validate_list(13, 1, 24, 1, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(5, 5, 5, 5, "%s!$F$2:$F$3" % q(CFG))
    s.validate_list(4, 5, 4, 5, '"Efectivo,Transferencia,Tarjeta,Nequi / Daviplata,Crédito"')
    return s


# ----------------------------------------------------------------------
# Ensamblar el .xlsx
# ----------------------------------------------------------------------
def build_workbook(path):
    sheets = [
        build_dashboard(), build_catalogo(), build_compras(), build_ventas(),
        build_traslados(), build_ajustes(), build_inventario(), build_clientes(),
        build_factura(), build_config(),
    ]

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + ('<numFmts count="%d">' % len(numfmts.items)) + ''.join(numfmts.items) + '</numFmts>'
        + ('<fonts count="%d">' % len(fonts.items)) + ''.join(fonts.items) + '</fonts>'
        + ('<fills count="%d">' % len(fills.items)) + ''.join(fills.items) + '</fills>'
        + ('<borders count="%d">' % len(borders.items)) + ''.join(borders.items) + '</borders>'
        + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        + ('<cellXfs count="%d">' % len(xfs.items)) + ''.join(xfs.items) + '</cellXfs>'
        + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )

    # workbook.xml
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    for i, sh in enumerate(sheets):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(sh.name), i + 1, i + 1))
    wb.append('</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    workbook_xml = ''.join(wb)

    # workbook rels
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(len(sheets)):
        rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1))
    style_rid = len(sheets) + 1
    rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % style_rid)
    rels.append('</Relationships>')
    workbook_rels = ''.join(rels)

    # content types
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(len(sheets)):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
    ct.append('</Types>')
    content_types = ''.join(ct)

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', workbook_xml)
        z.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        z.writestr('xl/styles.xml', styles_xml)
        for i, sh in enumerate(sheets):
            z.writestr('xl/worksheets/sheet%d.xml' % (i + 1), sh.xml())
    return path


if __name__ == "__main__":
    out = build_workbook("Sistema_Gestion_Inventario.xlsx")
    print("Generado:", out)
