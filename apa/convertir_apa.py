"""
Motor de formato APA 7 para documentos de Word (.docx)

Reglas del usuario (aplican a todo el documento):
  Times New Roman 12 | interlineado 1,5 | texto justificado | márgenes 2,5 cm

Reglas APA 7 que agrega:
  - Portada (opcional) con título en negrita, autor, institución, curso, asesor y fecha
  - Número de página arriba a la derecha
  - Títulos nivel 1 a 5 con el estilo APA (nivel 1 centrado, negrita, mayúscula tipo título)
  - Sangría de primera línea de 1,27 cm en los párrafos
  - Referencias: sangría francesa de 1,27 cm y orden alfabético
  - Tablas sin líneas verticales (solo superior, bajo la cabecera e inferior)
  - Citas en bloque (40 palabras o más) con sangría
  - Se eliminan párrafos vacíos y se convierten tablas hechas con barras '|' a tablas reales

Uso por terminal:  python convertir_apa.py [archivo.docx] [--conservar]
"""

import re
import sys
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.enum.text import WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FUENTE = "Times New Roman"
TAMANO = Pt(12)
MARGEN = Cm(2.5)
INTERLINEADO = 1.5
SANGRIA = Cm(1.27)
NEGRO = RGBColor(0, 0, 0)

SECCIONES_N1 = {
    "resumen", "abstract", "introducción", "introduccion", "conclusiones",
    "conclusión", "conclusion", "recomendaciones", "agradecimientos",
    "dedicatoria", "anexos", "apéndice", "apendice", "apéndices",
    "referencias", "bibliografía", "bibliografia", "referencias bibliográficas",
    "metodología", "metodologia", "resultados", "discusión", "discusion",
    "objetivos", "índice", "indice",
}
MENORES = {
    "a", "al", "ante", "con", "de", "del", "e", "el", "en", "la", "las", "lo",
    "los", "ni", "o", "para", "por", "que", "se", "sin", "sobre", "tras", "u",
    "un", "una", "y", "the", "of", "and", "in", "on", "for", "to", "an", "at", "by", "or",
}
RE_CAPTION = re.compile(r"^(tabla|figura)\s+\d+", re.I)
RE_ETIQUETA = re.compile(r"^(tabla|figura)\s+\d+\.?$", re.I)
RE_NOTA = re.compile(r"^nota[\.:]", re.I)
LETRA = re.compile(r"[^\W\d_]")
_CONSERVAR = False  # modo "conservar estructura" (formularios e informes con formato propio)


# ---------------------------------------------------------------- utilidades
def poner_fuente(font):
    font.name = FUENTE
    font.size = TAMANO
    rpr = font._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FUENTE)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def base_formato(p, alineacion=AL.JUSTIFY):
    pf = p.paragraph_format
    if _CONSERVAR and alineacion == AL.JUSTIFY:
        xml = p._p.xml
        if p.alignment in (AL.CENTER, AL.RIGHT):
            alineacion = p.alignment  # respeta lo centrado o a la derecha
        elif "<w:tab/>" in xml or "<w:br/>" in xml or "___" in p.text:
            alineacion = AL.LEFT  # líneas de formulario: justificar las deforma
    pf.alignment = alineacion
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = INTERLINEADO
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    for run in p.runs:
        poner_fuente(run.font)


def insertar_antes(padre, nuevo, etiquetas):
    """Inserta respetando el orden que exige el esquema de Word."""
    for et in etiquetas:
        ref = padre.find(qn(et))
        if ref is not None:
            ref.addprevious(nuevo)
            return
    padre.append(nuevo)


def borde(tag, val="single"):
    e = OxmlElement(tag)
    e.set(qn("w:val"), val)
    if val == "single":
        e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
    return e


# ------------------------------------------------------------------- títulos
def capitalizar(w):
    for i, c in enumerate(w):
        if c.isalpha():
            return w[:i] + c.upper() + w[i + 1:].lower()
    return w


def titulo_apa(texto):
    """Mayúscula tipo título: palabras cortas en minúscula salvo primera y última."""
    palabras = texto.split(" ")
    idx = [i for i, w in enumerate(palabras) if LETRA.search(w)]
    if not idx:
        return texto
    primero, ultimo = idx[0], idx[-1]
    todo_mayus = texto.isupper()
    salida, tras_dos_puntos = [], False
    for i, w in enumerate(palabras):
        if not LETRA.search(w):
            salida.append(w)
            continue
        base = re.sub(r"^\W+|\W+$", "", w)
        if len(base) > 1 and base.isupper() and not todo_mayus:
            salida.append(w)  # sigla (IA, APA...)
        elif i in (primero, ultimo) or tras_dos_puntos or base.lower() not in MENORES:
            salida.append(capitalizar(w))
        else:
            salida.append(w.lower())
        tras_dos_puntos = w.endswith(":")
    return " ".join(salida)


def nivel_titulo(p):
    nombre = (p.style.name or "").lower() if p.style is not None else ""
    m = re.match(r"(?:heading|t[ií]tulo)\s*(\d)", nombre)
    if m:
        return min(int(m.group(1)), 5)
    if nombre in ("title", "título", "titulo"):
        return 1
    t = p.text.strip().lower().rstrip(".:")
    if t in SECCIONES_N1 and len(p.text) < 40:
        return 1
    return 0


def es_referencias(texto):
    return bool(re.match(r"^\s*(\d+[\.\)]?\s*)?(referencias|bibliograf[ií]a)", texto.lower()))


def aplicar_titulo(p, nivel):
    if _CONSERVAR:
        texto = p.text.strip()
        centrar = nivel == 1 or (texto.isupper() and len(texto) < 70)
        base_formato(p, AL.CENTER if centrar else AL.LEFT)
        p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.keep_with_next = True
        for r in p.runs:
            r.font.bold = True
            r.font.italic = False
            r.font.color.rgb = NEGRO
        return
    if not p.runs:
        return
    nuevo = titulo_apa(p.text.strip())
    nuevo = nuevo.rstrip(".") if nivel <= 3 else (nuevo if nuevo.endswith(".") else nuevo + ".")
    if nuevo != p.text:
        p.runs[0].text = nuevo
        for r in p.runs[1:]:
            r.text = ""
    base_formato(p, AL.CENTER if nivel == 1 else AL.LEFT)
    pf = p.paragraph_format
    pf.left_indent = Cm(0)
    pf.first_line_indent = SANGRIA if nivel >= 4 else Cm(0)
    pf.keep_with_next = True
    for r in p.runs:
        r.font.bold = True
        r.font.italic = nivel in (3, 5)
        r.font.color.rgb = NEGRO


# --------------------------------------------------------------------- cuerpo
def quitar_vacios(doc):
    claves = ("w:drawing", "w:pict", "<w:br ", "<w:br/", "w:sectPr", "w:fldChar", "w:object")
    for p in list(doc.paragraphs):
        if p.text.strip():
            continue
        xml = p._p.xml
        if any(k in xml for k in claves):
            continue
        p._p.getparent().remove(p._p)


def quitar_vacios_finales(doc):
    """Quita los párrafos vacíos del final del documento (evitan hojas en blanco)."""
    claves = ("w:drawing", "w:pict", "<w:br ", "<w:br/", "w:sectPr", "w:fldChar", "w:object")
    body = doc.element.body
    while True:
        hijos = [c for c in body if c.tag != qn("w:sectPr")]
        if not hijos or hijos[-1].tag != qn("w:p"):
            return
        p = hijos[-1]
        if "".join(p.itertext()).strip() or any(k in p.xml for k in claves):
            return
        anterior = p.getprevious()
        if anterior is not None and anterior.tag == qn("w:tbl"):
            # Word exige un párrafo después de una tabla: se deja uno mínimo (1 pt)
            ppr = p.get_or_add_pPr()
            for e in ppr.findall(qn("w:spacing")):
                ppr.remove(e)
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "20")
            sp.set(qn("w:lineRule"), "exact")
            ppr.append(sp)
            return
        body.remove(p)


def convertir_tablas_texto(doc):
    """Convierte líneas tipo '| A | B |' en tablas reales de Word."""
    nuevas = []

    def volcar(grupo):
        if len(grupo) < 2:
            return
        filas = [[c.strip() for c in g.text.strip().strip("|").split("|")] for g in grupo]
        filas = [f for f in filas if any(f) and not all(re.fullmatch(r":?-{2,}:?", c) for c in f if c)]
        if not filas:
            return
        ncols = max(len(f) for f in filas)
        tabla = doc.add_table(rows=len(filas), cols=ncols)
        for i, f in enumerate(filas):
            for j in range(ncols):
                tabla.cell(i, j).text = f[j] if j < len(f) else ""
        grupo[0]._p.addprevious(tabla._tbl)
        nuevas.append(tabla._tbl)
        for g in grupo:
            g._p.getparent().remove(g._p)

    grupo = []
    for p in list(doc.paragraphs):
        t = p.text.strip()
        if t.startswith("|") and t.endswith("|") and t.count("|") >= 2:
            grupo.append(p)
        else:
            volcar(grupo)
            grupo = []
    volcar(grupo)
    return nuevas


def tiene_numeracion(p):
    ppr = p._p.pPr
    nombre = (p.style.name or "").lower() if p.style is not None else ""
    return (ppr is not None and ppr.numPr is not None) or nombre.startswith(("list", "lista"))


def se_indenta(p, texto, xml):
    if _CONSERVAR:
        pf = p.paragraph_format
        return (p.alignment == AL.JUSTIFY and len(texto) >= 80 and "___" not in texto
                and "<w:tab/>" not in xml and "<w:br/>" not in xml
                and not (pf.first_line_indent or 0) and not (pf.left_indent or 0))
    return len(texto) >= 40 or texto.endswith(".")


def es_cita_bloque(t):
    return len(t.split()) >= 40 and t[0] in "“\"«" and t[-1] in "”\"»"


def formatear_cuerpo(doc):
    refs, en_refs = [], False
    for p in list(doc.paragraphs):
        texto = p.text.strip()
        nivel = nivel_titulo(p)
        if nivel:
            aplicar_titulo(p, nivel)
            if nivel == 1:
                en_refs = es_referencias(texto)
            continue
        xml = p._p.xml
        con_imagen = "w:drawing" in xml or "w:pict" in xml
        if _CONSERVAR and not texto and not con_imagen:
            continue  # espacios de diseño del formulario: se dejan como están
        base_formato(p)
        if con_imagen:
            p.alignment = AL.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            continue
        if not texto:
            continue
        pf = p.paragraph_format
        if en_refs:
            p.alignment = AL.LEFT
            pf.left_indent = SANGRIA
            pf.first_line_indent = -SANGRIA
            refs.append(p)
        elif RE_CAPTION.match(texto) or RE_NOTA.match(texto):
            p.alignment = AL.LEFT
            pf.first_line_indent = Cm(0)
            if RE_ETIQUETA.match(texto):
                for r in p.runs:
                    r.font.bold = True
        elif tiene_numeracion(p):
            continue
        elif es_cita_bloque(texto):
            pf.left_indent = SANGRIA
            pf.first_line_indent = Cm(0)
        elif se_indenta(p, texto, xml):
            pf.first_line_indent = SANGRIA
    return refs


def clave_orden(texto):
    """Orden alfabético sin tener en cuenta tildes ni signos iniciales."""
    sin_tildes = unicodedata.normalize("NFD", texto)
    sin_tildes = "".join(c for c in sin_tildes if unicodedata.category(c) != "Mn")
    return re.sub(r"^\W+", "", sin_tildes).casefold()


def compactar_antes_de_tablas(doc):
    """Modo conservar: los títulos cortos pegados a una tabla van a espacio sencillo,
    para que la tabla no se desborde a una página nueva por el mayor interlineado."""
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    for tbl in [e for e in body if e.tag == qn("w:tbl")]:
        e, n = tbl.getprevious(), 0
        while e is not None and e.tag == qn("w:p") and n < 4:
            p = Paragraph(e, doc)
            texto = p.text.strip()
            if texto:
                if len(texto) > 80:
                    break
                p.paragraph_format.line_spacing = 1.0
            e, n = e.getprevious(), n + 1


def ordenar_referencias(refs):
    grupos = []
    for p in refs:
        if grupos and grupos[-1][-1]._p.getnext() is p._p:
            grupos[-1].append(p)
        else:
            grupos.append([p])
    for g in grupos:
        if len(g) < 2:
            continue
        padre = g[-1]._p.getparent()
        siguiente = g[-1]._p.getnext()
        for p in g:
            padre.remove(p._p)
        for p in sorted(g, key=lambda x: clave_orden(x.text)):
            if siguiente is not None:
                siguiente.addprevious(p._p)
            else:
                padre.append(p._p)


# -------------------------------------------------------------------- tablas
def formatear_tabla(tabla):
    tblpr = tabla._tbl.tblPr
    for e in tblpr.findall(qn("w:tblBorders")):
        tblpr.remove(e)
    b = OxmlElement("w:tblBorders")
    for tag, val in (("w:top", "single"), ("w:left", "nil"), ("w:bottom", "single"),
                     ("w:right", "nil"), ("w:insideH", "nil"), ("w:insideV", "nil")):
        b.append(borde(tag, val))
    insertar_antes(tblpr, b, ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
                              "w:tblCaption", "w:tblDescription"))
    for tc in tabla._tbl.iter(qn("w:tc")):
        tcpr = tc.tcPr
        if tcpr is not None:
            for t in ("w:tcBorders", "w:shd"):
                for e in tcpr.findall(qn(t)):
                    tcpr.remove(e)
    if tabla.rows:
        vistos = set()
        for celda in tabla.rows[0].cells:
            if celda._tc in vistos:
                continue
            vistos.add(celda._tc)
            tb = OxmlElement("w:tcBorders")
            tb.append(borde("w:bottom"))
            insertar_antes(celda._tc.get_or_add_tcPr(), tb,
                           ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                            "w:tcFitText", "w:vAlign", "w:hideMark"))
    for i, fila in enumerate(tabla.rows):
        for celda in fila.cells:
            for p in celda.paragraphs:
                base_formato(p, AL.LEFT)
                p.paragraph_format.first_line_indent = Cm(0)
                if i == 0:
                    for r in p.runs:
                        r.font.bold = True
            for sub in celda.tables:
                formatear_tabla(sub)


# ------------------------------------------------------ página, portada, estilos
def campo_pagina(p):
    def run_con(el):
        r = p.add_run()
        poner_fuente(r.font)
        r._r.append(el)

    def fld(tipo):
        e = OxmlElement("w:fldChar")
        e.set(qn("w:fldCharType"), tipo)
        return e

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    t = OxmlElement("w:t")
    t.text = "1"
    for el in (fld("begin"), instr, fld("separate"), t, fld("end")):
        run_con(el)


def numero_pagina(sec):
    h = sec.header
    h.is_linked_to_previous = False
    parrafos = h.paragraphs
    p = parrafos[0] if parrafos else h.add_paragraph()
    for extra in parrafos[1:]:
        extra._p.getparent().remove(extra._p)
    for hijo in list(p._p):
        if hijo.tag != qn("w:pPr"):
            p._p.remove(hijo)
    base_formato(p, AL.RIGHT)
    p.paragraph_format.first_line_indent = Cm(0)
    campo_pagina(p)


PORTADA_CAMPOS = (
    ("titulo", "Título del trabajo"),
    ("autor", "Autor"),
    ("institucion", "Institución"),
    ("programa", "Carrera o programa"),
    ("curso", "Curso"),
    ("asesor", "Asesor o docente"),
    ("fecha", "Fecha"),
)


def agregar_portada(doc, d):
    """Portada: cada dato con su etiqueta en negrita y, debajo, lo que escribió el estudiante."""
    ancla = doc.element.body[0]

    def linea(texto="", negrita=False):
        p = doc.add_paragraph()
        ancla.addprevious(p._p)
        base_formato(p, AL.CENTER)
        p.paragraph_format.first_line_indent = Cm(0)
        if texto:
            r = p.add_run(texto)
            poner_fuente(r.font)
            r.font.bold = negrita
        return p

    for _ in range(3):
        linea()
    ultimo = None
    llenos = [(et, (d.get(k) or "").strip(), k) for k, et in PORTADA_CAMPOS if (d.get(k) or "").strip()]
    for n, (etiqueta, valor, clave) in enumerate(llenos):
        linea(etiqueta + ":", True)
        ultimo = linea(titulo_apa(valor) if clave == "titulo" else valor)
        if n < len(llenos) - 1:
            linea()
    if ultimo is None:
        ultimo = linea()
    if not ultimo.runs:
        ultimo.add_run()
    ultimo.runs[-1].add_break(WD_BREAK.PAGE)


def formatear_estilos(doc):
    for estilo in doc.styles:
        if estilo.type not in (1, 2):  # solo párrafo y carácter
            continue
        try:
            poner_fuente(estilo.font)
            n = (estilo.name or "").lower()
            if not _CONSERVAR and n.startswith(("heading", "título", "titulo", "title")):
                estilo.font.color.rgb = NEGRO
        except Exception:
            pass


# ------------------------------------------------------------ función principal
def formatear_documento(doc, portada=None, ordenar_refs=True, conservar=False):
    global _CONSERVAR
    _CONSERVAR = conservar
    try:
        _formatear(doc, portada, ordenar_refs)
    finally:
        _CONSERVAR = False
    return doc


def _formatear(doc, portada, ordenar_refs):
    if not _CONSERVAR:
        quitar_vacios(doc)
    nuevas = convertir_tablas_texto(doc)
    if not _CONSERVAR:
        formatear_estilos(doc)

    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = MARGEN
        sec.left_margin = sec.right_margin = MARGEN
        for p in sec.footer.paragraphs:
            base_formato(p)

    refs = formatear_cuerpo(doc)
    if _CONSERVAR:
        compactar_antes_de_tablas(doc)
    for tabla in doc.tables:
        # En modo conservar las tablas del documento no se tocan; solo las nuevas (hechas con '|')
        if not _CONSERVAR or tabla._tbl in nuevas:
            formatear_tabla(tabla)
    if ordenar_refs:
        ordenar_referencias(refs)

    quitar_vacios_finales(doc)
    doc.settings.odd_and_even_pages_header_footer = False
    for i, sec in enumerate(doc.sections):
        sec.different_first_page_header_footer = False
        if i == 0 or not sec.header.is_linked_to_previous:
            numero_pagina(sec)

    if portada and (portada.get("titulo") or "").strip():
        agregar_portada(doc, portada)


def convertir(ruta_entrada, ruta_salida=None, conservar=False):
    ruta_entrada = Path(ruta_entrada)
    if ruta_salida is None:
        ruta_salida = ruta_entrada.with_name(ruta_entrada.stem + "_APA.docx")
    doc = formatear_documento(Document(str(ruta_entrada)), conservar=conservar)
    doc.save(str(ruta_salida))
    return ruta_salida


def elegir_archivo():
    import tkinter as tk
    from tkinter import filedialog

    raiz = tk.Tk()
    raiz.withdraw()
    return filedialog.askopenfilename(
        title="Elige tu documento de Word",
        filetypes=[("Documentos de Word", "*.docx")],
    )


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    conservar = "--conservar" in sys.argv
    ruta = args[0] if args else elegir_archivo()
    if not ruta:
        print("No se eligió ningún archivo.")
        return
    if not ruta.lower().endswith(".docx"):
        print("El archivo debe ser .docx (si es .doc, guárdalo como .docx en Word).")
        return
    print(f"Listo: {convertir(ruta, conservar=conservar)}")


if __name__ == "__main__":
    main()
