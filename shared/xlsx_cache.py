"""Conserva los resultados de las fórmulas al guardar un libro con openpyxl.

openpyxl guarda las fórmulas SIN su último resultado calculado (<v> vacío). Excel los recalcula al abrir el libro
en modo edición, pero en Vista protegida, en la vista previa de Windows/Outlook/OneDrive, en el celular y al leer
el archivo con Python esas celdas aparecen VACÍAS hasta que alguien edita y guarda en Excel.

Después de guardar, `restaurar_resultados` vuelve a escribir en cada fórmula el resultado que tenía en el archivo
original (misma hoja, misma celda y misma fórmula) y el valor calculado por el programa para las fórmulas nuevas.
"""
import os
import re
import zipfile
from xml.sax.saxutils import escape, unescape

_RE_CELDA_FORMULA = re.compile(
    r'<c r="([A-Z]+\d+)"([^>]*)>(<f(?:\s[^>]*)?/>|<f(?:\s[^>]*)?(?<!/)>.*?</f>)(<v>.*?</v>|<v\s*/>)?(</c>)', re.S)
_RE_SHEET = re.compile(r"<(?:\w+:)?sheet\b[^>]*>")
_RE_REL = re.compile(r"<(?:\w+:)?Relationship\b[^>]*>")


def _attrs(tag):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', tag))


def mapa_hojas(zf):
    """{nombre de hoja: ruta de la parte XML} de un .xlsx abierto."""
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rels = {a.get("Id"): a.get("Target", "") for a in map(_attrs, _RE_REL.findall(zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")))}
    hojas = {}
    for tag in _RE_SHEET.findall(wb):
        a = _attrs(tag)
        rid = next((v for k, v in a.items() if k.endswith(":id")), None)
        destino = rels.get(rid, "")
        if not destino:
            continue
        parte = destino.lstrip("/") if destino.startswith("/") else f"xl/{destino}"
        hojas[unescape(a.get("name", ""))] = os.path.normpath(parte).replace("\\", "/")
    return hojas


def _normalizar(formula):
    return re.sub(r"\s+", "", unescape(formula or "")).lstrip("=").upper()


def _resultados_originales(xml):
    """{celda: (fórmula normalizada, tipo, valor)} de las fórmulas con resultado guardado (incluye compartidas)."""
    from openpyxl.formula.translate import Translator

    maestras, salida = {}, {}
    for m in _RE_CELDA_FORMULA.finditer(xml):
        coord, attrs_c, f_xml, v_xml = m.group(1), m.group(2), m.group(3), m.group(4)
        if not v_xml or v_xml.startswith("<v/") or v_xml == "<v></v>":
            continue
        attrs_f = _attrs(re.match(r"<f([^>]*)>", f_xml).group(1))
        texto = re.sub(r"^<f[^>]*>|</f>$", "", f_xml) if not f_xml.endswith("/>") else ""
        if attrs_f.get("t") == "shared":
            si = attrs_f.get("si")
            if texto:
                maestras[si] = (unescape(texto), coord)
            elif si in maestras:
                base, origen = maestras[si]
                try:
                    texto = escape(Translator("=" + base, origin=origen).translate_formula(coord).lstrip("="))
                except Exception:
                    continue
        if not texto:
            continue
        tipo = _attrs(attrs_c).get("t", "")
        salida[coord] = (_normalizar(texto), tipo, v_xml[3:-4])
    return salida


def _valor_xml(valor):
    """(atributo t, contenido de <v>) para un valor calculado por el programa."""
    if isinstance(valor, bool):
        return "b", str(int(valor))
    if isinstance(valor, (int, float)):
        return "", repr(float(valor))
    return "str", escape(str(valor))


def _reemplazar(xml, originales, nuevos):
    def sustituir(m):
        coord, attrs_c, f_xml, _v, cierre = m.groups()
        if coord in nuevos:
            tipo, valor = _valor_xml(nuevos[coord])
        else:
            orig = originales.get(coord)
            texto = re.sub(r"^<f[^>]*>|</f>$", "", f_xml) if not f_xml.endswith("/>") else ""
            if not orig or orig[0] != _normalizar(texto):
                return m.group(0)
            _, tipo, valor = orig
        attrs = re.sub(r'\st="[^"]*"', "", attrs_c) + (f' t="{tipo}"' if tipo else "")
        return f'<c r="{coord}"{attrs}>{f_xml}<v>{valor}</v>{cierre}'
    return _RE_CELDA_FORMULA.sub(sustituir, xml)


def restaurar_resultados(ruta_original, ruta_nueva, valores_nuevos=None):
    """Reescribe ruta_nueva con los resultados de fórmula del original y los `valores_nuevos` {(hoja, celda): valor}."""
    valores_nuevos = valores_nuevos or {}
    with zipfile.ZipFile(ruta_original) as zo:
        hojas_o = mapa_hojas(zo)
        originales = {h: zo.read(p).decode("utf-8") for h, p in hojas_o.items() if p in zo.namelist()}
    originales = {h: _resultados_originales(x) for h, x in originales.items() if "<f" in x}

    with zipfile.ZipFile(ruta_nueva) as zn:
        hojas_n = mapa_hojas(zn)
        cambios = {}
        for hoja, parte in hojas_n.items():
            nuevos = {c: v for (h, c), v in valores_nuevos.items() if h == hoja}
            if not nuevos and not originales.get(hoja):
                continue
            xml = zn.read(parte).decode("utf-8")
            if "<f" not in xml:
                continue
            nuevo = _reemplazar(xml, originales.get(hoja, {}), nuevos)
            if nuevo != xml:
                cambios[parte] = nuevo.encode("utf-8")
        if not cambios:
            return 0
        tmp = f"{ruta_nueva}.cache"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zs:
            for item in zn.infolist():
                datos = cambios.get(item.filename)
                zs.writestr(item, datos if datos is not None else zn.read(item.filename))
    os.replace(tmp, ruta_nueva)
    return len(cambios)