"""Reemplazo/inserción de una hoja en un .xlsx/.xlsm editando el paquete ZIP directamente.

Evita cargar y re-serializar todo el libro con openpyxl (muy lento en libros grandes) y conserva intactas
las demás hojas, macros, gráficos y tablas dinámicas.
"""
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, Iterable, Optional
from xml.etree import ElementTree
from xml.sax.saxutils import escape

TIPO_HOJA = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
REL_HOJA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_ORDEN_ESTILOS = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
                  "tableStyles", "colors", "extLst")
_HIJO = {"numFmts": "numFmt", "fonts": "font", "fills": "fill", "borders": "border", "cellXfs": "xf", "dxfs": "dxf"}


def texto_xml(valor) -> str:
    return escape(_CONTROL.sub("", str(valor)), {'"': "&quot;"})


def _attrs(tag: str) -> Dict[str, str]:
    return dict(re.findall(r'([\w:]+)="([^"]*)"', tag))


class LibroXlsx:
    """Uso: with LibroXlsx(ruta) as libro: ids = libro.estilo(...); libro.escribir_hoja(nombre, partes_xml)."""

    def __init__(self, ruta):
        self.ruta = Path(ruta)
        self._zip = zipfile.ZipFile(self.ruta)
        self._nombres = set(self._zip.namelist())
        self._mod: Dict[str, bytes] = {}
        self._omitir = set()
        self._estilos = self._leer("xl/styles.xml")
        self._p = re.search(r"<(\w+:)?styleSheet\b", self._estilos).group(1) or ""

    def __enter__(self):
        return self

    def __exit__(self, tipo, *_):
        try:
            if tipo is None:
                self._guardar()
        finally:
            self._zip.close()

    def _leer(self, nombre) -> str:
        return self._mod[nombre].decode("utf-8") if nombre in self._mod else self._zip.read(nombre).decode("utf-8")

    # ---------------------------------------------------------------- estilos
    def _agregar(self, bloque: str, fragmento: str) -> int:
        """Añade un hijo al bloque de styles.xml (lo crea si no existe) y devuelve su índice."""
        p, xml, hijo = self._p, self._estilos, _HIJO[bloque]
        m = re.search(rf"<{p}{bloque}\b[^>]*?(/?)>", xml)
        if m is None:
            siguiente = next((re.search(rf"<{p}{b}\b", xml) for b in _ORDEN_ESTILOS[_ORDEN_ESTILOS.index(bloque) + 1:]
                              if re.search(rf"<{p}{b}\b", xml)), None)
            pos = siguiente.start() if siguiente else xml.rindex(f"</{p}styleSheet>")
            xml = f'{xml[:pos]}<{p}{bloque} count="0"></{p}{bloque}>{xml[pos:]}'
        elif m.group(1):
            xml = f'{xml[:m.start()]}<{p}{bloque} count="0"></{p}{bloque}>{xml[m.end():]}'
        m = re.search(rf"<{p}{bloque}\b[^>]*>(.*?)</{p}{bloque}>", xml, re.S)
        indice = len(re.findall(rf"<{p}{hijo}[\s/>]", m.group(1)))
        apertura = re.sub(r'\scount="\d*"', "", xml[m.start():m.start(1)])
        apertura = apertura[:-1] + f' count="{indice + 1}">'
        self._estilos = f"{xml[:m.start()]}{apertura}{m.group(1)}{fragmento}</{p}{bloque}>{xml[m.end():]}"
        return indice

    def _num_fmt(self, codigo: Optional[str]) -> int:
        if not codigo or codigo == "General":
            return 0
        ids = [int(i) for i in re.findall(r'numFmtId="(\d+)"', self._estilos)]
        nuevo = max([163] + ids) + 1
        self._agregar("numFmts", f'<{self._p}numFmt numFmtId="{nuevo}" formatCode="{texto_xml(codigo)}"/>')
        return nuevo

    def estilo(self, font: str, fill: str, border: str, alignment: str = "", num_fmt: Optional[str] = None) -> int:
        """Registra un xf (fragmentos SIN prefijo; se adaptan al del archivo) y devuelve su índice para s="..."."""
        p = self._p
        pre = (lambda x: re.sub(r"<(/?)(\w)", rf"<\1{p}\2", x)) if p else (lambda x: x)
        fid, flid, bid = self._agregar("fonts", pre(font)), self._agregar("fills", pre(fill)), self._agregar("borders", pre(border))
        nid = self._num_fmt(num_fmt)
        aplica = ' applyAlignment="1"' if alignment else ""
        cuerpo = f">{pre(alignment)}</{p}xf>" if alignment else "/>"
        return self._agregar("cellXfs", f'<{p}xf numFmtId="{nid}" fontId="{fid}" fillId="{flid}" borderId="{bid}" '
                                        f'xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" '
                                        f'applyBorder="1"{aplica}{cuerpo}')

    def dxf(self, fragmento: str) -> int:
        p = self._p
        pre = (lambda x: re.sub(r"<(/?)(\w)", rf"<\1{p}\2", x)) if p else (lambda x: x)
        return self._agregar("dxfs", pre(fragmento))

    # ---------------------------------------------------------------- hojas
    def escribir_hoja(self, nombre: str, partes: Iterable[str]) -> None:
        """Reemplaza la hoja `nombre` (o la agrega al final) con el XML de worksheet dado por partes."""
        wb, rels_nombre = self._leer("xl/workbook.xml"), "xl/_rels/workbook.xml.rels"
        rels = self._leer(rels_nombre)
        destino_de = {a["Id"]: a["Target"] for a in map(_attrs, re.findall(r"<(?:\w+:)?Relationship\b[^>]*>", rels))}

        parte = None
        for tag in re.findall(r"<(?:\w+:)?sheet\b[^>]*>", wb):
            a = _attrs(tag)
            if a.get("name", "").lower() == nombre.lower():
                rid = next(v for k, v in a.items() if k.endswith(":id"))
                destino = destino_de[rid]
                parte = destino.lstrip("/") if destino.startswith("/") else f"xl/{destino}"
                parte = os.path.normpath(parte).replace("\\", "/")
                break

        if parte is None:
            n = 1
            while f"xl/worksheets/sheet{n}.xml" in self._nombres:
                n += 1
            parte = f"xl/worksheets/sheet{n}.xml"
            ids_rel = {int(i) for i in re.findall(r'Id="rId(\d+)"', rels)}
            rid = f"rId{max(ids_rel | {0}) + 1}"
            rels = rels.replace("</Relationships>",
                                f'<Relationship Id="{rid}" Type="{REL_HOJA}" Target="worksheets/sheet{n}.xml"/></Relationships>')
            ejemplo = re.search(r"<(\w+:)?sheet\b[^>]*?\s(\w+):id=", wb)
            p_wb, p_r = (ejemplo.group(1) or "", ejemplo.group(2)) if ejemplo else ("", "r")
            sheet_id = max([int(i) for i in re.findall(r'sheetId="(\d+)"', wb)] + [0]) + 1
            wb = wb.replace(f"</{p_wb}sheets>",
                            f'<{p_wb}sheet xmlns:{p_r}="{NS_REL}" name="{texto_xml(nombre)}" sheetId="{sheet_id}" '
                            f'{p_r}:id="{rid}"/></{p_wb}sheets>')
            tipos = self._leer("[Content_Types].xml").replace(
                "</Types>", f'<Override PartName="/{parte}" ContentType="{TIPO_HOJA}"/></Types>')
            self._mod.update({"xl/workbook.xml": wb.encode(), rels_nombre: rels.encode(),
                              "[Content_Types].xml": tipos.encode()})
        else:
            carpeta, archivo = parte.rsplit("/", 1)
            self._omitir.add(f"{carpeta}/_rels/{archivo}.rels")

        self._omitir.add(parte)
        self._hoja = (parte, partes)

    def _guardar(self) -> None:
        self._mod["xl/styles.xml"] = self._estilos.encode()
        parte, partes = self._hoja
        for datos in self._mod.values():
            ElementTree.fromstring(datos)  # valida XML modificado antes de tocar el archivo
        fd, tmp = tempfile.mkstemp(suffix=self.ruta.suffix, dir=self.ruta.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zout:
                for item in self._zip.infolist():
                    if item.filename in self._omitir or item.filename in self._mod:
                        continue
                    zout.writestr(item, self._zip.read(item.filename), compress_type=item.compress_type)
                for nombre, datos in self._mod.items():
                    zout.writestr(nombre, datos)
                with zout.open(parte, "w", force_zip64=True) as f:
                    buffer = []
                    for trozo in partes:
                        buffer.append(trozo)
                        if len(buffer) >= 2000:
                            f.write("".join(buffer).encode("utf-8"))
                            buffer.clear()
                    f.write("".join(buffer).encode("utf-8"))
            self._zip.close()
            shutil.copymode(self.ruta, tmp)
            os.replace(tmp, self.ruta)
        except BaseException:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise