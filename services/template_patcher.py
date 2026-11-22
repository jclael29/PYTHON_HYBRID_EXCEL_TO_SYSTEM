"""
LOSSLESS TEMPLATE PATCHER
Direktang ine-edit ang XML ng sheet sa loob ng .xlsx, kaya hindi nawawala
ang images, logo, shapes, merged cells at format ng template.
"""
import io
import re
import zipfile
import posixpath
import xml.etree.ElementTree as ET

from openpyxl.utils import column_index_from_string, get_column_letter

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _q(tag):
    return "{%s}%s" % (NS_MAIN, tag)


def _split(coord):
    m = _CELL_RE.match(coord)
    return column_index_from_string(m.group(1)), int(m.group(2))


class TemplatePatcher:
    def __init__(self, template_path):
        self._load(template_path)
        for want, idx in (("FRONT", 0), ("BACK", 1)):
            self._resolve(want, idx)

    def _load(self, template_path):
        with zipfile.ZipFile(template_path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
            self.order = z.namelist()
        self.sheet_paths = self._sheet_paths()

    def _sheet_paths(self):
        wb = ET.fromstring(self.parts["xl/workbook.xml"])
        rels = ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        target = {r.get("Id"): r.get("Target") for r in rels}
        out = []
        for sh in wb.find(_q("sheets")):
            t = target[sh.get("{%s}id" % NS_REL)]
            path = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
            out.append((sh.get("name"), path))
        return out

    def _resolve(self, name, idx):
        def key(n):
            return re.sub(r"[^A-Z0-9]", "", n.upper())   # "BACK (2)" = "BACK 2" = "BACK2"
        names = {key(n): p for n, p in self.sheet_paths}
        if name == "BACK":
            # BACK (2): sheet na may ganyang pangalan, o kung apat ang sheet -> ang pang-apat
            if "BACK2" in names:
                return names["BACK2"]
            if len(self.sheet_paths) == 4:
                return self.sheet_paths[3][1]
        if key(name) in names:
            return names[key(name)]
        if idx < len(self.sheet_paths):
            return self.sheet_paths[idx][1]
        raise KeyError(f'Walang sheet na "{name}" sa template')

    def _patch_sheet(self, xml_bytes, writes):
        # i-register lahat ng namespace prefix para hindi magbago
        for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
            try:
                ET.register_namespace(prefix, uri)
            except ValueError:
                pass
        ET.register_namespace("", NS_MAIN)
        root = ET.fromstring(xml_bytes)
        # mc:Ignorable na tumutukoy sa prefix na hindi na ilalabas -> tanggalin
        for k in list(root.attrib):
            if k.endswith("}Ignorable"):
                del root.attrib[k]

        sd = root.find(_q("sheetData"))
        merges = []
        mc = root.find(_q("mergeCells"))
        if mc is not None:
            for m in mc:
                a, b = m.get("ref").split(":")
                (c1, r1), (c2, r2) = _split(a), _split(b)
                merges.append((c1, r1, c2, r2))

        rows = {int(r.get("r")): r for r in sd.findall(_q("row"))}

        def anchor(col, row):
            for c1, r1, c2, r2 in merges:
                if c1 <= col <= c2 and r1 <= row <= r2:
                    return c1, r1
            return col, row

        for coord, value in writes.items():
            col, row = anchor(*_split(coord))
            ref = f"{get_column_letter(col)}{row}"
            rel = rows.get(row)
            if rel is None:
                rel = ET.Element(_q("row"), {"r": str(row)})
                after = [n for n in rows if n < row]
                if after:
                    idx = list(sd).index(rows[max(after)]) + 1
                    sd.insert(idx, rel)
                else:
                    sd.insert(0, rel)
                rows[row] = rel
            cell = None
            insert_at = len(rel)
            for c in rel.findall(_q("c")):
                cc, _ = _split(c.get("r"))
                if cc == col:
                    cell = c
                    break
                if cc > col:
                    insert_at = list(rel).index(c)
                    break
            if cell is None:
                cell = ET.Element(_q("c"), {"r": ref})
                rel.insert(insert_at, cell)
            for child in list(cell):
                cell.remove(child)
            cell.attrib.pop("t", None)
            if isinstance(value, bool):
                value = str(value)
            if isinstance(value, (int, float)):
                ET.SubElement(cell, _q("v")).text = repr(value)
            else:
                cell.set("t", "inlineStr")
                is_ = ET.SubElement(cell, _q("is"))
                t = ET.SubElement(is_, _q("t"))
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = str(value)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _write_zip(self, patched, out):
        """patched: {sheet xml path: {cell: value}}. out: path o file-like."""
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.order:
                data = self.parts[name]
                if name in patched:
                    data = self._patch_sheet(data, patched[name])
                elif name == "xl/calcChain.xml":
                    continue  # iiwas error kung nagbago ang formula cells
                elif name == "[Content_Types].xml":
                    data = data.replace(
                        b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', b"")
                elif name == "xl/_rels/workbook.xml.rels":
                    data = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"", data)
                z.writestr(name, data)

    def build(self, front_writes, back_writes, out_path):   # out_path: path o file-like
        self._write_zip({
            self._resolve("FRONT", 0): front_writes,
            self._resolve("BACK", 1): back_writes,
        }, out_path)

    def build_bytes(self, front_writes, back_writes):
        """Gumawa ng SF10 sa memory (walang isinusulat sa disk)."""
        buf = io.BytesIO()
        self.build(front_writes, back_writes, buf)
        return buf.getvalue()


class SingleSheetPatcher(TemplatePatcher):
    """Para sa template na isa lang ang sheet (Form 9): ang unang sheet ang sinusulatan."""

    def __init__(self, template_path):
        self._load(template_path)
        if not self.sheet_paths:
            raise KeyError("Walang sheet ang template")

    def build_bytes(self, writes):
        buf = io.BytesIO()
        self._write_zip({self.sheet_paths[0][1]: writes}, buf)
        return buf.getvalue()