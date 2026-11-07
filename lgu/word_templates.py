"""
Custom Word (.docx) template auto-fill for LGU records.

Templates are ordinary .docx files containing Jinja-style placeholders such as
{{ mayor_name }} or {{ seal }} (seal image). They are stored in word_templates/
next to the data file and filled with docxtpl.
"""

import datetime
import io
import os
import re
import zipfile
from typing import Any, Dict, Iterable, List, Optional, Tuple

# Placeholder reference shown in the UI and in the sample template: (name, description)
PLACEHOLDERS: List[Tuple[str, str]] = [
    ("mayor_name", "Mayor full name (uses the Name Casing setting)"),
    ("mayor_name_upper", "Mayor full name in UPPERCASE"),
    ("mayor_first", "Mayor first name"),
    ("mayor_middle", "Mayor middle name"),
    ("mayor_last", "Mayor last name"),
    ("mayor_suffix", "Mayor suffix (Jr., III, ...)"),
    ("mayor_title", "\"City Mayor\" or \"Municipal Mayor\""),
    ("mayor_email", "Mayor email"),
    ("mayor_contact", "Mayor contact number"),
    ("vice_mayor", "Vice mayor full name"),
    ("vice_mayor_upper", "Vice mayor full name in UPPERCASE"),
    ("vice_mayor_title", "\"City Vice Mayor\" or \"Municipal Vice Mayor\""),
    ("vm_email", "Vice mayor email"),
    ("vm_contact", "Vice mayor contact number"),
    ("town", "Town / city name without parentheses, e.g. \"Calanasan\", \"City Of Lamitan\""),
    ("town_upper", "Town / city name without parentheses, in UPPERCASE"),
    ("town_name", "Name without \"City of\", e.g. \"Lamitan\""),
    ("town_name_upper", "Clean name in UPPERCASE, e.g. \"LAMITAN\""),
    ("lgu_name", "\"Municipality of Akbar\", \"City of Lamitan\", \"Quezon City\""),
    ("lgu_name_upper", "\"MUNICIPALITY OF AKBAR\", \"CITY OF LAMITAN\", \"QUEZON CITY\""),
    ("office", "\"Office of the Municipal Mayor\" or \"Office of the City Mayor\""),
    ("office_upper", "\"OFFICE OF THE MUNICIPAL MAYOR\" / \"OFFICE OF THE CITY MAYOR\""),
    ("province", "Province"),
    ("province_upper", "Province in UPPERCASE"),
    ("place", "\"Town, Province\""),
    ("region", "Region"),
    ("lgu_type", "\"City\" or \"Municipality\""),
    ("address", "LGU hall address"),
    ("date", "Today's date, e.g. November 7, 2026"),
    ("seal", "Seal image (blank if no seal saved)"),
]
KNOWN_NAMES = {name for name, _ in PLACEHOLDERS}

# Image formats Word / python-docx can embed
DOCX_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"}

MAX_TEMPLATE_BYTES = 15 * 1024 * 1024


def _safe_filename(text: str) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", text)
    return re.sub(r"\s+", " ", text).strip(" .") or "document"


def _require_docxtpl():
    try:
        import docxtpl  # noqa: F401
        return docxtpl
    except ImportError as ex:
        raise RuntimeError(
            "Word template support needs the 'docxtpl' package. Run: pip install -r requirements.txt"
        ) from ex


def strip_parens(text: str) -> str:
    """Remove "(...)" parts and tidy spacing: "Calanasan (bayag)" -> "Calanasan"."""
    text = re.sub(r"\s*\([^)]*\)", "", text or "")
    return re.sub(r"\s{2,}", " ", text).strip()


def build_context(record: Dict[str, str], raw: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Build the placeholder values for one LGU (without the seal image)."""
    raw = raw or {}
    m = raw.get("raw_mayor") or {}
    vm = raw.get("raw_vm") or {}
    is_city = record.get("LGU Type") == "City"
    mayor = record.get("Mayor Name", "") or ""
    vice = record.get("Vice Mayor", "") or ""
    # Drop parenthetical notes: "Calanasan (bayag)" -> "Calanasan", "City Of Lamitan (capital)" -> "City Of Lamitan"
    town = strip_parens(record.get("Town", "") or "")
    province = record.get("Province", "") or ""
    today = datetime.date.today()

    def cap(s: str) -> str:
        return " ".join(w.capitalize() for w in (s or "").split())

    # "City Of Lamitan (capital)" -> "Lamitan"; "Quezon City" stays "Quezon City"
    town_name = re.sub(r"^city\s+of\s+", "", town, flags=re.I).strip()
    if is_city:
        lgu_name = town_name if re.search(r"\bcity$", town_name, re.I) else f"City of {town_name}"
    else:
        lgu_name = f"Municipality of {town_name}"
    office = "Office of the City Mayor" if is_city else "Office of the Municipal Mayor"

    return {
        "mayor_name": mayor,
        "mayor_name_upper": mayor.upper(),
        "mayor_first": cap(m.get("first", "")),
        "mayor_middle": cap(m.get("middle", "")),
        "mayor_last": cap(m.get("last", "")),
        "mayor_suffix": m.get("suffix", "") or "",
        "mayor_title": "City Mayor" if is_city else "Municipal Mayor",
        "mayor_email": record.get("Mayor Email", "") or m.get("email", "") or "",
        "mayor_contact": record.get("Mayor Contact", "") or m.get("contact", "") or "",
        "vice_mayor": vice,
        "vice_mayor_upper": vice.upper(),
        "vice_mayor_title": "City Vice Mayor" if is_city else "Municipal Vice Mayor",
        "vm_email": record.get("Vice Mayor Email", "") or vm.get("email", "") or "",
        "vm_contact": record.get("Vice Mayor Contact", "") or vm.get("contact", "") or "",
        "town": town,
        "town_upper": town.upper(),
        "town_name": town_name,
        "town_name_upper": town_name.upper(),
        "lgu_name": lgu_name,
        "lgu_name_upper": lgu_name.upper(),
        "office": office,
        "office_upper": office.upper(),
        "province": province,
        "province_upper": province.upper(),
        "place": f"{town}, {province}",
        "region": record.get("Region", "") or "",
        "lgu_type": "City" if is_city else "Municipality",
        "address": record.get("Municipal Hall Address", "") or "",
        "date": f"{today.strftime('%B')} {today.day}, {today.year}",
        "seal": "",
    }


class WordTemplateStore:
    def __init__(self, base_dir: str):
        self.dir = os.path.join(base_dir, "word_templates")

    # ---------- storage ----------
    def _path(self, name: str) -> str:
        name = os.path.basename(name or "")
        if not name.lower().endswith(".docx"):
            raise ValueError("Template must be a .docx file")
        path = os.path.join(self.dir, name)
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Template not found: {name}")
        return path

    def placeholders_in(self, path_or_bytes) -> List[str]:
        docxtpl = _require_docxtpl()
        src = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else path_or_bytes
        tpl = docxtpl.DocxTemplate(src)
        return sorted(tpl.get_undeclared_template_variables())

    def list(self) -> List[Dict[str, Any]]:
        if not os.path.isdir(self.dir):
            return []
        items = []
        for fn in sorted(os.listdir(self.dir), key=str.lower):
            if not fn.lower().endswith(".docx") or fn.startswith("~$"):
                continue
            path = os.path.join(self.dir, fn)
            info: Dict[str, Any] = {
                "name": fn,
                "size": os.path.getsize(path),
                "updated": datetime.datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="seconds"),
            }
            try:
                found = self.placeholders_in(path)
                info["placeholders"] = found
                info["unknown"] = [p for p in found if p not in KNOWN_NAMES]
            except Exception as ex:
                info["error"] = str(ex)
            items.append(info)
        return items

    def save(self, filename: str, content: bytes) -> Dict[str, Any]:
        filename = _safe_filename(os.path.basename(filename or ""))
        if not filename.lower().endswith(".docx"):
            raise ValueError("Only .docx Word files are supported (save .doc files as .docx first)")
        if len(content) > MAX_TEMPLATE_BYTES:
            raise ValueError("Template too large (max 15 MB)")
        if not zipfile.is_zipfile(io.BytesIO(content)):
            raise ValueError("File is not a valid .docx document")
        found = self.placeholders_in(content)  # also validates the Jinja syntax
        os.makedirs(self.dir, exist_ok=True)
        tmp = os.path.join(self.dir, filename + ".tmp")
        with open(tmp, "wb") as f:
            f.write(content)
        os.replace(tmp, os.path.join(self.dir, filename))
        return {"name": filename, "placeholders": found, "unknown": [p for p in found if p not in KNOWN_NAMES]}

    def delete(self, name: str):
        os.remove(self._path(name))

    # ---------- filling ----------
    def _render(self, template_path: str, context: Dict[str, Any], seal_path: Optional[str], seal_width_in: float) -> bytes:
        docxtpl = _require_docxtpl()
        from docx.shared import Inches

        tpl = docxtpl.DocxTemplate(template_path)
        ctx = dict(context)
        if seal_path and os.path.isfile(seal_path) and seal_path.rsplit(".", 1)[-1].lower() in DOCX_IMAGE_EXTS:
            ctx["seal"] = docxtpl.InlineImage(tpl, seal_path, width=Inches(seal_width_in))
        tpl.render(ctx, autoescape=True)
        out = io.BytesIO()
        tpl.save(out)
        return out.getvalue()

    def fill_one(self, name: str, record: Dict[str, str], raw: Optional[Dict[str, Any]],
                 seal_path: Optional[str], seal_width_in: float = 1.0) -> Tuple[bytes, str]:
        path = self._path(name)
        content = self._render(path, build_context(record, raw), seal_path, seal_width_in)
        stem = os.path.splitext(os.path.basename(path))[0]
        filename = _safe_filename(f"{stem} - {strip_parens(record.get('Town', ''))}, {record.get('Province', '')}") + ".docx"
        return content, filename

    def _merge(self, docs: List[bytes]) -> bytes:
        """
        Concatenate filled documents into one file separated by page breaks
        (no section breaks). Because a single-section document can only have one
        header, each LGU's header (letterhead with its own seal/town) is moved into
        the top of that LGU's page. Footers come from the first document.
        """
        from docx import Document
        try:
            from docxcompose.composer import Composer
        except ImportError as ex:
            raise RuntimeError(
                "Merging needs the 'docxcompose' package. Run: pip install -r requirements.txt"
            ) from ex
        if len(docs) == 1:
            return docs[0]
        master = Document(io.BytesIO(docs[0]))
        self._inline_header(master)
        composer = Composer(master)
        for content in docs[1:]:
            doc = Document(io.BytesIO(content))
            self._inline_header(doc)
            master.add_page_break()
            composer.append(doc)
        out = io.BytesIO()
        composer.save(out)
        return out.getvalue()

    @staticmethod
    def _inline_header(doc):
        """Move the document's page header into the start of its body and remove the
        header, keeping images. The top margin is set to the header distance so the
        letterhead sits where the header used to be."""
        from copy import deepcopy
        from lxml import etree
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml.ns import nsmap, qn

        body = doc.element.body
        sect = body.find(qn("w:sectPr"))
        if sect is None:
            return
        refs = {r.get(qn("w:type")) or "default": r for r in sect.findall(qn("w:headerReference"))}
        if not refs:
            return
        use_first = sect.find(qn("w:titlePg")) is not None and "first" in refs
        ref = refs["first"] if use_first else refs.get("default")
        for r in refs.values():
            sect.remove(r)
        if ref is None:
            return

        hdr_part = doc.part.related_parts[ref.get(qn("r:id"))]
        remap: Dict[str, str] = {}
        for rid, rel in hdr_part.rels.items():
            if rel.is_external:
                remap[rid] = doc.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.IMAGE:
                remap[rid] = doc.part.relate_to(rel.target_part, RT.IMAGE)

        r_ns = "{%s}" % nsmap["r"]
        anchor = body[0]
        for el in hdr_part.element:
            if el.tag not in (qn("w:p"), qn("w:tbl"), qn("w:sdt")):
                continue
            el = deepcopy(el)
            for node in el.iter(etree.Element):
                for attr, val in list(node.attrib.items()):
                    if attr.startswith(r_ns) and val in remap:
                        node.set(attr, remap[val])
            anchor.addprevious(el)

        pg_mar = sect.find(qn("w:pgMar"))
        if pg_mar is not None and pg_mar.get(qn("w:header")):
            pg_mar.set(qn("w:top"), pg_mar.get(qn("w:header")))

    def fill_batch(self, name: str,
                   items: Iterable[Tuple[Dict[str, str], Optional[Dict[str, Any]], Optional[str]]],
                   seal_width_in: float = 1.0, mode: str = "single",
                   label: str = "") -> Tuple[bytes, str, str, int, int]:
        """
        Fill the template for every item (record, raw, seal_path), one LGU per page,
        sorted by province then town.
          mode="single":   everything merged into ONE .docx
          mode="province": one merged .docx per province (a single province is returned
                           as that .docx; several provinces come as a ZIP)
        `label` names the single file (e.g. the selected region); defaults to the province(s).
        Returns (content, filename, mime, lgu_count, province_count).
        """
        path = self._path(name)
        stem = os.path.splitext(os.path.basename(path))[0]
        docx_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        groups: Dict[str, List[Tuple[Dict[str, str], Optional[Dict[str, Any]], Optional[str]]]] = {}
        for item in items:
            groups.setdefault(item[0].get("Province", "") or "Unknown", []).append(item)
        if not groups:
            raise ValueError("No LGUs to fill")
        provinces = sorted(groups, key=str.lower)
        for province in provinces:
            groups[province].sort(key=lambda it: (it[0].get("Town", "") or "").lower())

        def render_all(rows) -> List[bytes]:
            return [self._render(path, build_context(rec, raw), seal, seal_width_in) for rec, raw, seal in rows]

        if mode != "province" or len(provinces) == 1:
            rows = [it for p in provinces for it in groups[p]]
            content = self._merge(render_all(rows))
            if len(provinces) == 1:
                label = provinces[0]
            label = label or f"{len(provinces)} provinces"
            return content, _safe_filename(f"{stem} - {label}") + ".docx", docx_mime, len(rows), len(provinces)

        merged: List[Tuple[str, bytes]] = []
        total = 0
        for province in provinces:
            docs = render_all(groups[province])
            merged.append((province, self._merge(docs)))
            total += len(docs)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for province, content in merged:
                zf.writestr(_safe_filename(f"{stem} - {province}") + ".docx", content)
        filename = _safe_filename(f"{stem} - {len(merged)} provinces") + ".zip"
        return buf.getvalue(), filename, "application/zip", total, len(merged)

    @staticmethod
    def sample_template() -> bytes:
        """A starter .docx showing a seal header, a sample letter and the placeholder reference."""
        _require_docxtpl()
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        from docx.enum.table import WD_ALIGN_VERTICAL
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Inches

        doc = Document()
        section = doc.sections[0]
        usable = section.page_width - section.left_margin - section.right_margin

        # Letterhead in the Word header: a borderless 1x3 table keeps seal | text | logo
        # aligned without tabs or spaces, and each LGU keeps its own header when merged.
        header = section.header
        table = header.add_table(rows=1, cols=3, width=usable)
        table.autofit = False
        side = Inches(1.3)
        widths = [side, usable - 2 * side, side]
        for cell, w in zip(table.rows[0].cells, widths):
            cell.width = w
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        left, mid, right = table.rows[0].cells
        left.paragraphs[0].text = "{{ seal }}"
        left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        lines = [("Republic of the Philippines", False), ("Province of {{ province }}", False),
                 ("{{ lgu_name_upper }}", True), ("{{ office_upper }}", False)]
        mp = mid.paragraphs[0]
        mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for i, (text, bold) in enumerate(lines):
            run = mp.add_run(text)
            run.bold = bold
            if i < len(lines) - 1:
                run.add_break()
        right.paragraphs[0].text = "(insert your logo here)"
        right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Horizontal rule under the letterhead (paragraph bottom border, not a drawn line)
        rule = header.paragraphs[0] if not header.paragraphs[0].text else header.add_paragraph()
        header._element.append(rule._p)  # make sure the rule sits below the table
        pPr = rule._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        for k, v in {"val": "single", "sz": "8", "space": "1", "color": "000000"}.items():
            bottom.set(qn(f"w:{k}"), v)
        pbdr.append(bottom)
        pPr.append(pbdr)

        doc.add_paragraph("{{ date }}")
        p = doc.add_paragraph()
        p.add_run("HON. {{ mayor_name_upper }}").bold = True
        p.add_run("\n{{ mayor_title }}\n{{ place }}\n{{ address }}")
        doc.add_paragraph("Dear Mayor {{ mayor_last }},")
        doc.add_paragraph("Write your letter here. Any placeholder below can be used anywhere in the document, "
                          "including headers, footers and tables.")

        doc.add_page_break()
        doc.add_heading("Placeholder reference (delete this page)", level=2)
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        table.rows[0].cells[0].text = "Type this in your template"
        table.rows[0].cells[1].text = "Replaced with"
        for name, desc in PLACEHOLDERS:
            cells = table.add_row().cells
            # Wrap in {% raw %} so the reference itself is not filled in
            cells[0].text = "{% raw %}{{ " + name + " }}{% endraw %}"
            cells[1].text = desc
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)
        out = io.BytesIO()
        doc.save(out)
        return out.getvalue()
