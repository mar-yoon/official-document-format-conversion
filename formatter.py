"""Dependency-free, conservative DOCX formatting engine.

Each template is data. The classifier may be replaced or extended for a thesis
template without changing the upload/download application.
"""
from __future__ import annotations

import io
import json
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
for prefix, url in {
    "w": W,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}.items():
    ET.register_namespace(prefix, url)

MAX_FILE = 15 * 1024 * 1024
MAX_UNCOMPRESSED = 100 * 1024 * 1024
MAX_MEMBERS = 500
TEMPLATE_DIR = Path(__file__).parent / "templates"


class FormatError(ValueError):
    pass


def q(name: str) -> str:
    return f"{{{W}}}{name}"


def mm_twips(mm: float) -> str:
    return str(round(mm * 1440 / 25.4))


def pt_half(pt: float) -> str:
    return str(round(pt * 2))


def load_templates() -> List[Dict]:
    templates = []
    for path in sorted(TEMPLATE_DIR.glob("*.json"), key=lambda item: (item.stem != "gov_official", item.stem)):
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("id") != path.stem:
            raise FormatError(f"模板 ID 与文件名不符：{path.name}")
        templates.append(data)
    return templates


def get_template(template_id: str) -> Dict:
    for template in load_templates():
        if template["id"] == template_id:
            return template
    raise FormatError("未找到所选模板")


def read_package(data: bytes) -> Dict[str, bytes]:
    if len(data) > MAX_FILE:
        raise FormatError("文件超过 15 MB 上限")
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise FormatError("文件不是有效的 .docx")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            infos = zf.infolist()
            if len(infos) > MAX_MEMBERS or sum(i.file_size for i in infos) > MAX_UNCOMPRESSED:
                raise FormatError("文档包过大或结构异常")
            names = {i.filename for i in infos}
            if "word/document.xml" not in names or "[Content_Types].xml" not in names:
                raise FormatError("缺少 Word 文档主体")
            if any("vbaProject" in n or "activeX" in n for n in names):
                raise FormatError("暂不处理包含宏或 ActiveX 的文档")
            if len(names) != len(infos):
                raise FormatError("文档包包含重复条目")
            return {i.filename: zf.read(i) for i in infos}
    except (zipfile.BadZipFile, RuntimeError, OSError) as exc:
        raise FormatError("无法读取文档包") from exc


def paragraphs(root: ET.Element) -> List[ET.Element]:
    body = root.find("w:body", NS)
    if body is None:
        raise FormatError("文档缺少正文")
    return list(body.iter(q("p")))


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(q("t"))).strip()


def classify(items: List[ET.Element], skip_indices=None, attachment_label=False) -> List[str]:
    texts = [paragraph_text(p) for p in items]
    skip_indices = skip_indices or set()
    nonempty = [i for i, text in enumerate(texts) if text and i not in skip_indices]
    roles = ["empty"] * len(items)
    if not nonempty:
        return roles

    first = nonempty[0]
    title_index = first
    if re.search(r"文件$", texts[first]) and len(texts[first]) < 45:
        roles[first] = "agency"
        title_index = next((i for i in nonempty[1:] if not re.search(r"〔\d{4}〕\d+号", texts[i])), first)
        for i in nonempty[1:]:
            if i >= title_index:
                break
            if re.search(r"〔\d{4}〕\d+号", texts[i]):
                roles[i] = "doc_number"
    roles[title_index] = "title"

    for i in nonempty:
        if roles[i] != "empty":
            continue
        text = texts[i]
        if attachment_label and re.fullmatch(r"附件\s*[一二三四五六七八九十百0-9]*[：:]?", text):
            roles[i] = "attachment_label"
        elif re.match(r"^附件(?:\s*[一二三四五六七八九十百0-9]+)?[：:]", text):
            roles[i] = "attachment"
        elif re.match(r"^[一二三四五六七八九十百]+、", text):
            roles[i] = "heading_1"
        elif re.match(r"^（[一二三四五六七八九十百]+）", text):
            roles[i] = "heading_2"
        elif re.match(r"^\d+[.．]", text):
            roles[i] = "heading_3"
        elif re.match(r"^（\d+）", text):
            roles[i] = "heading_4"
        elif i == next((j for j in nonempty if j > title_index), -1) and text.endswith(("：", ":")):
            roles[i] = "recipient"
        elif re.fullmatch(r"\d{4}年\d{1,2}月\d{1,2}日", text):
            roles[i] = "date"
        else:
            roles[i] = "body"

    date_indices = [i for i, role in enumerate(roles) if role == "date"]
    if date_indices:
        date_i = date_indices[-1]
        previous = next((i for i in reversed(nonempty) if i < date_i), None)
        if previous is not None and roles[previous] == "body" and len(texts[previous]) <= 30:
            roles[previous] = "signature"
    return roles


def classify_document(root: ET.Element, items: List[ET.Element], template: Dict) -> List[str]:
    if "table" not in template["roles"] and "image" not in template["roles"]:
        return classify(items)
    table_ids = {id(p) for table in root.iter(q("tbl")) for p in table.iter(q("p"))}
    table_indices = {i for i, p in enumerate(items) if id(p) in table_ids}
    roles = classify(items, table_indices, "attachment_label" in template["roles"])
    if "table" in template["roles"]:
        for i in table_indices:
            roles[i] = "table"
    if "image" in template["roles"]:
        for i, p in enumerate(items):
            if roles[i] == "empty" and p.find(".//w:drawing", NS) is not None:
                roles[i] = "image"
    return roles


def inspect(data: bytes, template_id: str = "gov_official") -> Dict:
    template = get_template(template_id)
    package = read_package(data)
    try:
        root = ET.fromstring(package["word/document.xml"])
    except ET.ParseError as exc:
        raise FormatError("文档主体 XML 损坏") from exc
    items = paragraphs(root)
    roles = classify_document(root, items, template)
    counts = {role: roles.count(role) for role in sorted(set(roles)) if role != "empty"}
    title = next((paragraph_text(p) for p, role in zip(items, roles) if role == "title"), "")
    warnings = []
    if not title:
        warnings.append("没有检测到正文文字，请检查上传的文件。")
    if counts.get("agency", 0) == 0:
        warnings.append("未识别发文机关标志；此程序不会补写缺失的公文要素。")
    if counts.get("doc_number", 0) == 0:
        warnings.append("未识别发文字号；请按实际发文信息人工核对。")
    if counts.get("date", 0) == 0:
        warnings.append("未识别成文日期；请人工核对。")
    if not any(name.startswith("word/footer") and name.endswith(".xml") for name in package):
        warnings.append("原文没有页脚；首版不会自动生成符合单双页位置要求的页码。")
    if any("<w:drawing" in part.decode("utf-8", "ignore") for name, part in package.items() if name.startswith("word/") and name.endswith(".xml")):
        warnings.append("文档含图片或图形，调整后请检查位置与印章。")
    if root.find(".//w:tbl", NS) is not None:
        warnings.append("文档含表格，调整后请检查跨页和表格宽度。")
    if template.get("id") == "enterprise_training":
        warnings.append("附件分页、图片环绕及页码插入需在 Word 中复核；程序不会补写或删除正文内容。")
    return {"template": template["name"], "basis": template["basis"], "title": title, "paragraphs": len(items), "roles": counts, "warnings": warnings}


def ensure(parent: ET.Element, name: str, first: bool = False) -> ET.Element:
    found = parent.find(q(name))
    if found is None:
        found = ET.Element(q(name))
        if first:
            parent.insert(0, found)
        else:
            parent.append(found)
    return found


def set_child_attr(parent: ET.Element, name: str, attr: str, value: str) -> None:
    ensure(parent, name).set(q(attr), value)


def set_format(p: ET.Element, role: str, template: Dict) -> None:
    defaults = template["default"]
    settings = dict(defaults, **template["roles"].get(role, {}))
    if not settings.get("preserve_paragraph"):
        ppr = ensure(p, "pPr", first=True)
        set_child_attr(ppr, "spacing", "before", str(round(settings.get("space_before_pt", 0) * 20)))
        set_child_attr(ppr, "spacing", "after", "0")
        if settings.get("line_rule") == "auto":
            set_child_attr(ppr, "spacing", "line", "240")
            set_child_attr(ppr, "spacing", "lineRule", "auto")
        else:
            set_child_attr(ppr, "spacing", "line", str(round(settings["line_pt"] * 20)))
            set_child_attr(ppr, "spacing", "lineRule", "exact")
        indent = ensure(ppr, "ind")
        indent.set(q("left"), "0")
        indent.set(q("firstLineChars"), str(settings.get("first_line_chars", 0) * 100))
        for obsolete in ("firstLine", "hanging", "hangingChars"):
            indent.attrib.pop(q(obsolete), None)
        set_child_attr(ppr, "jc", "val", settings.get("align", "left"))

    # Apply to all text runs, including runs inside hyperlinks. Other run content
    # (images, fields, bookmarks) remains in place.
    for run in p.iter(q("r")):
        if run.find(q("t")) is None:
            continue
        rpr = ensure(run, "rPr", first=True)
        fonts = ensure(rpr, "rFonts")
        for key in ("ascii", "hAnsi"):
            fonts.set(q(key), settings.get("latin_font", settings["font"]))
        for key in ("eastAsia", "cs"):
            fonts.set(q(key), settings["font"])
        set_child_attr(rpr, "sz", "val", pt_half(settings["size_pt"]))
        set_child_attr(rpr, "szCs", "val", pt_half(settings["size_pt"]))
        set_child_attr(rpr, "color", "val", settings.get("color", "000000"))
        bold = ensure(rpr, "b")
        bold.set(q("val"), "1" if settings.get("bold", False) else "0")
        for obsolete in ("highlight", "shd", "u"):
            element = rpr.find(q(obsolete))
            if element is not None:
                rpr.remove(element)


def set_page(root: ET.Element, template: Dict) -> None:
    page = template["page"]
    section_props = list(root.iter(q("sectPr")))
    if not section_props:
        body = root.find("w:body", NS)
        section_props = [ET.SubElement(body, q("sectPr"))]
    for sect in section_props:
        size = ensure(sect, "pgSz", first=True)
        size.set(q("w"), mm_twips(page["width_mm"]))
        size.set(q("h"), mm_twips(page["height_mm"]))
        margin = ensure(sect, "pgMar")
        for key, source in (("top", "top_mm"), ("bottom", "bottom_mm"), ("left", "left_mm"), ("right", "right_mm")):
            margin.set(q(key), mm_twips(page[source]))


def set_existing_footer_page_font(package: Dict[str, bytes], template: Dict) -> None:
    settings = template.get("footer_page_number")
    if not settings:
        return
    for name, content in list(package.items()):
        if not re.fullmatch(r"word/footer\d+\.xml", name):
            continue
        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            continue
        changed = False
        for paragraph in root.iter(q("p")):
            instructions = " ".join((node.text or "") for node in paragraph.iter(q("instrText")))
            simple_fields = " ".join((node.get(q("instr")) or "") for node in paragraph.iter(q("fldSimple")))
            if not re.search(r"\bPAGE\b", instructions + " " + simple_fields, re.I):
                continue
            for run in paragraph.iter(q("r")):
                rpr = ensure(run, "rPr", first=True)
                fonts = ensure(rpr, "rFonts")
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    fonts.set(q(key), settings["font"])
                set_child_attr(rpr, "sz", "val", pt_half(settings["size_pt"]))
                set_child_attr(rpr, "szCs", "val", pt_half(settings["size_pt"]))
                changed = True
        if changed:
            package[name] = serialize_document(root, content)


def serialize_document(root: ET.Element, original: bytes) -> bytes:
    """Keep namespace declarations referenced by mc:Ignorable in Word files.

    ElementTree otherwise drops declarations for namespaces used only in that
    attribute's space-separated prefix list, which can make Word repair a file.
    """
    output = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    root_tag = re.search(rb"<[^!?][^>]*>", original)
    output_tag = re.search(rb"<[^!?][^>]*>", output)
    if not root_tag or not output_tag:
        return output
    original_decls = re.findall(rb"\s(xmlns(?::[A-Za-z_][\w.-]*)?)=([\"'][^\"']*[\"'])", root_tag.group())
    existing = {name for name, _ in re.findall(rb"\s(xmlns(?::[A-Za-z_][\w.-]*)?)=([\"'][^\"']*[\"'])", output_tag.group())}
    missing = b"".join(b" " + name + b"=" + value for name, value in original_decls if name not in existing)
    if not missing:
        return output
    insert_at = output_tag.end() - 1
    return output[:insert_at] + missing + output[insert_at:]


def format_docx(data: bytes, template_id: str = "gov_official") -> Tuple[bytes, Dict]:
    report = inspect(data, template_id)
    template = get_template(template_id)
    package = read_package(data)
    root = ET.fromstring(package["word/document.xml"])
    items = paragraphs(root)
    roles = classify_document(root, items, template)
    set_page(root, template)
    for p, role in zip(items, roles):
        if role != "empty":
            set_format(p, role, template)
    package["word/document.xml"] = serialize_document(root, package["word/document.xml"])
    set_existing_footer_page_font(package, template)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in package.items():
            zf.writestr(name, content)
    return out.getvalue(), report
