from __future__ import annotations

import re
import shutil
import zipfile
import csv
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
PAPER_DIR = ROOT / "\u8bba\u6587"
DOCX = PAPER_DIR / "\u8ba1\uff08\u521b\uff092302\u73ed-\u5218\u4e16\u7fd4-XXXXXXX-\u6570\u5b57\u56fe\u50cf\u5904\u7406\u6280\u672f\u7ed3\u8bfe\u8bba\u6587.docx"


FIG_CAPTION_PREFIXES = (
    "\u56fe1 DE-Cividis",
    "\u56fe1 \u7ec6\u8282\u589e\u5f3a",
    "\u56fe2 \u4e0d\u540c\u8272\u56fe",
    "\u56fe3 \u53c2\u6570",
    "\u56fe4 \u6700\u7ec8\u8bc4\u4f30\u96c6",
    "\u56fe5 X",
    "\u56fe6 DE-Cividis",
    "\u56fe6 \u65b9\u6cd5\u6d88\u878d",
    "\u56fe7 Kodak",
    "\u56fe7 \u8865\u5145\u9a8c\u8bc1",
    "\u56fe8 \u590d\u6742\u7eb9\u7406",
)

TABLE_CAPTION_PREFIXES = (
    "\u88681 \u6700\u7ec8\u8bc4\u4f30\u96c6",
    "\u88682 \u989d\u5916 Kodak",
)

FULLWIDTH_DIGITS = str.maketrans("0123456789", "０１２３４５６７８９")
HALFWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def is_figure_caption(text: str) -> bool:
    probe = re.sub(r"^([图表])\s+([0-9０-９]+)", r"\1\2", text.translate(HALFWIDTH_DIGITS))
    return probe.startswith(FIG_CAPTION_PREFIXES)


def is_table_caption(text: str) -> bool:
    probe = re.sub(r"^([图表])\s+([0-9０-９]+)", r"\1\2", text.translate(HALFWIDTH_DIGITS))
    return probe.startswith(TABLE_CAPTION_PREFIXES)


def compact_caption_number(paragraph) -> None:
    text = paragraph.text.strip()
    updated = re.sub(
        r"^([图表])\s*([0-9０-９]+)\s*",
        lambda match: match.group(1) + match.group(2).translate(HALFWIDTH_DIGITS) + " ",
        text,
    )
    if updated != text:
        set_text(paragraph, updated)


def normalize_body_reference_number(paragraph) -> None:
    text = paragraph.text.strip()
    updated = re.sub(
        r"^([图表])([０-９]+)",
        lambda match: match.group(1) + match.group(2).translate(HALFWIDTH_DIGITS),
        text,
    )
    if updated != text:
        set_text(paragraph, updated)


def set_text(paragraph, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def set_text_hard(paragraph, text: str) -> None:
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    paragraph.add_run(text)


def remove_paragraph(paragraph) -> None:
    p = paragraph._element
    p.getparent().remove(p)


def set_run_font(run, east: str = "\u5b8b\u4f53", latin: str = "Times New Roman", size: float | None = None,
                 bold: bool | None = None, italic: bool | None = None) -> None:
    run.font.name = latin
    if run._element.rPr is None:
        run._element.get_or_add_rPr()
    if run._element.rPr.rFonts is None:
        run._element.rPr.append(OxmlElement("w:rFonts"))
    run._element.rPr.rFonts.set(qn("w:ascii"), latin)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), latin)
    run._element.rPr.rFonts.set(qn("w:cs"), latin)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def style_paragraph(paragraph, *, east="\u5b8b\u4f53", latin="Times New Roman", size=10.5,
                    bold=None, italic=False, align=None, first_line=True,
                    line_spacing=1.25, before=0, after=0) -> None:
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(0.74) if first_line else None
    if align is not None:
        paragraph.alignment = align
    for run in paragraph.runs:
        if run.text:
            set_run_font(run, east=east, latin=latin, size=size, bold=bold, italic=italic)


def disable_east_asian_auto_spacing(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    for name in ("autoSpaceDE", "autoSpaceDN"):
        element = p_pr.find(qn("w:" + name))
        if element is None:
            element = OxmlElement("w:" + name)
            p_pr.append(element)
        element.set(qn("w:val"), "0")


CITATION_RE = re.compile(r"\[((?:[1-9]\d?)(?:[-,，](?:[1-9]\d?))*)\]")


def _reference_number(text: str) -> int | None:
    match = re.match(r"^\[(\d{1,2})\]", text.strip())
    return int(match.group(1)) if match else None


def _existing_bookmark_names(doc: Document) -> set[str]:
    names: set[str] = set()
    for start in doc._element.iter(qn("w:bookmarkStart")):
        name = start.get(qn("w:name"))
        if name:
            names.add(name)
    return names


def _next_bookmark_id(doc: Document) -> int:
    ids = []
    for start in doc._element.iter(qn("w:bookmarkStart")):
        value = start.get(qn("w:id"))
        if value and value.isdigit():
            ids.append(int(value))
    return max(ids, default=0) + 1


def _add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _add_text_run(paragraph, text: str, *, superscript: bool = False) -> None:
    if not text:
        return
    run = paragraph.add_run(text)
    set_run_font(run, east="\u5b8b\u4f53", latin="Times New Roman", size=9 if superscript else 10.5, bold=False)
    if superscript:
        run.font.superscript = True


def _append_citation_hyperlink(paragraph, text: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")

    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for attr, value in (
        ("w:ascii", "Times New Roman"),
        ("w:hAnsi", "Times New Roman"),
        ("w:cs", "Times New Roman"),
        ("w:eastAsia", "\u5b8b\u4f53"),
    ):
        fonts.set(qn(attr), value)
    r_pr.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "18")
    r_pr.append(size)
    vert = OxmlElement("w:vertAlign")
    vert.set(qn("w:val"), "superscript")
    r_pr.append(vert)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    r_pr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "none")
    r_pr.append(underline)
    run.append(r_pr)

    node = OxmlElement("w:t")
    node.text = text
    run.append(node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def apply_reference_crosslinks(doc: Document) -> None:
    """Make body citations look like Word cross-references to the bibliography."""
    existing = _existing_bookmark_names(doc)
    next_id = _next_bookmark_id(doc)
    in_refs = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text == "\u53c2\u8003\u6587\u732e":
            in_refs = True
            continue
        if not in_refs:
            continue
        number = _reference_number(text)
        if number is None:
            continue
        name = f"ref_{number}"
        if name not in existing:
            _add_bookmark(paragraph, name, next_id)
            existing.add(name)
            next_id += 1

    in_refs = False
    for paragraph in doc.paragraphs:
        text = xml_text(paragraph._p) or paragraph.text
        if text.strip() == "\u53c2\u8003\u6587\u732e":
            in_refs = True
        text = re.sub(
            r"(对红绿色觉障碍用户非常不友好。)(?:\[(?:[1-9]\d?)(?:[-,，](?:[1-9]\d?))*\])+",
            r"\1",
            text,
        )
        if in_refs or not CITATION_RE.search(text):
            continue
        p = paragraph._p
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)
        pos = 0
        for match in CITATION_RE.finditer(text):
            _add_text_run(paragraph, text[pos:match.start()])
            first_number = re.split(r"[-,，]", match.group(1))[0]
            _append_citation_hyperlink(paragraph, match.group(0), f"ref_{first_number}")
            pos = match.end()
        _add_text_run(paragraph, text[pos:])


def set_paragraph_border(paragraph, *, top=False, bottom=False) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        p_pr.append(borders)
    for edge, enabled in (("top", top), ("bottom", bottom)):
        tag = qn("w:" + edge)
        element = borders.find(tag)
        if element is None:
            element = OxmlElement("w:" + edge)
            borders.append(element)
        if enabled:
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "6")
            element.set(qn("w:space"), "3")
            element.set(qn("w:color"), "000000")
        else:
            element.set(qn("w:val"), "nil")


def add_paragraph_after(paragraph, text: str, style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style is not None:
        new_para.style = style
    if text:
        new_para.add_run(text)
    return new_para


def make_raw_paragraph(text: str, style_id: str | None = None) -> OxmlElement:
    p = OxmlElement("w:p")
    if style_id:
        p_pr = OxmlElement("w:pPr")
        p_style = OxmlElement("w:pStyle")
        p_style.set(qn("w:val"), style_id)
        p_pr.append(p_style)
        p.append(p_pr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def remove_caption_and_previous_drawing(doc: Document, caption_prefix: str) -> None:
    for paragraph in list(body_paragraphs(doc)):
        probe = paragraph.text.strip().translate(HALFWIDTH_DIGITS)
        if probe.startswith(caption_prefix):
            paragraphs = body_paragraphs(doc)
            idx = next((i for i, p in enumerate(paragraphs) if p._p is paragraph._p), None)
            if idx is not None and idx > 0 and "w:drawing" in paragraphs[idx - 1]._p.xml:
                remove_paragraph(paragraphs[idx - 1])
            remove_paragraph(paragraph)
            return


def insert_evaluation_overview(doc: Document) -> None:
    image_path = ROOT / "figs" / "fig4_eval_overview.png"
    if not image_path.exists():
        return
    remove_caption_and_previous_drawing(doc, "图4 最终评估集")

    anchor = None
    for paragraph in body_paragraphs(doc):
        if paragraph.text.strip().startswith("Entropy 方面"):
            anchor = paragraph
            break
    if anchor is None:
        for paragraph in body_paragraphs(doc):
            if paragraph.text.strip().startswith("在 KSCgain 上"):
                anchor = paragraph
                break
    if anchor is None:
        return

    image_para = add_paragraph_after(anchor, "")
    image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_para.paragraph_format.first_line_indent = None
    image_para.paragraph_format.space_before = Pt(3)
    image_para.paragraph_format.space_after = Pt(3)
    image_para.add_run().add_picture(str(image_path), width=Cm(15.8))

    caption = add_paragraph_after(image_para, "图4 最终评估集 27 幅图像的 DE-Cividis 输出总览")
    style_paragraph(caption, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                    bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=1, after=0)
    disable_east_asian_auto_spacing(caption)


def insert_kodak_overview(doc: Document) -> None:
    image_path = ROOT / "figs" / "fig7_kodak_overview.png"
    if not image_path.exists():
        return
    remove_caption_and_previous_drawing(doc, "图7 Kodak")

    anchor = None
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if text.startswith("这组结果与主实验基本一致") or text.startswith("表2 表明"):
            anchor = paragraph
    if anchor is None:
        return

    image_para = add_paragraph_after(anchor, "")
    image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_para.paragraph_format.first_line_indent = None
    image_para.paragraph_format.space_before = Pt(3)
    image_para.paragraph_format.space_after = Pt(3)
    image_para.add_run().add_picture(str(image_path), width=Cm(15.8))

    caption = add_paragraph_after(image_para, "图7 Kodak 01-24 补充验证图像的 DE-Cividis 输出总览")
    style_paragraph(caption, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                    bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=1, after=0)
    disable_east_asian_auto_spacing(caption)


def ensure_section_headings(doc: Document) -> None:
    for paragraph in list(body_paragraphs(doc)):
        if paragraph.text.strip() == "5.2 典型图像对比":
            remove_paragraph(paragraph)
    for paragraph in body_paragraphs(doc):
        probe = paragraph.text.strip().translate(HALFWIDTH_DIGITS)
        if probe.startswith("图5 X"):
            paragraphs = body_paragraphs(doc)
            idx = next((i for i, p in enumerate(paragraphs) if p._p is paragraph._p), None)
            if idx is not None and idx > 0:
                target = paragraphs[idx - 1] if "w:drawing" in paragraphs[idx - 1]._p.xml else paragraph
                target._p.addprevious(make_raw_paragraph("5.2 典型图像对比", "Heading2"))
            return


def body_paragraphs(doc: Document):
    return list(doc.paragraphs)


def fix_figure_sequence(doc: Document) -> None:
    # Drop the weak old Figure 4: drawing paragraph followed by "图4 工业检测合成图上的多方法对比".
    paragraphs = body_paragraphs(doc)
    for i, para in enumerate(paragraphs):
        if para.text.strip().startswith("\u56fe4 \u5de5\u4e1a\u68c0\u6d4b\u5408\u6210\u56fe"):
            if i > 0 and "w:drawing" in paragraphs[i - 1]._p.xml:
                remove_paragraph(paragraphs[i - 1])
            remove_paragraph(para)
            break

    replacements = {
        "\u56fe 1 DE-Cividis \u65b9\u6cd5\u603b\u4f53\u6d41\u7a0b": "\u56fe1 DE-Cividis \u65b9\u6cd5\u603b\u4f53\u6d41\u7a0b",
        "\u56fe2 \u7684\u4f5c\u7528\u662f\u89e3\u91ca\u8272\u56fe\u9009\u62e9\u3002Jet \u5728\u6b63\u5e38\u89c6\u89c9\u4e0b\u989c\u8272\u5f88\u5f3a\uff0c\u4f46\u7ecf\u8fc7 protan \u548c deutan \u6a21\u62df\u540e\uff0c\u90e8\u5206\u989c\u8272\u533a\u57df\u4f1a\u53d8\u5f97\u63a5\u8fd1\uff0c\u5bb9\u6613\u5f71\u54cd\u7070\u5ea6\u5c42\u6b21\u5224\u65ad\uff1bCividis \u7684\u8272\u5f69\u53d8\u5316\u6ca1\u6709 Jet \u90a3\u4e48\u5f3a\uff0c\u4f46\u5728\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u66f4\u7a33\u5b9a\uff0c\u56e0\u6b64\u9002\u5408\u4f5c\u4e3a\u672c\u6587\u7684\u57fa\u7840\u8272\u56fe\u3002\u56fe3 \u53ea\u662f\u5e2e\u52a9\u8bf4\u660e KSCgain \u7684\u8ba1\u7b97\u4fa7\u91cd\uff1a\u5b83\u4e0d\u662f\u770b\u6574\u5e45\u56fe\u7684\u5e73\u5747\u53d8\u5316\uff0c\u800c\u662f\u66f4\u5173\u6ce8\u539f\u56fe\u4e2d\u8fb9\u7f18\u548c\u5c40\u90e8\u5bf9\u6bd4\u660e\u663e\u7684\u4f4d\u7f6e\u3002\u56fe4 \u4f7f\u7528\u5de5\u4e1a\u68c0\u6d4b\u5408\u6210\u56fe\u505a\u5c0f\u8303\u56f4\u5bf9\u6bd4\uff0c\u4e3b\u8981\u7528\u6765\u89c2\u5bdf\u4e0d\u540c\u65b9\u6cd5\u5728\u7ea2\u7eff\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u7684\u8272\u5f69\u7a33\u5b9a\u6027\u3002":
            "\u56fe2 \u7528\u540c\u4e00\u7070\u5ea6\u56fe\u5bf9\u6bd4 Jet \u548c Cividis\uff0c\u4e3b\u8981\u8bf4\u660e\u8272\u56fe\u9009\u62e9\u7684\u7406\u7531\u3002Jet \u5728\u6b63\u5e38\u89c6\u89c9\u4e0b\u989c\u8272\u5f88\u5f3a\uff0c\u4f46\u7ecf\u8fc7 protan \u548c deutan \u6a21\u62df\u540e\uff0c\u90e8\u5206\u989c\u8272\u533a\u57df\u53d8\u5f97\u63a5\u8fd1\uff0c\u5bb9\u6613\u5f71\u54cd\u7070\u5ea6\u5c42\u6b21\u5224\u65ad\uff1bCividis \u8272\u5f69\u4e0d\u90a3\u4e48\u5938\u5f20\uff0c\u4f46\u5728\u7ea2\u7eff\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u66f4\u7a33\u5b9a\uff0c\u56e0\u6b64\u9002\u5408\u4f5c\u4e3a\u672c\u6587\u7684\u57fa\u7840\u8272\u56fe\u3002\u56fe3 \u7528\u6765\u89e3\u91ca KSCgain \u8fd9\u4e2a\u8f85\u52a9\u6307\u6807\uff1a\u5b83\u4e0d\u770b\u6574\u5e45\u56fe\u7684\u5e73\u5747\u53d8\u5316\uff0c\u800c\u662f\u66f4\u5173\u6ce8\u539f\u56fe\u4e2d\u8fb9\u7f18\u548c\u5c40\u90e8\u5bf9\u6bd4\u660e\u663e\u7684\u4f4d\u7f6e\u3002",
        "\u56fe5 X \u5c04\u7ebf\u3001\u710a\u7f1d\u548c\u964d\u96e8\u56fe\u573a\u666f\u5bf9\u6bd4": "\u56fe4 X \u5c04\u7ebf\u3001\u710a\u7f1d\u548c\u964d\u96e8\u56fe\u573a\u666f\u5bf9\u6bd4",
        "\u56fe5 \u7ed9\u51fa\u4e86\u4e09\u4e2a\u5178\u578b\u573a\u666f\u3002X \u5c04\u7ebf\u56fe\u4e2d\uff0cJet \u867d\u7136\u989c\u8272\u5f3a\u70c8\uff0c\u4f46\u5728\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u6df1\u8272\u533a\u57df\u548c\u80cc\u666f\u5bb9\u6613\u6df7\u5728\u4e00\u8d77\uff1bCividis \u66f4\u7a33\u5b9a\uff0c\u4f46\u5c40\u90e8\u5f31\u8fb9\u754c\u4e0d\u591f\u9192\u76ee\u3002DE-Cividis \u5bf9\u7ec6\u8282\u5c42\u8fdb\u884c\u589e\u5f3a\u540e\uff0c\u77e9\u5f62\u8fb9\u754c\u548c\u6563\u70b9\u72b6\u7ec6\u8282\u66f4\u5bb9\u6613\u5206\u8fa8\uff0c\u540c\u65f6\u6574\u4f53\u660e\u6697\u5173\u7cfb\u4ecd\u63a5\u8fd1\u539f\u7070\u5ea6\u56fe\u3002":
            "\u56fe4 \u7ed9\u51fa\u4e86\u4e09\u4e2a\u5178\u578b\u573a\u666f\uff0c\u5b83\u6bd4\u5355\u5f20\u5de5\u4e1a\u68c0\u6d4b\u5408\u6210\u56fe\u66f4\u80fd\u770b\u51fa\u65b9\u6cd5\u5dee\u522b\u3002X \u5c04\u7ebf\u56fe\u4e2d\uff0cJet \u989c\u8272\u5f3a\u70c8\uff0c\u4f46\u5728\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u6df1\u8272\u533a\u57df\u548c\u80cc\u666f\u5bb9\u6613\u6df7\u5728\u4e00\u8d77\uff1bCividis \u66f4\u7a33\u5b9a\uff0c\u4f46\u5c40\u90e8\u5f31\u8fb9\u754c\u4e0d\u591f\u9192\u76ee\u3002DE-Cividis \u5bf9\u7ec6\u8282\u5c42\u8fdb\u884c\u589e\u5f3a\u540e\uff0c\u77e9\u5f62\u8fb9\u754c\u548c\u6563\u70b9\u72b6\u7ec6\u8282\u66f4\u5bb9\u6613\u5206\u8fa8\uff0c\u540c\u65f6\u6574\u4f53\u660e\u6697\u5173\u7cfb\u4ecd\u63a5\u8fd1\u539f\u7070\u5ea6\u56fe\u3002",
        "\u56fe6 DE-Cividis \u6d88\u878d\u5b9e\u9a8c\u53ef\u89c6\u5316\u7ed3\u679c": "\u56fe5 DE-Cividis \u6d88\u878d\u5b9e\u9a8c\u53ef\u89c6\u5316\u7ed3\u679c",
        "\u56fe7 \u590d\u6742\u7eb9\u7406\u548c\u566a\u58f0\u573a\u666f\u4e0b\u7684\u5c40\u9650\u6027\u5206\u6790": "\u56fe6 \u590d\u6742\u7eb9\u7406\u548c\u566a\u58f0\u573a\u666f\u4e0b\u7684\u5c40\u9650\u6027\u5206\u6790",
    }
    for para in body_paragraphs(doc):
        text = para.text.strip()
        if text in replacements:
            set_text(para, replacements[text])

    # Move "5.2 典型图像对比" before the typical-comparison drawing.
    paragraphs = body_paragraphs(doc)
    heading = next((p for p in paragraphs if p.text.strip() == "5.2 \u5178\u578b\u56fe\u50cf\u5bf9\u6bd4"), None)
    cap4 = next((p for p in paragraphs if p.text.strip().startswith("\u56fe4 X")), None)
    if heading is not None and cap4 is not None:
        paragraphs = body_paragraphs(doc)
        cap_index = next(i for i, p in enumerate(paragraphs) if p._p is cap4._p)
        if cap_index > 0 and "w:drawing" in paragraphs[cap_index - 1]._p.xml:
            drawing_p = paragraphs[cap_index - 1]
            for p in list(body_paragraphs(doc)):
                if p.text.strip() == "5.2 \u5178\u578b\u56fe\u50cf\u5bf9\u6bd4":
                    remove_paragraph(p)
            drawing_p._p.addprevious(make_raw_paragraph("5.2 \u5178\u578b\u56fe\u50cf\u5bf9\u6bd4", "Heading2"))


def apply_page_and_styles(doc: Document) -> None:
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.3)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.1)
    sec.right_margin = Cm(2.1)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "\u5b8b\u4f53")
    normal.font.size = Pt(10.5)

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if "w:drawing" in paragraph._p.xml:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(3)
            continue
        if not text:
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            continue

        if paragraph.style.name == "Heading 1":
            style_paragraph(paragraph, east="\u9ed1\u4f53", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                            first_line=False, before=8, after=4)
        elif paragraph.style.name == "Heading 2":
            style_paragraph(paragraph, east="\u9ed1\u4f53", size=12, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                            first_line=False, before=5, after=3)
        elif is_figure_caption(text):
            compact_caption_number(paragraph)
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                            align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=1, after=0)
            disable_east_asian_auto_spacing(paragraph)
        elif text.startswith("Fig."):
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                            italic=False, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=3)
        elif is_table_caption(text):
            compact_caption_number(paragraph)
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                            align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=3, after=2)
            disable_east_asian_auto_spacing(paragraph)
        elif paragraph in doc.paragraphs[:11]:
            # Front matter is handled separately below.
            pass
        else:
            normalize_body_reference_number(paragraph)
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                            first_line=True, before=0, after=0)

    # Front matter.
    for i in (0, 1):
        style_paragraph(doc.paragraphs[i], east="\u9ed1\u4f53", size=22, bold=True,
                        align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=0)
    style_paragraph(doc.paragraphs[2], east="\u5b8b\u4f53", size=14, bold=False,
                    align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=0)
    style_paragraph(doc.paragraphs[3], east="\u5b8b\u4f53", size=9, bold=False,
                    align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=6)
    style_paragraph(doc.paragraphs[4], east="\u6977\u4f53", size=10.5, bold=False,
                    align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=False, before=0, after=3)

    # Chinese keywords: label SimHei bold, content KaiTi.
    kw = doc.paragraphs[5]
    set_text(kw, "")
    r1 = kw.add_run("\u5173\u952e\u8bcd\uff1a")
    set_run_font(r1, east="\u9ed1\u4f53", size=10.5, bold=True)
    r2 = kw.add_run("\u8272\u89c9\u969c\u788d\uff1b\u4f2a\u5f69\u8272\uff1bCividis\uff1b\u7ec6\u8282\u589e\u5f3a\uff1b\u53cd\u9510\u5316\uff1b\u79d1\u5b66\u53ef\u89c6\u5316")
    set_run_font(r2, east="\u6977\u4f53", size=10.5, bold=False)
    kw.alignment = WD_ALIGN_PARAGRAPH.LEFT
    kw.paragraph_format.first_line_indent = None
    kw.paragraph_format.line_spacing = 1.25
    kw.paragraph_format.space_after = Pt(6)

    style_paragraph(doc.paragraphs[6], east="\u5b8b\u4f53", latin="Times New Roman", size=10.5, bold=True,
                    align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=3)
    style_paragraph(doc.paragraphs[7], east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                    align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=0)
    style_paragraph(doc.paragraphs[8], east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                    align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=3)
    style_paragraph(doc.paragraphs[9], east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                    align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=False, before=0, after=2)
    style_paragraph(doc.paragraphs[10], east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False,
                    align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=0, after=8)


def add_english_figure_captions(doc: Document) -> None:
    # Remove English figure captions that are not immediately after a real Chinese figure caption.
    for para in list(body_paragraphs(doc)):
        if para.text.strip().startswith("Fig."):
            remove_paragraph(para)

    for para in list(body_paragraphs(doc)):
        text = para.text.strip()
        if not re.match(r"^\u56fe\d+\s", text):
            continue
        key = text.split()[0]
        continue

    for para in body_paragraphs(doc):
        if para.text.strip().startswith("Fig."):
            style_paragraph(para, east="\u5b8b\u4f53", latin="Times New Roman", size=9, bold=False, italic=False,
                            align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=0, after=3)


def xml_text(element) -> str:
    return "".join(t.text for t in element.iter(qn("w:t")) if t.text).strip()


def finalize_figure_xml(doc: Document) -> None:
    body = doc._body._body

    # Remove all generated English figure captions and all misplaced 5.2 headings.
    for child in list(body.iterchildren()):
        if child.tag.split("}")[-1] != "p":
            continue
        text = xml_text(child)
        if text.startswith("Fig.") or text == "5.2 \u5178\u578b\u56fe\u50cf\u5bf9\u6bd4":
            body.remove(child)

    # Insert 5.2 before the typical-comparison figure drawing.
    children = list(body.iterchildren())
    for i, child in enumerate(children):
        if child.tag.split("}")[-1] != "p":
            continue
        if xml_text(child).startswith("\u56fe4 X"):
            drawing_index = i - 1
            if drawing_index >= 0:
                body.insert(drawing_index, make_raw_paragraph("5.2 \u5178\u578b\u56fe\u50cf\u5bf9\u6bd4", "Heading2"))
            break



def normalize_section_titles(doc: Document) -> None:
    replacements = {
        "2.3 图像细节增强方法": "2.3 反锐化与细节增强",
        "2.3 细节层增强方法": "2.3 反锐化与细节增强",
        "2.3 细节增强与反锐化方法": "2.3 反锐化与细节增强",
        "2.3 空间域细节增强与反锐化": "2.3 反锐化与细节增强",
        "3.3 方法特点": "3.3 候选方法比较",
        "3.3 方法设计取舍": "3.3 候选方法比较",
        "3.3 内容自适应 LUT 与细节增强路线比较": "3.3 候选方法比较",
        "3.3 候选方案比较与最终流程确定": "3.3 候选方法比较",
        "4.3 参数设置与实验流程": "4.3 参数设置",
        "5.1 最终评估集结果": "5.1 主评估集结果",
        "5.3 消融实验与补充验证": "5.3 消融与补充验证",
        "6.1 方法特点分析": "6.1 结果讨论",
        "6.2 指标解释与适用范围": "6.2 适用范围",
        "6.3 参数选择与实验复现": "6.3 实现说明",
        "6.3 参数与实现细节": "6.3 实现说明",
        "6.4 局限性分析": "6.4 局限性",
    }
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if text in replacements:
            set_text(paragraph, replacements[text])


def polish_wording(doc: Document) -> None:
    replacements = {
        "图4 X 射线、焊缝和降雨图场景对比": "图4 X 射线、焊缝和显微纹理图场景对比",
        "图4 给出了三个典型场景，它比单张工业检测合成图更能看出方法差别。X 射线图中，Jet 颜色强烈，但在色觉障碍模拟下深色区域和背景容易混在一起；Cividis 更稳定，但局部弱边界不够醒目。DE-Cividis 对细节层进行增强后，矩形边界和散点状细节更容易分辨，同时整体明暗关系仍接近原灰度图。":
            "图4 选取了三类灰度结构较清楚的图像，用来观察方法差异。X 射线图主要看矩形边界和内部暗块，焊缝图主要看细长裂纹和局部亮点，显微纹理图主要看纤维状边缘。Jet 的颜色最强，但在色觉障碍模拟下部分区域容易混淆；固定 Cividis 更稳定，但弱边界不够突出。DE-Cividis 在保持原灰度明暗关系的同时，让这些边缘和细节位置出现更明显的颜色差。",
        "焊缝图和降雨图也表现出类似现象。焊缝图更关注细长裂纹和局部过渡，降雨图更关注强度带和区域边界。DE-Cividis 的作用不是单纯提高饱和度，而是在结构变化较明显的位置增加颜色差异。":
            "该样例显示，DE-Cividis 主要在边缘、裂纹和纹理位置形成更明显的颜色差异。对于平坦区域，DE-Cividis 与 Cividis 的差异较小；对于局部变化明显的区域，增强后的灰度差会转化为更容易观察的色彩差异。",
        "前期实验中，本文尝试过内容自适应 LUT 方法。其思路是根据灰度级上的结构重要性构造一条单调变换曲线，把 Cividis 的更多颜色距离分配给重要灰度区间。这个想法与“关键结构优先颜色分配”是一致的，但实验中存在两个问题：一是增益不够明显，二是稳定性不如细节增强。":
            "在确定最终流程前，本文先比较了几种候选方案。其中内容自适应 LUT 的思路，是根据灰度级上的结构重要性调整 Cividis 的查表位置，让重要灰度区间获得更多颜色距离。这个方案逻辑上可行，但在实验中提升不够稳定。",
        "自适应 LUT 只根据灰度级重新分配颜色，如果同一灰度值同时出现在背景和结构区域，它无法区分二者。DE-Cividis 则直接在空间域增强局部细节，同样灰度值在不同邻域中可以得到不同的增强结果。因此，DE-Cividis 并不是放弃“关键结构优先”的思想，而是把实现方式从“改色图”换成“改输入灰度”。":
            "主要原因是，自适应 LUT 只改变灰度值到颜色的对应关系，仍然是一维映射。如果同一灰度值同时出现在背景和结构区域，它很难区分二者。DE-Cividis 改在空间域处理局部细节，同样的灰度值处在不同邻域时，增强结果可以不同，更符合本文关注边缘和纹理结构的目标。",
        "这种变化也使方法目标更加集中。前期尝试过的多模块方法步骤较多，参数之间容易互相影响。DE-Cividis 只保留细节增强、Cividis 映射和亮度回写三个环节，分别对应结构增强、色觉障碍友好上色和亮度结构保持，便于复现实验并分析每个环节的作用。":
            "最终采用细节增强、Cividis 映射和亮度回写作为统一实验流程，分别对应局部结构增强、色觉障碍友好上色和亮度结构保持。",
        "主要原因是，自适应 LUT 只改变灰度值到颜色的对应关系，仍然是一维映射。如果同一灰度值同时出现在背景和结构区域，它很难区分二者。DE-Cividis 改在空间域处理局部细节，同样的灰度值处在不同邻域时，增强结果可以不同，更符合本文想突出边缘和纹理的目标。":
            "主要原因是，自适应 LUT 只改变灰度值到颜色的对应关系，仍然是一维映射。如果同一灰度值同时出现在背景和结构区域，它很难区分二者。DE-Cividis 改在空间域处理局部细节，同样的灰度值处在不同邻域时，增强结果可以不同，更符合本文关注边缘和纹理结构的目标。",
        "因此，最终流程保留三个最容易解释和复现的环节：先做细节增强，再用 Cividis 上色，最后回写原灰度亮度。这样每一步都有对应目的，参数也少，比较适合作为课程实验中的主方法。":
            "最终采用细节增强、Cividis 映射和亮度回写作为统一实验流程，分别对应局部结构增强、色觉障碍友好上色和亮度结构保持。",
        "在确定最终方法前，我还尝试过几种思路，例如自适应 LUT、CLAHE-Cividis、Retinex-Cividis、双边滤波细节增强和多尺度反锐化等。这些方法有的能让局部对比更强，有的能降低部分色差，但整体稳定性不如单尺度 DE-Cividis。":
            "在确定最终方法前，本文还尝试过自适应 LUT、CLAHE-Cividis、Retinex-Cividis、双边滤波细节增强和多尺度反锐化等方案。这些方法有的能让局部对比更强，有的能降低部分色差，但综合效果不如单尺度 DE-Cividis 稳定。",
        "多尺度反锐化也有一定效果，但参数更多，调参过程更复杂。本文最终采用单尺度反锐化方案，是因为它在当前实验中已经能带来稳定提升，同时实现和复现都更直接。":
            "多尺度反锐化也有一定效果，但参数更多，调参过程更复杂。本文采用单尺度反锐化，是因为它在当前实验中已经能带来稳定提升，同时实现和复现都更直接。",
    }
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if text in replacements:
            set_text(paragraph, replacements[text])


def revise_current_draft_text(doc: Document) -> None:
    replacements = {
        "本文最初尝试过内容自适应色图路线，即根据每张图的灰度分布和结构重要性重新分配 Cividis 的查表位置。这条路线听起来更像“自适应方法”，但实验结果并不理想：它比固定 Cividis 有提升，却不如更直接的细节层增强稳定。经过多轮候选方法比较，本文最终采用 DE-Cividis，即 Detail-Enhanced Cividis。它不重新设计色图，而是在进入 Cividis 前增强图像细节层，再通过亮度回写保持原灰度结构。":
            "在前期实验中，本文重点比较了自适应 LUT、CLAHE-Cividis、融合基线和 DE-Cividis 等可复现实验方案。同时，Retinex-Cividis、双边滤波细节增强、Guided Filter 细节增强和多尺度反锐化也作为候选思路做过尝试，但没有纳入最终统计表。综合最终表格和典型图像后，本文采用 DE-Cividis，即 Detail-Enhanced Cividis。它不重新设计色图，而是在进入 Cividis 前增强灰度细节层，再通过亮度回写保持原灰度明暗结构。",
        "本文主要做了三件事。第一，以空间域锐化、伪彩色处理和彩色空间变换思想为基础，查阅资料后选用 Cividis 色图，并加入 YCbCr 亮度回写，组合成一个轻量流程，用来做红绿色觉障碍友好的灰度伪彩色显示。第二，设计了一个辅助评价量 KSCgain，用来观察原图边缘、纹理等关键区域在色觉障碍模拟后是否更容易区分。第三，在 27 幅最终评估图像和 16 幅额外 Kodak 图像上进行对比实验，并结合典型图像和消融结果分析方法效果。":
            "本文主要做了三件事。第一，组合反锐化细节增强、Cividis 色图和 YCbCr 亮度回写，形成一个面向红绿色觉障碍友好显示的灰度伪彩色流程。第二，使用 KSCgain 辅助观察边缘、纹理等关键区域在色觉障碍模拟后的可分辨性。第三，在 27 幅最终评估图像和 24 幅额外 Kodak 图像上进行对比，并用可视化图像和消融实验分析各步骤的作用。",
        "色觉障碍图像处理通常包括模拟和补偿两类任务。模拟的目标是估计色觉障碍用户看到的图像，例如 Brettel、Viénot 和 Mollon 提出的模型在 LMS 锥体响应空间中进行投影，能够模拟 protan、deutan 和 tritan 等情况[7]；Machado 等提出的生理模型也常用于色觉障碍模拟。补偿的目标则是重新调整颜色，让色觉障碍用户更容易区分原本混淆的颜色。":
            "色觉障碍图像处理通常包括模拟和补偿两类任务。模拟的目标是估计色觉障碍用户看到的图像，例如 Brettel、Viénot 和 Mollon 提出的模型在 LMS 锥体响应空间中进行投影，能够模拟 protan（红色盲）、deutan（绿色盲）和 tritan（蓝黄色觉障碍）等情况[7]；Machado 等提出的生理模型也常用于色觉障碍模拟。补偿的目标则是重新调整颜色，让色觉障碍用户更容易区分原本混淆的颜色。",
        "因此，最终流程保留三个容易检查的环节：先做细节增强，再用 Cividis 上色，最后回写原灰度亮度。这样每一步都有对应目的，参数也少，便于复现和对比。":
            "因此，最终流程分为三个环节：先做细节增强，再用 Cividis 上色，最后回写原灰度亮度。前两步负责增加结构区域的颜色差异，最后一步负责保持原图的明暗关系。",
        "本文主要从三个角度看结果：一是色差和结构相似度[16-17]，用来判断输出是否偏离原灰度图太多；二是图像边缘是否更明显，用 Sobel 梯度做一个简单近似；三是颜色信息是否更丰富，用信息熵作参考。这些指标不能完全代替人眼观察，所以本文还结合图像对比一起分析。":
            "本文使用五类指标评价结果。CIEDE2000 色差（ΔE2000）[16]用于观察输出图与原灰度结构之间的感知差异，数值越小表示改动越保守；Y-SSIM（亮度通道结构相似度）和 RGB-SSIM（RGB 三通道结构相似度）[17]用于观察结构保持情况；EdgeGain（边缘增益）用 Sobel 梯度比较色觉障碍模拟后的边缘强度；Entropy（信息熵）用于描述输出颜色变化的丰富程度。",
        "为了更贴近本文关心的边缘和纹理区域，实验中还加入了一个自定义辅助量 KSCgain。它的想法很简单：先找出原灰度图里边缘强、局部对比明显的位置，再看这些位置在色觉障碍模拟后是否仍有较明显的颜色或亮度差异。KSCgain 大于 1 时，表示它比固定 Cividis 在这些位置更容易拉开差异。":
            "除上述指标外，本文还计算 KSCgain（关键结构对比增益）。该指标先根据原灰度图中的 Sobel 梯度和局部对比度得到结构权重，再比较色觉障碍模拟后 Lab 空间中的多通道梯度。KSCgain 大于 1 表示该方法在边缘、纹理和局部突变区域的可见对比高于固定 Cividis。",
        "KSCgain 不是用户实验，也不能直接代表真实色觉障碍用户的主观体验。它和本文的细节增强思路都用到了边缘和局部对比，因此只适合作为辅助观察指标，最终判断仍需结合色差、结构相似度和可视化结果。":
            "该指标只作为辅助观察量。后文分析结果时，仍需要结合 ΔE2000、SSIM、EdgeGain、Entropy 和图像效果一起判断。",
        "图2 主要用于说明本文为什么选择 Cividis 作为基础色图。Jet 在正常视觉下颜色差异明显，但在 protan 和 deutan 模拟后，部分颜色区域会变得接近，容易造成灰度层次混淆。Cividis 的颜色变化更克制，在两类红绿色觉障碍模拟下仍能保持较稳定的明暗过渡，因此更适合作为后续细节增强的颜色基底。":
            "图2 对比了 Jet 和 Cividis 在正常视觉、protan（红色觉缺陷）和 deutan（绿色觉缺陷）模拟下的显示效果。Jet 的颜色更鲜艳，但模拟后部分颜色区域容易接近；Cividis 的色彩变化较平缓，明暗过渡也更稳定。基于这一观察，本文后续实验以 Cividis 作为基础色图。",
        "图3 不是新的算法模块，而是为了说明 KSCgain 的关注区域。Sobel 梯度突出边缘，局部对比度突出灰度变化集中的区域，二者合成后得到的权重 W 用来让评价更关注边缘、纹理和局部突变。这样做可以避免只看整幅图平均变化，但它仍然只是辅助指标，不能替代真实用户实验。":
            "图3 展示了 KSCgain 中结构权重的来源。Sobel 梯度主要响应边缘，局部对比度主要响应灰度变化集中的区域，二者合成后的权重 W 会突出边缘、纹理和局部突变位置。后续计算 KSCgain 时，这些位置在平均值中占更大权重。",
        "表1 显示，固定 Cividis 的色差最低，说明它最保守；但它的 KSCgain 被定义为 1.00，并没有主动增强关键结构。DE-Cividis 的色差接近 Cividis 和 CLAHE-Cividis，远低于 Jet、Viridis 和多模块融合基线。这说明 DE-Cividis 没有为了增强结构而大幅偏离原图。":
            "表1 汇总了 27 幅最终评估图像上的平均结果。固定 Cividis 的色差最低，说明它对原灰度变化最保守；DE-Cividis 的色差为 15.16，与固定 Cividis、CLAHE-Cividis 和自适应 LUT 接近，明显低于 Jet、Hot、Viridis 和多模块融合基线。",
        "在 Y-SSIM 上，DE-Cividis 与自适应 LUT 和 CLAHE-Cividis 基本相当。这个结果与亮度回写有关：不管细节增强怎样影响颜色，最终输出的亮度通道仍然由原灰度图约束。因此，本文同时列出 RGB-SSIM，避免只用 Y-SSIM 造成过度乐观的解释。":
            "在结构相似度方面，DE-Cividis 的 Y-SSIM 为 0.986，与自适应 LUT 和 CLAHE-Cividis 基本一致。其 RGB-SSIM 低于固定 Cividis 和自适应 LUT，说明方法主要通过色彩差异提示结构，而不是完全保持三个颜色通道都接近灰度图。",
        "在 KSCgain 上，DE-Cividis 高于 Cividis、自适应 LUT 和 CLAHE-Cividis，说明细节增强确实把更多色彩差异分配给了关键结构区域。但 Jet 和显著性重着色在部分指标上也可能较高，因此不能简单说 DE-Cividis 在所有指标上最好。本文更关注的是在色差和结构保真仍可接受的前提下，获得更稳定的关键结构提示。":
            "在 KSCgain 上，DE-Cividis 的平均值为 1.60，高于固定 Cividis 的 1.00、自适应 LUT 的 1.25 和 CLAHE-Cividis 的 1.45。这说明细节增强后，边缘和纹理区域在色觉障碍模拟下更容易形成可见差异。",
        "EdgeGain 的结果需要谨慎解释。DE-Cividis 的平均 EdgeGain 接近 Jet；Daltonization 的 EdgeGain 更高，但其 Y-SSIM、RGB-SSIM 和色差表现较差，并且它处理的是已经着色的图像。换句话说，单看 EdgeGain 会误导结论。":
            "EdgeGain 方面，DE-Cividis 与固定 Cividis 和 Jet 接近，说明它没有单纯依靠提高亮度边缘来获得效果。Daltonization 和显著性重着色的部分指标较高，但色差也明显增大，图像整体变化更强。",
        "Entropy 方面，DE-Cividis 高于 Cividis、多模块融合基线和自适应 LUT，说明输出包含更丰富的颜色变化。颜色信息量增加并不必然是好事，因为过高的熵也可能来自噪声或过度增强；因此需要结合色差、Y-SSIM、RGB-SSIM 和可视化结果综合判断。":
            "Entropy 方面，DE-Cividis 高于 Cividis、多模块融合基线和自适应 LUT，说明输出包含更丰富的颜色变化。结合色差、Y-SSIM 和 KSCgain 可以看出，DE-Cividis 的主要特点是在较小色差增加下提升关键结构区域的可见差异。",
        "图4 X 射线、焊缝和显微纹理图场景对比":
            "图5 X 射线、焊缝和地图场景对比",
        "图4 选取了三类灰度结构较清楚的图像，用来观察方法差异。X 射线图主要看矩形边界和内部暗块，焊缝图主要看细长裂纹和局部亮点，显微纹理图主要看纤维状边缘。Jet 的颜色最强，但在色觉障碍模拟下部分区域容易混淆；固定 Cividis 更稳定，但弱边界不够突出。DE-Cividis 在保持原灰度明暗关系的同时，让这些边缘和细节位置出现更明显的颜色差。":
            "图5 选取了 KSCgain 提升较明显的三类样例。X 射线图中，DE-Cividis 对矩形边界和内部暗块的区分更清楚；焊缝图中，细长裂纹和局部亮点比固定 Cividis 更醒目；纹理图像中，毛发、褶皱等纹理边界的颜色差异也更明显。CLAHE-Cividis 在局部对比上也有明显改善，但部分背景纹理会被一起拉强。相比之下，DE-Cividis 的颜色变化更克制，整体明暗关系仍接近原灰度图。",
        "这张图放在正文中的作用，是说明 DE-Cividis 的优势主要出现在边缘、裂纹和纹理等结构位置，而不是让整幅图变得更花。对于平坦区域，DE-Cividis 与 Cividis 的差异较小；对于局部变化明显的区域，增强后的灰度差会转化为更容易观察的色彩差异。":
            "这些现象与方法设计是一致的：高斯背景层去除后，裂纹、边界和细纹理会进入细节层；增强后的灰度差再经过 Cividis 映射，最终表现为局部结构处更明显的颜色差异。",
        "图5 DE-Cividis 消融实验可视化结果":
            "图6 DE-Cividis 消融实验可视化结果",
        "在确定最终方法前，本文还尝试过自适应 LUT、CLAHE-Cividis、Retinex-Cividis、双边滤波细节增强和多尺度反锐化等方案。这些方法有的能让局部对比更强，有的能降低部分色差，但综合效果不如单尺度 DE-Cividis 稳定。":
            "图6 展示了 DE-Cividis 的消融结果。去掉细节增强后，边界位置的颜色差异变弱；去掉亮度回写后，局部对比虽然更强，但明暗关系更容易偏离原灰度图；改用 Jet 后颜色变化更鲜艳，但在色觉障碍模拟下不够稳定。",
        "其中 CLAHE-Cividis 是最值得比较的候选。它通过局部直方图均衡增强对比度，很多图像看起来更清楚，但有时也会把小区域的噪声或纹理拉得过强。DE-Cividis 的增强方式更简单，只增强相对于高斯背景的细节层，因此结果相对克制。":
            "因此，本文保留“细节增强 + Cividis + 亮度回写”三个步骤。细节增强负责拉开结构区域，Cividis 提供相对色觉障碍友好的颜色基底，亮度回写用于保持原始灰度的明暗层次。",
        "多尺度反锐化也有一定效果，但参数更多，调参过程更复杂。本文采用单尺度反锐化，是因为它在当前实验中已经能带来稳定提升，同时实现和复现都更直接。":
            "在最终统计的候选方法中，CLAHE-Cividis 的局部对比增强最明显，自适应 LUT 的色差和结构相似度较接近 DE-Cividis，融合基线则体现了多模块组合的效果。综合表1、表2和典型图像，单尺度 DE-Cividis 在结构提示和亮度保真之间更均衡。",
        "额外 Kodak 09-24 共 16 幅图像用于补充验证。这部分不再比较所有候选方法，只观察 Cividis、CLAHE-Cividis 和 DE-Cividis 三种关键方法在固定参数下是否仍然稳定。":
            "额外 Kodak 09-24 共 16 幅图像用于补充验证。这部分只比较 Cividis、CLAHE-Cividis 和 DE-Cividis 三种方法，用来观察固定参数在另一组自然图像上的表现。",
        "表2 表明，在额外 Kodak 图像上，DE-Cividis 的色差低于 CLAHE-Cividis，Y-SSIM 略高于 CLAHE-Cividis，KSCgain 与 CLAHE-Cividis 接近。Cividis 的色差和 RGB-SSIM 仍然更保守，但 KSCgain 为 1.00，说明它没有提供额外结构增强。这个结果说明 DE-Cividis 在固定参数下仍有一定稳定性，但结论仍限定在本文测试的图像范围内。":
            "表2 表明，在额外 Kodak 图像上，DE-Cividis 的色差低于 CLAHE-Cividis，Y-SSIM 略高于 CLAHE-Cividis，KSCgain 也略高于 CLAHE-Cividis。固定 Cividis 的色差和 RGB-SSIM 更保守，但 KSCgain 为 1.00，说明它没有额外增强关键结构。",
        "留出集的 EdgeGain 没有列入表2，是因为自然图像灰度结构和科学可视化图不同，EdgeGain 对图像内容非常敏感。本文把额外 Kodak 图像作为补充观察，而不是重新调参依据。所有参数在留出验证前已经固定为 σ=2.0、α=2.0。":
            "这组结果与主实验基本一致：DE-Cividis 相比固定 Cividis 提供了更高的关键结构对比，相比 CLAHE-Cividis 则保持了略低的色差。所有参数在补充验证前已经固定为 σ=2.0、α=2.0。",
        "6.2 指标解释与适用范围":
            "6.2 适用场景",
        "多个指标需要一起解释。色差低说明输出保守，但不一定更醒目；边缘增益高说明边缘更强，但也可能来自噪声或伪边界。因此，本文把色差、结构相似度、边缘变化、KSCgain 和图像对比放在一起看，不把某一个指标作为唯一目标。":
            "从实验现象看，DE-Cividis 更适合单通道强度图像的辅助显示。对于 X 射线、焊缝、降雨、热成像、显微结构和地形高度图等输入，原始数据通常具有明确的灰度强度含义，伪彩色的作用是帮助观察强度变化和局部结构。",
        "KSCgain 提供的是一个辅助观察角度。它先找原图中边缘和局部对比明显的位置，再看这些位置经过色觉障碍模拟后是否还容易分辨。由于这个指标与细节增强存在一定同源性，解释时需要结合色差、SSIM 和图像对比一起判断。":
            "如果输入本身是普通 RGB 彩色照片，直接套用本文流程并不合适，因为彩色照片中的颜色往往已经带有语义信息。此时更适合采用面向彩色图像的重着色或 Daltonization 方法。",
        "DE-Cividis 更适合灰度科学可视化，而不是普通彩色照片的重着色。对于 X 射线、焊缝、降雨、热成像、显微结构和地形高度图等输入，原始数据通常具有明确的单通道强度含义，伪彩色的目标是辅助观察强度变化。对于已经具有语义颜色的 RGB 图像，应优先采用面向彩色图像的 Daltonization 或重着色方法，而不应简单转灰度后再套用本文流程。":
            "",
        "本文没有使用深度学习模型。这样做并不是否定学习式方法，而是出于任务边界和可复现性的考虑。灰度伪彩色映射本身可以由经典滤波、查表映射和颜色空间变换完成；在没有大规模标注数据和真实色觉障碍用户反馈的情况下，使用轻量可解释方法更容易控制变量，也更便于分析每个模块对结果的影响。":
            "",
        "本文实验代码按功能划分为数据读取、伪彩色方法、色觉障碍模拟、评价指标和最终实验几个部分。经典色图、CLAHE-Cividis 以及 DE-Cividis 均以独立函数实现，避免不同方法之间共享中间状态。实验结果由程序统计后整理成表，减少人工誊写造成的误差。":
            "本文实验代码按功能划分为数据读取、伪彩色方法、色觉障碍模拟、评价指标和最终统计几个部分。经典色图、CLAHE-Cividis 和 DE-Cividis 均以独立函数实现，所有表格结果由程序重新计算后汇总。",
        "参数固定后，程序使用同一组 σ = 2.0、α = 2.0 处理所有主实验图像和额外 Kodak 图像。这样做的目的是让结果可重复，也避免某些图像因为单独调参而看起来过好。因此，表格中的指标和正文的典型图像展示都来自同一个固定流程。":
            "参数固定后，程序使用同一组 σ = 2.0、α = 2.0 处理所有最终评估图像和额外 Kodak 图像。表格中的指标和正文中的图像展示都来自这组固定参数。",
        "图6 复杂纹理和噪声场景下的局限性分析":
            "图7 复杂纹理和噪声场景下的局限性分析",
        "第一，本文使用色觉障碍模拟和客观指标进行评价，没有招募真实色觉障碍用户完成任务实验。KSCgain 能描述关键结构区域的可见对比，但它仍然是模型化近似。后续工作可以设计边界识别、异常定位或强度排序任务，让真实用户参与评价，从而检验客观指标与主观体验之间的一致性。":
            "对于噪声较重的图像，可以先进行轻度去噪，或者适当减小细节增强系数 α。对于灰度变化较弱的图像，DE-Cividis 的提升也会相应变小，因为可增强的细节层本身有限。",
        "第二，反锐化会增强高频成分，因此在噪声很重的图像中可能同时放大噪声。本文采用固定参数是为了保持方法简洁和对比公平，但在实际医学或工业检测场景中，可以加入噪声估计模块，使 α 随局部噪声水平自适应变化，或者在细节增强前加入边缘保持去噪。":
            "本文的评价主要基于色觉障碍模拟和客观指标。后续可设计边界识别、异常定位或强度排序等任务，并引入真实用户参与测试，从而进一步验证客观指标和实际观察体验之间的关系。",
        "本文的评价主要基于色觉障碍模拟和客观指标。后续如果继续完善，可以设计边界识别、异常定位或强度排序等任务，请真实用户参与测试，从而进一步验证客观指标和实际观察体验之间的关系。":
            "本文的评价主要基于色觉障碍模拟和客观指标。后续可设计边界识别、异常定位或强度排序等任务，并引入真实用户参与测试，从而进一步验证客观指标和实际观察体验之间的关系。",
        "第三，本文主要讨论 protan 和 deutan 两类红绿色觉障碍，对 tritan 蓝黄色觉障碍涉及较少。原因是红绿色觉障碍更常见，且 Cividis 的设计重点也在红绿混淆条件下的可读性。未来可以扩展到不同色觉障碍类型和不同严重程度，并研究个性化参数设置。":
            "",
        "第四，本文采用的测试图像规模仍然有限。虽然最终评估图像覆盖自然图、地图、Ishihara 合成图和若干科学可视化场景，额外 Kodak 图像也提供了补充验证，但更严格的评估仍需要更大规模、分类型的数据集。未来可以按医学影像、工业检测、遥感可视化和材料显微图像分别建立测试子集。":
            "",
        "第五，DE-Cividis 当前没有对图像语义进行建模。它强调的是边缘、纹理和局部强度变化，而不是识别病灶、裂纹或目标类别。因此，该方法适合作为可视化增强工具，而不能替代专业检测算法。若要面向具体应用，还需要与领域任务指标结合评价。":
            "",
    }
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text == "红绿色觉障碍友好灰度图像伪彩色方法":
            set_text_hard(paragraph, "红绿色觉障碍友好的灰度图像伪彩色方法")
            continue
        if text.startswith("摘要：灰度图像伪彩色处理常用于"):
            set_text_hard(
                paragraph,
                "摘要：灰度图像伪彩色处理常用于医学影像、工业检测、地理可视化等场景，但以 Jet 为代表的部分彩虹色图依赖明显颜色差异，"
                "Hot 等色图也可能存在高亮饱和或颜色信息量不足的问题，色觉障碍用户容易丢失关键信息。固定 Cividis 色图具有较好的色觉障碍友好性，"
                "但它只按灰度值进行全局映射，不能主动增强边缘、裂缝、纹理等关键结构。针对这一问题，本文提出 DE-Cividis 方法：先用高斯平滑获得背景层，"
                "再以反锐化方式增强细节层，将增强后的灰度送入 Cividis 色图，最后在 YCbCr 空间把亮度通道回写为原灰度图。该方法流程较轻量，参数较少，便于解释。"
                "27 幅最终评估图像和 24 幅额外 Kodak 图像上的实验表明，DE-Cividis 在保持较高亮度结构相似度的同时，有助于提升色觉障碍模拟视角下的关键结构对比。"
                "实验基于模拟模型和客观指标，后续仍可结合真实用户观察进一步验证。",
            )
            continue
        if text.startswith("Abstract: Pseudocolor mapping helps reveal structures"):
            set_text_hard(
                paragraph,
                "Abstract: Pseudocolor mapping helps reveal structures in grayscale images, but rainbow colormaps such as Jet rely on strong color transitions, "
                "and Hot may suffer from saturation or limited color variation. This paper proposes DE-Cividis, a lightweight method that enhances a detail layer, "
                "maps the result with Cividis, and writes back the original luminance in YCbCr space. Experiments on 27 evaluation images and 24 additional Kodak images "
                "show that DE-Cividis improves key-structure contrast under simulated color vision deficiency while maintaining luminance structure. The evaluation is based "
                "on simulation models and objective metrics, and can be further checked with real-user observation.",
            )
            continue
        if text.startswith("伪彩色处理并不是简单地让图像更鲜艳"):
            set_text_hard(
                paragraph,
                "伪彩色处理的目的不是单纯增强图像鲜艳程度，而是借助人眼对颜色差异的敏感性，把原本难以分辨的强度变化表达出来。"
                "一个好的伪彩色方法应该让观察者更快发现结构，而不是制造不存在的结构。许多常用色图并不满足这个要求。"
                "以 Jet 为代表的彩虹色图颜色鲜艳，但亮度变化不单调，容易形成伪边界[3]；同时它大量使用红绿方向的差异，对红绿色觉障碍用户非常不友好。",
            )
            continue
        if text.startswith("在前期实验中，本文比较了自适应 LUT、CLAHE-Cividis、Retinex-Cividis"):
            set_text_hard(
                paragraph,
                "在前期实验中，本文重点比较了自适应 LUT、CLAHE-Cividis、融合基线和 DE-Cividis 等可复现实验方案。"
                "同时，Retinex-Cividis、双边滤波细节增强、Guided Filter 细节增强和多尺度反锐化也作为候选思路做过尝试，"
                "但没有纳入最终统计表。综合最终表格和典型图像后，本文采用 DE-Cividis，即 Detail-Enhanced Cividis。"
                "它不重新设计色图，而是在进入 Cividis 前增强灰度细节层，再通过亮度回写保持原灰度明暗结构。",
            )
            continue
        if text.startswith("双边滤波和 Guided Filter 都可以得到边缘保持的背景层"):
            set_text_hard(
                paragraph,
                "双边滤波和 Guided Filter 都可以得到边缘保持的背景层[12-13]。理论上，它们比高斯滤波更能避免跨边缘模糊，适合分离纹理和结构。"
                "不过这类方法会增加窗口半径、正则化系数等设置。为了保持流程简洁，DE-Cividis 采用参数更少的高斯背景层估计。",
            )
            continue
        if text.startswith("多尺度反锐化进一步把不同尺度的细节层加权融合"):
            set_text_hard(
                paragraph,
                "多尺度反锐化可以把不同尺度的细节层加权融合，用来同时增强细边缘和中尺度纹理。"
                "本文采用单尺度反锐化作为实现，主要考虑参数数量、结果稳定性和方法解释的清晰度。",
            )
            continue
        if text.startswith("本文主要做了三件事。第一"):
            set_text(
                paragraph,
                "本文主要做了三件事。第一，组合反锐化细节增强、Cividis 色图和 YCbCr 亮度回写，形成一个面向红绿色觉障碍友好显示的灰度伪彩色流程。"
                "第二，使用 KSCgain 辅助观察边缘、纹理等关键区域在色觉障碍模拟后的可分辨性。第三，在 27 幅主评估图像和 24 幅 Kodak 补充图像上进行对比，"
                "并用可视化图像和消融实验分析各步骤的作用。",
            )
            continue
        if text.startswith("This paper") and "16 additional Kodak" in text:
            set_text(paragraph, text.replace("16 additional Kodak images", "24 additional Kodak images"))
            continue
        if text.startswith("这里的“关键结构”"):
            set_text(
                paragraph,
                "本文所说的关键结构，指图像处理意义上的边缘、纹理、局部突变和细节层，不对应语义识别中的目标类别。"
                "对于 X 射线图像，它可能表现为缺陷边界；对于降雨图，它可能表现为强度变化带；对于焊缝图，它可能表现为裂纹或过渡区域。"
                "本文不使用训练数据，而是用局部频率和局部对比信息近似这种结构重要性。",
            )
            continue
        if text.startswith("在确定最终流程前，本文先比较了几种候选方案"):
            set_text(
                paragraph,
                "在确定最终流程前，本文比较了几种候选方案。融合基线把早期的显著性、色觉补偿和亮度约束放在同一流程中，模块较多但色差偏大；"
                "自适应 LUT 根据灰度级上的结构重要性调整 Cividis 的查表位置，提升不够稳定；CLAHE-Cividis 通过局部直方图均衡增强输入，视觉上较明显，"
                "但有时会把背景纹理也一起拉强。表格中的候选方法主要用于说明这些不同路线与最终 DE-Cividis 的差别。",
            )
            continue
        if text.startswith("在确定最终流程前，本文比较了几种候选方案"):
            set_text(
                paragraph,
                "在确定最终流程前，本文比较了几种候选方案。融合基线把早期的显著性、色觉补偿和亮度约束放在同一流程中，模块较多但色差偏大；"
                "自适应 LUT 根据灰度级上的结构重要性调整 Cividis 的查表位置，提升不够稳定；CLAHE-Cividis 通过局部直方图均衡增强输入，视觉上较明显，"
                "但有时会把背景纹理也一起拉强。表格中的候选方法主要用于说明这些不同路线与最终 DE-Cividis 的差别。",
            )
            continue
        if text.startswith("主实验重点比较 Jet、Hot、Viridis、Cividis、CLAHE-Cividis 和 DE-Cividis"):
            set_text(
                paragraph,
                "主实验重点比较 Jet、Hot、Viridis、Cividis、CLAHE-Cividis 和 DE-Cividis；同时保留 Daltonization、显著性重着色、融合基线和自适应 LUT 作为参考。"
                "这些方法分别代表直接色图、已有彩色图补偿、多模块组合、一维自适应查表和局部均衡增强等思路，便于观察不同处理路线在同一批灰度输入上的差异。",
            )
            continue
        if text.startswith("除上述指标外，本文还计算 KSCgain"):
            set_text(
                paragraph,
                "除上述指标外，本文还计算 KSCgain（关键结构对比增益）。该指标先根据原灰度图中的 Sobel 梯度和局部对比度得到结构权重，"
                "再比较色觉障碍模拟后 Lab 空间中的多通道梯度。KSCgain 大于 1 表示该方法在边缘、纹理和局部突变区域的可见对比高于固定 Cividis。",
            )
            continue
        if text.startswith("KSCgain 的作用是"):
            set_text(
                paragraph,
                "KSCgain 用于补充观察结构区域的变化。后文分析结果时，仍结合 ΔE2000、SSIM、EdgeGain、Entropy 和图像效果一起判断。",
            )
            continue
        if text.startswith("这种划分不是机器学习意义上的训练/测试流程"):
            set_text(
                paragraph,
                "为了避免不同图像单独调参带来的偏差，候选观察完成后，后续统计都使用固定的方法列表和固定参数。"
                "所有主要方法都在同一组输入图像上运行，输入预处理也保持一致，因此横向比较具有可比性。",
            )
            continue
        if text.startswith("图5 选取了 KSCgain 提升较明显"):
            set_text(
                paragraph,
                "图5 选取了 KSCgain 提升较明显的三类样例。X 射线图中，DE-Cividis 对矩形边界和内部暗块的区分更清楚；"
                "焊缝图中，细长裂纹和局部亮点比固定 Cividis 更醒目；纹理图像中，毛发、褶皱等纹理边界的颜色差异也更明显。"
                "CLAHE-Cividis 在局部对比上也有明显改善，但部分背景纹理会被一起拉强。相比之下，DE-Cividis 的颜色变化更克制，整体明暗关系仍接近原灰度图。",
            )
            continue
        if text.startswith("因此，本文保留“细节增强 + Cividis + 亮度回写”"):
            set_text(
                paragraph,
                "因此，本文保留细节增强、Cividis 和亮度回写三个步骤。细节增强负责拉开结构区域，Cividis 提供相对色觉障碍友好的颜色基底，亮度回写用于保持原始灰度的明暗层次。",
            )
            continue
        if text.startswith("其他候选方法也有可取之处"):
            set_text(
                paragraph,
                "在最终统计的候选方法中，CLAHE-Cividis 的局部对比增强最明显，自适应 LUT 的色差和结构相似度较接近 DE-Cividis，融合基线则体现了多模块组合的效果。"
                "综合表1、表2和典型图像，单尺度 DE-Cividis 在结构提示和亮度保真之间更均衡。",
            )
            continue
        if text.startswith("与多模块融合基线相比"):
            set_text(
                paragraph,
                "与多模块融合基线相比，DE-Cividis 的处理链更短。它先增强细节，再用 Cividis 上色，最后把原灰度亮度写回去。三个步骤分别对应结构、颜色和亮度三个问题，模块之间的关系比较清楚，代码实现时也更容易检查。",
            )
            continue
        if text.startswith("从指标上看，DE-Cividis 并没有追求"):
            set_text(
                paragraph,
                "从指标上看，固定 Cividis 的 ΔE 最低，说明它与原灰度亮度变化最接近，但它的 KSCgain 被定义为 1，缺少额外结构增强能力。"
                "Jet 和部分补偿方法在边缘增益上可能较高，但其 Y-SSIM、RGB-SSIM 和 ΔE 较差，说明这类方法容易引入伪边界或破坏原始结构。"
                "DE-Cividis 的结果位于两者之间：它增加少量色差，换取更高的关键结构对比增益，同时保持较高的亮度结构相似度。",
            )
            continue
        if text.startswith("从指标上看，固定 Cividis 的 ΔE 最低"):
            set_text(
                paragraph,
                "从指标上看，固定 Cividis 的 ΔE 最低，说明它与原灰度亮度变化最接近，但它的 KSCgain 被定义为 1，缺少额外结构增强能力。"
                "Jet 和部分补偿方法在边缘增益上可能较高，但其 Y-SSIM、RGB-SSIM 和 ΔE 较差，说明这类方法容易引入伪边界或破坏原始结构。"
                "DE-Cividis 的结果位于两者之间：它增加少量色差，换取更高的关键结构对比增益，同时保持较高的亮度结构相似度。",
            )
            continue
        if text.startswith("DE-Cividis 的实现只依赖常见图像处理库"):
            set_text(
                paragraph,
                "DE-Cividis 的实现使用常见图像处理库完成。高斯滤波由 OpenCV 实现，Cividis 色图由 Matplotlib 提供，Lab 色彩空间、SSIM 和色差计算由 scikit-image 提供。"
                "对大小为 H×W 的图像，主要计算量来自卷积、查表和颜色空间转换，整体复杂度近似为 O(HW)。",
            )
            continue
        if text.startswith("本文围绕红绿色觉障碍用户"):
            set_text(
                paragraph,
                "本文围绕红绿色觉障碍用户的灰度图像伪彩色可读性问题，提出了 DE-Cividis 方法。该方法通过高斯背景层估计和反锐化细节增强，"
                "使边缘、纹理和局部突变区域在进入色图前获得更大的灰度差；随后使用 Cividis 提供相对色觉障碍友好的颜色基底；"
                "最后在 YCbCr 空间回写原灰度亮度，以保持原始结构。整体流程较简单，参数较少，便于解释。",
            )
            continue
        if text.startswith("实验结果表明，DE-Cividis 相比固定 Cividis"):
            set_text(
                paragraph,
                "实验结果表明，DE-Cividis 相比固定 Cividis 能提高 KSCgain；相比多模块融合基线，它具有更低色差和更高亮度结构一致性；"
                "相比自适应 LUT，它的色差和 Y-SSIM 接近，但 KSCgain 更高；相比 CLAHE-Cividis，它在最终评估图像和额外 Kodak 图像上表现较均衡。"
                "在本文测试的图像范围内，DE-Cividis 更适合作为结构保真要求较高的灰度科学可视化场景中的轻量伪彩色方案。",
            )
            continue
        if text.startswith("除最终评估图像外，本文额外使用 Kodak"):
            remove_paragraph(paragraph)
            continue
        if text.startswith("实验共使用 51 幅正式实验图像"):
            set_text(
                paragraph,
                "实验共使用 51 幅图像。主评估集包含 27 幅图像，来源包括 USC-SIPI、地图类图像、Ishihara 合成图以及若干灰度科学可视化合成样例；"
                "补充验证集包含 Kodak 01-24 共 24 幅图像。主评估集用于完整横向比较，Kodak 图像用于观察固定参数在另一组自然图像上的表现。",
            )
            continue
        if text.startswith("实验共准备 35 幅候选图像"):
            set_text(
                paragraph,
                "实验共使用 51 幅正式实验图像。主评估集包含 27 幅图像，来源包括 USC-SIPI、地图类图像、Ishihara 合成图以及若干灰度科学可视化合成样例；"
                "补充验证集包含 Kodak 01-24 共 24 幅图像。主评估集用于完整横向比较，Kodak 图像用于观察固定参数在另一组自然图像上的表现。",
            )
            continue
        if text.startswith("实验过程分为两步。第一步先用几张图观察"):
            set_text(
                paragraph,
                "实验过程分为两步。第一步先用少量代表图像观察不同参数和候选方法的效果，确定大致方向；第二步固定参数和方法列表，在 27 幅主评估图像上重新计算指标，再用 24 幅 Kodak 图像做补充观察。",
            )
            continue
        if (
            text.startswith("色觉障碍图像处理通常包括模拟和补偿两类任务")
            and "Brettel" in text
            and "Machado" in text
        ):
            set_text_hard(
                paragraph,
                "色觉障碍图像处理通常包括模拟和补偿两类任务。模拟的目标是估计色觉障碍用户看到的图像，"
                "例如 Brettel、Viénot 和 Mollon 提出的模型在 LMS 锥体响应空间中进行投影，"
                "能够模拟 protan（红色觉缺陷）、deutan（绿色觉缺陷）和 tritan（蓝黄色觉缺陷）等情况[7]；"
                "Machado 等提出的生理模型也常用于色觉障碍模拟[8]。补偿的目标则是重新调整颜色，"
                "让色觉障碍用户更容易区分原本混淆的颜色。",
            )
            continue
        text = text.replace("。[7][8]", "。")
        simple_replacements = {
            "27 幅最终评估图像和 16 幅额外 Kodak 图像": "27 幅最终评估图像和 24 幅额外 Kodak 图像",
            "27 幅最终评估图像和额外 16 幅 Kodak 图像": "27 幅最终评估图像和额外 24 幅 Kodak 图像",
            "27 幅主评估图像和额外 16 幅 Kodak 图像": "27 幅主评估图像和额外 24 幅 Kodak 图像",
            "16 additional Kodak images": "24 additional Kodak images",
            "额外 Kodak 09-24 共 16 幅图像": "额外 Kodak 01-24 共 24 幅图像",
            "Kodak 09-24 共 16 幅图像": "Kodak 01-24 共 24 幅图像",
            "额外 Kodak 16 幅图像": "额外 Kodak 24 幅图像",
            "Kodak 09-24": "Kodak 01-24",
            "图5 X 射线、焊缝和地图场景对比": "图5 X 射线、焊缝和纹理图像场景对比",
            "地图图像中，道路和区域边界的颜色差异也更明显": "纹理图像中，毛发、褶皱等纹理边界的颜色差异也更明显",
            "KSCgain 也略高于 CLAHE-Cividis": "KSCgain 与 CLAHE-Cividis 接近",
            "表2 额外 Kodak 16 幅图像上的关键方法验证": "表2 额外 Kodak 24 幅图像上的关键方法验证",
            "伪彩色处理并不是简单地“让图变好看”。": "伪彩色处理并不是简单地让图像更鲜艳。",
            "实际问题不是“灰图和彩图二选一”，而是“能否设计一种上色方式，在不破坏原灰度结构的前提下，让红绿色觉障碍用户也能获得额外的结构提示”。": "实际问题不是在灰图和彩色图之间二选一，而是能否设计一种上色方式，在不破坏原灰度结构的前提下，让红绿色觉障碍用户也能获得额外的结构提示。",
            "对红绿色觉障碍用户非常不友好。[3]": "对红绿色觉障碍用户非常不友好。",
        }
        normalized_text = text.translate(HALFWIDTH_DIGITS)
        for old, new in simple_replacements.items():
            text = text.replace(old, new)
        text = re.sub(
            r"(对红绿色觉障碍用户非常不友好。)(?:\[(?:[1-9]\d?)(?:[-,，](?:[1-9]\d?))*\])+",
            r"\1",
            text,
        )
        if text != paragraph.text.strip():
            set_text_hard(paragraph, text)
        key = text if text in replacements else text.translate(HALFWIDTH_DIGITS)
        if key in replacements:
            new_text = replacements[key]
            if new_text:
                set_text_hard(paragraph, new_text)
            else:
                remove_paragraph(paragraph)
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        norm = text.translate(HALFWIDTH_DIGITS)
        if norm.startswith("图7 给出了两种"):
            remove_paragraph(paragraph)


def insert_after_text_once(doc: Document, anchor_prefix: str, new_text: str) -> None:
    paragraphs = body_paragraphs(doc)
    if any(p.text.strip() == new_text for p in paragraphs):
        return
    for paragraph in paragraphs:
        if paragraph.text.strip().startswith(anchor_prefix):
            added = add_paragraph_after(paragraph, new_text)
            style_paragraph(added, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            return


def add_course_knowledge_sentence(doc: Document) -> None:
    return
    text = (
        "从数字图像处理知识点看，本文主要涉及空间域锐化、伪彩色增强和颜色空间变换，"
        "与图像增强和伪彩色显示内容相对应。"
    )
    insert_after_text_once(doc, "本文主要做了三件事", text)


def add_figure_overview_explanations(doc: Document) -> None:
    fig4_text = (
        "图4 给出主评估集 27 幅图像在固定参数下的 DE-Cividis 输出，用来说明主实验覆盖的图像类型。"
        "这些图像既包括自然图像、地图截图和 Ishihara 合成图，也包括 X 射线、焊缝、热成像、显微和降雨等灰度强度图像。"
        "由于每张图内容差异较大，后文只选取变化更容易观察的样例进行细节分析。"
    )
    fig7_text = (
        "图7 给出 Kodak 01-24 在固定参数下的 DE-Cividis 输出。Kodak 图像内容更接近自然照片，"
        "放在这里主要用于观察固定参数在另一组图像上是否出现明显异常。该图只展示 DE-Cividis 输出，"
        "Cividis、CLAHE-Cividis 和 DE-Cividis 的数值比较见表2。"
    )
    insert_after_text_once(doc, "图4 最终评估集", fig4_text)
    insert_after_text_once(doc, "图7 Kodak", fig7_text)


def split_fig2_fig3_explanations(doc: Document) -> None:
    old_prefixes = (
        "图2 用同一灰度图对比 Jet 和 Cividis",
        "图2 的作用是解释色图选择",
        "图2 对比了 Jet 和 Cividis",
        "图3 展示了 KSCgain",
        "图3 展示了参数",
        "图3 给出",
    )
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text.startswith(old_prefixes):
            remove_paragraph(paragraph)

    fig2_text = (
        "图2 对比了 Jet、Turbo、Viridis 和 Cividis 在正常视觉、protan（红色觉缺陷）和 deutan（绿色觉缺陷）模拟下的显示效果。"
        "Jet 和 Turbo 的颜色区分度较强，但在色觉障碍模拟下部分区域仍容易接近；Viridis 与 Cividis 的过渡更平稳，"
        "其中 Cividis 更适合作为红绿色觉障碍友好的基础色图。"
    )
    fig3_text = (
        "图3 给出 σ、α 和 λ 的小范围参数观察。σ 较小时，背景层过于接近原图，细节增强不明显；"
        "σ 或 α 过大时，边缘和噪声都更容易被拉强。λ 越接近 1，亮度越接近原灰度图；"
        "λ 较低时，增强后的灰度会更多进入亮度通道，局部对比更强，但明暗关系也更容易偏离原图。"
        "综合多类样例观察，后续实验固定为 σ = 2.0、α = 2.0、λ = 0.9。"
    )
    insert_after_text_once(doc, "图2 不同色图", fig2_text)
    insert_after_text_once(doc, "图3 参数", fig3_text)


FINAL_METHODS = [
    ("经典方法", "密度分层", "M1_intensity_slicing"),
    ("经典方法", "三相位正弦", "M2_sin_three_phase"),
    ("传统色图", "Jet", "M4_jet"),
    ("传统色图", "Hot", "M5_hot"),
    ("改进色图", "Turbo", "M5b_turbo"),
    ("感知均匀", "Viridis", "M6_viridis"),
    ("色觉友好", "Cividis", "M6b_cividis"),
    ("本文方法", "DE-Cividis", "DE_Cividis"),
]


def _set_row_values(row, values: list[str]) -> None:
    for cell, value in zip(row.cells, values):
        set_text(cell.paragraphs[0], value)
        for extra in cell.paragraphs[1:]:
            set_text(extra, "")
    for cell in row.cells[len(values):]:
        set_text(cell.paragraphs[0], "")
        for extra in cell.paragraphs[1:]:
            set_text(extra, "")


def _resize_table_rows(table, target_rows: int) -> None:
    while len(table.rows) < target_rows:
        table.add_row()
    while len(table.rows) > target_rows:
        table._tbl.remove(table.rows[-1]._tr)


def _resize_table_columns(table, target_cols: int) -> None:
    for row in table.rows:
        while len(row.cells) > target_cols:
            row._tr.remove(row.cells[-1]._tc)
    grid = table._tbl.tblGrid
    if grid is not None:
        grid_cols = list(grid.gridCol_lst)
        for grid_col in grid_cols[target_cols:]:
            grid.remove(grid_col)


def update_metric_tables_v2(doc: Document) -> None:
    summary_path = ROOT / "results" / "final_results_summary.csv"
    if not summary_path.exists() or len(doc.tables) < 2:
        return
    with open(summary_path, encoding="utf-8") as f:
        summary_rows = list(csv.DictReader(f))

    table1_header = ["类型", "方法", "ΔE2000 ↓", "Y-SSIM ↑", "RGB-SSIM ↑", "AvgGrad ↑", "Entropy ↑"]
    table1_metrics = ["delta_e2000", "ssim", "ssim_rgb", "avg_gradient", "entropy"]
    _resize_table_columns(doc.tables[0], len(table1_header))
    _resize_table_rows(doc.tables[0], len(FINAL_METHODS) + 1)
    _set_row_values(doc.tables[0].rows[0], table1_header)
    for row, (group, label, method) in zip(doc.tables[0].rows[1:], FINAL_METHODS):
        values = [group, label]
        values.extend(_fmt_metric(_summary_metric(summary_rows, "primary27", method, metric), metric) for metric in table1_metrics)
        _set_row_values(row, values)

    table2_header = ["类型", "方法", "ΔE2000 ↓", "Y-SSIM ↑", "RGB-SSIM ↑", "AvgGrad ↑", "Entropy ↑"]
    table2_metrics = ["delta_e2000", "ssim", "ssim_rgb", "avg_gradient", "entropy"]
    _resize_table_columns(doc.tables[1], len(table2_header))
    _resize_table_rows(doc.tables[1], len(FINAL_METHODS) + 1)
    _set_row_values(doc.tables[1].rows[0], table2_header)
    for row, (group, label, method) in zip(doc.tables[1].rows[1:], FINAL_METHODS):
        values = [group, label]
        values.extend(_fmt_metric(_summary_metric(summary_rows, "external_kodak24", method, metric), metric) for metric in table2_metrics)
        _set_row_values(row, values)


def finalize_metric_table_labels(doc: Document) -> None:
    if len(doc.tables) < 2:
        return
    headers = [
        ["类型", "方法", "ΔE2000 ↓", "Y-SSIM ↑", "RGB-SSIM ↑", "AvgGrad ↑", "Entropy ↑"],
        ["类型", "方法", "ΔE2000 ↓", "Y-SSIM ↑", "RGB-SSIM ↑", "AvgGrad ↑", "Entropy ↑"],
    ]
    for table, header in zip(doc.tables[:2], headers):
        _set_row_values(table.rows[0], header)


def rewrite_references_v2(doc: Document) -> None:
    refs = [
        "[1] Birch J. Worldwide prevalence of red-green color deficiency[J]. Journal of the Optical Society of America A, 2012, 29(3): 313-320. DOI: 10.1364/JOSAA.29.000313.",
        "[2] Gonzalez R C, Woods R E. Digital Image Processing[M]. 4th ed. New York: Pearson, 2018.",
        "[3] 章毓晋. 图像处理和分析基础[M]. 北京: 高等教育出版社, 2002.",
        "[4] Borland D, Taylor R M II. Rainbow color map (still) considered harmful[J]. IEEE Computer Graphics and Applications, 2007, 27(2): 14-17. DOI: 10.1109/MCG.2007.323435.",
        "[5] Nuñez J R, Anderton C R, Renslow R S. Optimizing colormaps with consideration for color vision deficiency to enable accurate interpretation of scientific data[J]. PLOS ONE, 2018, 13(7): e0199239. DOI: 10.1371/journal.pone.0199239.",
        "[6] Smith N J, van der Walt S. A better default colormap for matplotlib[EB/OL]. (2015-09-15)[2026-06-25]. https://bids.github.io/colormap/.",
        "[7] Mikhailov A. Turbo, an improved rainbow colormap for visualization[EB/OL]. (2019-08-20)[2026-06-25]. https://research.google/blog/turbo-an-improved-rainbow-colormap-for-visualization/.",
        "[8] Brettel H, Viénot F, Mollon J D. Computerized simulation of color appearance for dichromats[J]. Journal of the Optical Society of America A, 1997, 14(10): 2647-2655. DOI: 10.1364/JOSAA.14.002647.",
        "[9] Machado G M, Oliveira M M, Fernandes L A F. A physiologically-based model for simulation of color vision deficiency[J]. IEEE Transactions on Visualization and Computer Graphics, 2009, 15(6): 1291-1298. DOI: 10.1109/TVCG.2009.113.",
        "[10] Kuhn G R, Oliveira M M, Fernandes L A F. An efficient naturalness-preserving image-recoloring method for dichromats[J]. IEEE Transactions on Visualization and Computer Graphics, 2008, 14(6): 1747-1754. DOI: 10.1109/TVCG.2008.112.",
        "[11] Zhu Z, Mao X. Image recoloring for color vision deficiency compensation: a survey[J]. The Visual Computer, 2021, 37(12): 2999-3018. DOI: 10.1007/s00371-021-02240-0.",
        "[12] Tomasi C, Manduchi R. Bilateral filtering for gray and color images[C]//Proceedings of the 6th International Conference on Computer Vision. Bombay: IEEE, 1998: 839-846. DOI: 10.1109/ICCV.1998.710815.",
        "[13] He K, Sun J, Tang X. Guided image filtering[J]. IEEE Transactions on Pattern Analysis and Machine Intelligence, 2013, 35(6): 1397-1409. DOI: 10.1109/TPAMI.2012.213.",
        "[14] Sharma G, Wu W, Dalal E N. The CIEDE2000 color-difference formula: implementation notes, supplementary test data, and mathematical observations[J]. Color Research & Application, 2005, 30(1): 21-30. DOI: 10.1002/col.20070.",
        "[15] Wang Z, Bovik A C, Sheikh H R, Simoncelli E P. Image quality assessment: from error visibility to structural similarity[J]. IEEE Transactions on Image Processing, 2004, 13(4): 600-612. DOI: 10.1109/TIP.2003.819861.",
    ]
    paragraphs = list(doc.paragraphs)
    start = None
    for i, paragraph in enumerate(paragraphs):
        if paragraph.text.strip() == "参考文献":
            start = i
            break
    if start is None:
        return
    ref_paragraphs = [p for p in paragraphs[start + 1:] if p.text.strip()]
    for paragraph, text in zip(ref_paragraphs, refs):
        set_text(paragraph, text)
    for paragraph in ref_paragraphs[len(refs):]:
        remove_paragraph(paragraph)


def apply_final_baseline_rewrite(doc: Document) -> None:
    replacements = {
        "在前期实验中，本文重点比较了自适应 LUT、CLAHE-Cividis、融合基线": "综合固定色图的稳定性和结构增强需求后，最终采用 DE-Cividis，即 Detail-Enhanced Cividis。它不重新设计色图，而是在进入 Cividis 前增强灰度细节层，再通过亮度回写保持原灰度明暗结构。后续实验只把 DE-Cividis 作为本文方法，与经典伪彩色方法和公开色图进行比较。",
        "本文主要做了三件事": "本文主要完成三项工作。第一，组合反锐化细节增强、Cividis 色图和 YCbCr 亮度回写，形成一个面向红绿色觉障碍友好显示的灰度伪彩色流程。第二，使用 KSCgain 辅助观察边缘、纹理等关键区域在色觉障碍模拟后的可分辨性。第三，在 27 幅主评估图像和 24 幅 Kodak 补充图像上，与密度分层、三相位正弦变换、Jet、Hot、Turbo、Viridis 和 Cividis 进行比较。",
        "经典伪彩色方法通常可以分为三类": "经典伪彩色方法通常可以分为三类。第一类是密度分层，即把灰度区间划分为若干段，每一段赋予一种固定颜色[1-2]。这种方法易实现、易解释，适合突出特定灰度范围，但颜色变化不连续，分段边界可能被误认为真实结构。第二类是解析函数映射，例如用三相位正弦函数分别生成 R、G、B 通道[2]。这种方法连续性较好，但频率、相位等参数依赖人工设定。第三类是查表色图，如 Jet、Hot、Turbo、Viridis、Cividis 等[3,5-7]。",
        "Hot 色图相对单调": "Hot 色图相对单调，结构保真度通常比 Jet 更好，但颜色信息量较低，且在高亮区域容易饱和。Turbo 保留了彩虹色图较强的视觉区分度，同时改善了 Jet 的部分跳变问题[7]。Cividis 在设计时进一步考虑了色觉障碍下的可读性[5]，Viridis 则是常用的感知均匀色图之一[6]。因此，本文选择 Cividis 作为基础色图。",
        "色觉障碍图像处理通常包括模拟和补偿两类任务": "色觉障碍图像处理通常包括模拟和补偿两类任务。模拟的目标是估计色觉障碍用户看到的图像，例如 Brettel、Viénot 和 Mollon 提出的模型在 LMS 锥体响应空间中进行投影，能够模拟 protan（红色觉缺陷）、deutan（绿色觉缺陷）和 tritan（蓝黄色觉缺陷）等情况[8]；Machado 等提出的生理模型也常用于色觉障碍模拟[9]。补偿的目标则是重新调整颜色，让色觉障碍用户更容易区分原本混淆的颜色。",
        "Daltonization 是典型的色觉补偿思路": "色觉补偿和重着色方法会重新调整已有 RGB 图像的颜色，使色觉障碍用户更容易区分原本混淆的区域。例如 Kuhn 等提出自然性保持的图像重着色方法，在增强可分辨性的同时尽量保持图像自然性[10]。这类方法主要面向已有 RGB 彩色图像；本文关注的是灰度图到伪彩色图的生成。",
        "也有一些更复杂的图像重着色方法": "也有一些更复杂的图像重着色方法[11]，可以处理更一般的彩色图像，但通常需要更多参数或优化过程。本文关注的是灰度图到伪彩色图的轻量转换，因此主实验采用经典方法和公开色图作为对比对象。",
        "双边滤波和 Guided Filter": "双边滤波和 Guided Filter 都可以得到边缘保持的背景层[12-13]。理论上，它们比高斯滤波更能避免跨边缘模糊，适合分离纹理和结构。不过这类方法会增加窗口半径、正则化系数等设置。为了保持流程简洁，DE-Cividis 采用参数更少的高斯背景层估计。",
        "多尺度反锐化进一步": "多尺度反锐化可以把不同尺度的细节层加权融合，用来同时增强细边缘和中尺度纹理。本文最终没有采用多尺度形式，主要是因为单尺度反锐化已经能提供较稳定的结构提示，参数也更容易控制。",
        "在确定最终流程前，本文比较了几种候选方案": "DE-Cividis 的设计重点是把结构增强放在色图映射之前。固定 Cividis 只根据灰度值查表，无法区分同一灰度值处在背景还是边缘位置；DE-Cividis 先在空间域计算细节层，使边缘和纹理在进入色图前获得更大的灰度差。",
        "主要原因是，自适应 LUT": "因此，最终流程分为三个环节：先做细节增强，再用 Cividis 上色，最后回写原灰度亮度。前两步负责增加结构区域的颜色差异，最后一步负责保持原图的明暗关系。",
        "主实验重点比较 Jet、Hot、Viridis、Cividis、CLAHE-Cividis": "主实验比较 8 种方法：密度分层、三相位正弦变换、Jet、Hot、Turbo、Viridis、Cividis 和 DE-Cividis。前两种属于经典伪彩色方法，Jet、Hot 和 Turbo 代表常见色图，Viridis 和 Cividis 代表感知均匀或色觉友好的公开色图，DE-Cividis 是本文方法。",
        "本文使用五类指标评价结果": "本文使用五类指标评价结果。CIEDE2000 色差（ΔE2000）[14]用于观察输出图与原灰度结构之间的感知差异，数值越小表示改动越保守；Y-SSIM（亮度通道结构相似度）和 RGB-SSIM（RGB 三通道结构相似度）[15]用于观察结构保持情况；EdgeGain（边缘增益）用 Sobel 梯度比较色觉障碍模拟后的边缘强度；Entropy（信息熵）用于描述输出颜色变化的丰富程度。",
        "所有对比方法均在相同预处理后的灰度图上运行": "所有对比方法均在相同预处理后的灰度图上运行，并统一进行 protan 和 deutan 两类色觉障碍模拟[8-9]。密度分层、三相位正弦变换和各类色图都直接从同一归一化灰度图生成输出；DE-Cividis 额外进行细节增强和亮度回写。这样可以保证比较对象对应同一输入信息。",
        "图2 对比了 Jet 和 Cividis": "图2 对比了 Jet、Turbo、Viridis 和 Cividis 在正常视觉、protan（红色觉缺陷）和 deutan（绿色觉缺陷）模拟下的显示效果。Jet 和 Turbo 的颜色区分度较强，但在色觉障碍模拟下部分区域仍容易接近；Viridis 与 Cividis 的过渡更平稳，其中 Cividis 更符合本文的红绿色觉障碍友好目标。",
        "表1 汇总了 27 幅": "表1 汇总了 27 幅主评估图像上的平均结果。固定 Cividis 的 ΔE2000 较低，说明它对原灰度变化最保守；DE-Cividis 的 ΔE2000 略高于 Cividis，但明显低于 Jet、Hot、Turbo 等颜色变化更强的色图。",
        "在结构相似度方面": "在结构相似度方面，DE-Cividis 通过亮度回写保持了较高的 Y-SSIM。RGB-SSIM 会低于固定 Cividis，原因是 DE-Cividis 有意在边缘和纹理区域增加颜色差异，而不是让三个颜色通道完全接近灰度图。",
        "在 KSCgain 上": "在 KSCgain 上，DE-Cividis 相比固定 Cividis 有明显提升，说明细节增强后，边缘和纹理区域在色觉障碍模拟下更容易形成可见差异。密度分层、三相位正弦和 Jet 也可能产生较强颜色变化，但色差和结构稳定性通常不如 Cividis 系方法。",
        "EdgeGain 方面": "EdgeGain 方面，DE-Cividis 的提升并不只来自整体亮度边缘变强。结合 KSCgain 可以看出，方法主要把变化集中到原图已有的边缘、纹理和局部突变位置。",
        "Entropy 方面": "Entropy 方面，DE-Cividis 高于固定 Cividis，说明输出包含更丰富的颜色变化。颜色信息量增加本身不一定代表效果更好，因此仍需结合 ΔE2000、SSIM、KSCgain 和可视化结果综合判断。",
        "图5 选取了 KSCgain": "图5 选取了 X 射线、焊缝和纹理图像三类样例。Jet 和 Turbo 的色彩变化更强，但容易带来较夸张的颜色层次；Viridis 和 Cividis 更平稳，不过弱边缘不一定突出。DE-Cividis 在 Cividis 的基础上增强细节层，因此裂纹、边界和细纹理位置更容易观察。",
        "在最终统计的候选方法中": "表2 使用同一套 8 种方法在 Kodak 01-24 上做补充观察。Kodak 图像更接近自然照片，放在这里主要用于检查固定参数在另一组图像上是否仍能稳定运行，而不是重新选择参数。",
        "额外 Kodak 01-24": "额外 Kodak 01-24 共 24 幅图像用于补充验证。与主评估集相同，这部分也使用密度分层、三相位正弦变换、Jet、Hot、Turbo、Viridis、Cividis 和 DE-Cividis。",
        "表2 表明": "表2 表明，在 Kodak 图像上，固定 Cividis 仍然是较保守的基线，DE-Cividis 则在保持亮度结构的同时提高 KSCgain。由于 Kodak 图像含有更多自然场景纹理，结果主要作为固定参数的补充观察。",
        "这组结果与主实验基本一致": "这组结果与主实验的趋势基本一致：DE-Cividis 相比固定 Cividis 提供了更高的关键结构对比。所有参数在补充验证前已经固定为 σ = 2.0、α = 2.0。",
        "图7 给出 Kodak": "图7 给出 Kodak 01-24 在固定参数下的 DE-Cividis 输出。Kodak 图像内容更接近自然照片，用于观察固定参数在另一组图像上是否出现明显异常。相关数值比较见表2。",
        "从指标上看，固定 Cividis": "从指标上看，固定 Cividis 的 ΔE2000 较低，说明它与原灰度亮度变化更接近，但它的 KSCgain 被定义为 1，缺少额外结构增强能力。Jet、Hot 和 Turbo 颜色更强，可能增加局部可见差异，也更容易带来较大色差。DE-Cividis 的位置介于两者之间：它增加少量色差，换取更高的关键结构对比，同时保持较高的亮度结构相似度。",
        "与多模块融合基线相比": "DE-Cividis 的处理链较短。它先增强细节，再用 Cividis 上色，最后把原灰度亮度写回去。三个步骤分别对应结构、颜色和亮度三个问题，模块关系比较清楚。",
        "与自适应 LUT 方法相比": "固定查表色图只改变灰度值到颜色的对应关系，本质上仍是一维映射；如果同一灰度值同时出现在背景和边缘区域，它无法根据空间邻域改变映射。DE-Cividis 在空间域计算细节层，因此同一灰度值在不同局部结构中可以产生不同的增强结果。",
        "与 CLAHE-Cividis 相比": "参数方面，σ 决定背景层平滑尺度，α 决定细节层回加强度。两者取 2.0 是在多类样例上观察后的折中：取值过小，弱边缘变化不明显；取值过大，噪声和局部过冲更容易出现。",
        "本文实验代码按功能划分": "实验代码按功能划分为数据读取、经典伪彩色方法、DE-Cividis、色觉障碍模拟、评价指标和最终统计几个部分。表格结果由程序重新计算后汇总，避免手工誊写造成误差。",
        "实验结果表明": "实验结果表明，DE-Cividis 相比固定 Cividis 能提高 KSCgain，并在亮度结构保持和关键结构提示之间取得较好的折中。与密度分层、三相位正弦变换以及 Jet、Hot、Turbo 等色图相比，它的颜色变化更克制；与 Viridis、Cividis 相比，它对边缘和纹理的提示更明显。在本文测试的图像范围内，DE-Cividis 可作为一种轻量的灰度伪彩色显示方案。",
    }
    remove_prefixes = (
        "从数字图像处理知识点看",
        "本文没有使用深度学习模型",
        "表2 使用同一套 8 种方法",
    )
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text.startswith(remove_prefixes):
            remove_paragraph(paragraph)
            continue
        for prefix, replacement in replacements.items():
            if text.startswith(prefix):
                set_text_hard(paragraph, replacement)
                break
    phrase_replacements = {
        "这也是本文工作的出发点。": "这一点也是方法设计的出发点。",
        "本文主要完成三项工作。": "主要完成三项工作。",
        "因此，本文选择 Cividis 作为基础色图。": "因此，Cividis 被选作基础色图。",
        "但本文讨论的是灰度图到伪彩色图的生成任务。": "但这里讨论的是灰度图到伪彩色图的生成任务。",
        "本文关注的是灰度图到伪彩色图的轻量转换": "研究对象是灰度图到伪彩色图的轻量转换",
        "本文最终没有采用多尺度形式": "最终没有采用多尺度形式",
        "本文的核心问题是": "核心问题是",
        "本文所说的关键结构": "这里所说的关键结构",
        "本文不使用训练数据": "方法不使用训练数据",
        "本文方法有三个约束": "该方法有三个约束",
        "本文使用高斯滤波得到": "使用高斯滤波得到",
        "本文把初始伪彩色图": "将初始伪彩色图",
        "本文使用五类指标评价结果": "实验使用五类指标评价结果",
        "本文还计算 KSCgain": "还计算 KSCgain",
        "本文方法在边缘、裂纹和纹理较明显的灰度图上": "该方法在边缘、裂纹和纹理较明显的灰度图上",
        "本文的评价主要基于": "评价主要基于",
        "本文围绕红绿色觉障碍用户的": "围绕红绿色觉障碍用户的",
        "在本文测试的图像范围内": "在本实验测试的图像范围内",
        "本文将初始伪彩色结果": "将初始伪彩色结果",
        "本文选择 YCbCr": "选择 YCbCr",
        "更符合本文的结构保真目标": "更符合结构保真目标",
        "单通道强度分布[1-2]。": "单通道强度分布。",
        "尤其 Cividis 在设计时考虑了色觉障碍视角，适合作为灰度伪彩色的安全基底[5]。": "尤其 Cividis 在设计时考虑了色觉障碍视角，适合作为灰度伪彩色的安全基底。",
        "例如用三相位正弦函数分别生成 R、G、B 通道[2]。": "例如用三相位正弦函数分别生成 R、G、B 通道。",
        "Turbo 保留了彩虹色图较强的视觉区分度，同时改善了 Jet 的部分跳变问题[7]。Viridis 和 Cividis 属于更现代的感知均匀色图[5-6]，其中 Cividis 在设计时进一步考虑了色觉障碍下的可读性。": "Cividis 在设计时进一步考虑了色觉障碍下的可读性[5]，Viridis 则是常用的感知均匀色图之一[6]。Turbo 保留了彩虹色图较强的视觉区分度，同时改善了 Jet 的部分跳变问题[7]。",
    }
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if not text or text == "参考文献":
            continue
        updated = text
        for old, new in phrase_replacements.items():
            updated = updated.replace(old, new)
        if updated != text:
            set_text_hard(paragraph, updated)
    dedupe_prefixes = (
        "图4 给出主评估集 27 幅图像",
        "图7 给出 Kodak 01-24",
    )
    seen_prefixes: set[str] = set()
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        for prefix in dedupe_prefixes:
            if text.startswith(prefix):
                if prefix in seen_prefixes:
                    remove_paragraph(paragraph)
                else:
                    seen_prefixes.add(prefix)
                break
    rewrite_references_v2(doc)


def add_limitations_intro(doc: Document) -> None:
    intro = (
        "结合可视化和指标结果，本文方法在边缘、裂纹和纹理较明显的灰度图上效果更容易观察。"
        "但它仍然是一个轻量级图像处理流程，使用时需要注意以下几类情况。"
    )
    paragraphs = body_paragraphs(doc)
    if any(p.text.strip() == intro for p in paragraphs):
        return
    for paragraph in paragraphs:
        if paragraph.text.strip() in {"6.4 局限性分析", "6.4 局限性"}:
            added = add_paragraph_after(paragraph, intro)
            style_paragraph(added, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            return


def polish_algorithm_block(doc: Document) -> None:
    in_block = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text.startswith("算法 1") or text.startswith("算法1"):
            set_text(paragraph, text.replace("算法 1", "算法1"))
            in_block = True
            style_paragraph(paragraph, east="\u9ed1\u4f53", size=10.5, bold=True,
                            align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=4, after=2,
                            line_spacing=1.15)
            paragraph.paragraph_format.left_indent = Cm(0.35)
            paragraph.paragraph_format.right_indent = Cm(0.35)
            set_paragraph_border(paragraph, top=True, bottom=False)
            continue
        if in_block:
            if not text:
                in_block = False
                continue
            if text.startswith("输入") or text.startswith("输出"):
                cleaned = text.rstrip("。.")
                if cleaned != text:
                    set_text(paragraph, cleaned)
                    text = cleaned
                style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                                align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=0, after=1,
                                line_spacing=1.15)
                paragraph.paragraph_format.left_indent = Cm(0.35)
                paragraph.paragraph_format.right_indent = Cm(0.35)
                paragraph.paragraph_format.first_line_indent = Cm(0.70)
                set_paragraph_border(paragraph, top=False, bottom=text.startswith("输出"))
                continue
            if re.match(r"^\d+\.", text):
                cleaned = text.rstrip("。.")
                if cleaned != text:
                    set_text(paragraph, cleaned)
                style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                                align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=0, after=1,
                                line_spacing=1.15)
                paragraph.paragraph_format.left_indent = Cm(1.30)
                paragraph.paragraph_format.right_indent = Cm(0.35)
                paragraph.paragraph_format.first_line_indent = Cm(-0.35)
                continue
            in_block = False


def set_cell_borders(cell, *, top=None, bottom=None, left="nil", right="nil") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right),
                      ("insideH", "nil"), ("insideV", "nil")):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        if val is None:
            element.set(qn("w:val"), "nil")
        elif val == "nil":
            element.set(qn("w:val"), "nil")
        else:
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), str(val))
            element.set(qn("w:space"), "0")
            element.set(qn("w:color"), "000000")


def set_table_borders_nil(table) -> None:
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
        element = borders.find(qn("w:" + edge))
        if element is None:
            element = OxmlElement("w:" + edge)
            borders.append(element)
        element.set(qn("w:val"), "nil")


def polish_tables(doc: Document) -> None:
    widths = [
        [Cm(1.9), Cm(3.0), Cm(1.7), Cm(1.7), Cm(1.8), Cm(1.8), Cm(1.7)],
        [Cm(1.9), Cm(3.0), Cm(1.7), Cm(1.7), Cm(1.8), Cm(1.8), Cm(1.7)],
    ]
    header_labels = [
        ["类型", "方法", "ΔE2000 ↓", "Y-SSIM ↑", "RGB-SSIM ↑", "AvgGrad ↑", "Entropy ↑"],
        ["类型", "方法", "ΔE2000 ↓", "Y-SSIM ↑", "RGB-SSIM ↑", "AvgGrad ↑", "Entropy ↑"],
    ]
    label_replacements = {
        "感知均匀色图": "感知均匀",
        "色觉友好基线": "色觉友好",
        "结构重着色": "结构重着",
        "Jet 彩虹色图": "Jet",
        "Hot 热色图": "Hot",
        "多模块融合基线": "融合基线",
        "显著性重着色": "显著性重着",
    }
    for ti, table in enumerate(doc.tables):
        table.style = "Normal Table"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_borders_nil(table)
        col_widths = widths[min(ti, len(widths) - 1)]
        last_row = len(table.rows) - 1
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                if ci < len(col_widths):
                    cell.width = col_widths[ci]
                if ri == 0 and ti < len(header_labels) and ci < len(header_labels[ti]):
                    set_text_hard(cell.paragraphs[0], header_labels[ti][ci])
                    for extra in cell.paragraphs[1:]:
                        set_text(extra, "")
                raw_text = cell.text.strip()
                if raw_text in label_replacements:
                    for paragraph in cell.paragraphs:
                        set_text(paragraph, "")
                    cell.paragraphs[0].add_run(label_replacements[raw_text])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.first_line_indent = None
                    paragraph.paragraph_format.line_spacing = 1.15
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    for run in paragraph.runs:
                        emphasis = ri == 0
                        set_run_font(run, east="\u5b8b\u4f53", size=8.5, bold=emphasis)
                top = 8 if ri == 0 else None
                bottom = 8 if ri in (0, last_row) else None
                set_cell_borders(cell, top=top, bottom=bottom)


def _summary_metric(summary_rows: list[dict[str, str]], dataset: str, method: str, metric: str) -> float | None:
    values = [
        float(row[f"{metric}_mean"])
        for row in summary_rows
        if row.get("dataset") == dataset and row.get("method") == method and row.get("cvd_type") in CVD_TYPES
    ]
    if not values:
        return None
    return sum(values) / len(values)


CVD_TYPES = ("protan", "deutan")


def _fmt_metric(value: float | None, metric: str | None = None) -> str:
    if value is None:
        return ""
    if metric in {"ssim", "ssim_rgb", "avg_gradient"}:
        return f"{value:.3f}"
    return f"{value:.2f}"


def update_metric_tables(doc: Document) -> None:
    summary_path = ROOT / "results" / "final_results_summary.csv"
    if not summary_path.exists() or len(doc.tables) < 2:
        return
    with open(summary_path, encoding="utf-8") as f:
        summary_rows = list(csv.DictReader(f))

    table1_methods = {
        "Jet": "M4_jet",
        "Hot": "M5_hot",
        "Viridis": "M6_viridis",
        "Cividis": "M6b_cividis",
        "Daltonization": "M7_daltonize",
        "显著性重着": "M8_saliency_aware",
        "融合基线": "M9_old",
        "自适应 LUT": "M10_adaptive_LUT",
        "CLAHE-Cividis": "M12_CLAHE_cividis",
        "DE-Cividis": "DE_Cividis",
    }
    table1_metrics = ["delta_e2000", "ssim", "ssim_rgb", "avg_gradient", "entropy"]
    for row in doc.tables[0].rows[1:]:
        method_label = row.cells[1].text.strip()
        method = table1_methods.get(method_label)
        if not method:
            continue
        for offset, metric in enumerate(table1_metrics, start=2):
            set_text(row.cells[offset].paragraphs[0], _fmt_metric(_summary_metric(summary_rows, "primary27", method, metric), metric))

    table2_methods = {
        "Cividis": "M6b_cividis",
        "CLAHE-Cividis": "M12_CLAHE_cividis",
        "DE-Cividis": "DE_Cividis",
    }
    table2_metrics = ["delta_e2000", "ssim", "ssim_rgb", "avg_gradient", "entropy"]
    for row in doc.tables[1].rows[1:]:
        method_label = row.cells[1].text.strip()
        method = table2_methods.get(method_label)
        if not method:
            continue
        for offset, metric in enumerate(table2_metrics, start=2):
            set_text(row.cells[offset].paragraphs[0], _fmt_metric(_summary_metric(summary_rows, "external_kodak24", method, metric), metric))


def add_page_number_footer(doc: Document) -> None:
    section = doc.sections[0]
    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    set_text(paragraph, "")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = None
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_sep, text, fld_end])
    set_run_font(run, east="\u5b8b\u4f53", size=9, bold=False)


def resize_images(doc: Document) -> None:
    figure_files = [
        "fig1_pipeline.png",
        "fig2_colormap_cvd.png",
        "fig3_saliency.png",
        "fig4_eval_overview.png",
        "fig5_typical_cases.png",
        "fig6_ablation.png",
        "fig7_kodak_overview.png",
    ]
    max_width = Cm(15.8)
    max_height = Cm(10.8)
    for shape, filename in zip(doc.inline_shapes, figure_files):
        image_path = ROOT / "figs" / filename
        if not image_path.exists():
            continue
        with Image.open(image_path) as image:
            px_w, px_h = image.size
        width = max_width
        height = int(width * px_h / px_w)
        if height > max_height:
            height = max_height
            width = int(height * px_w / px_h)
        shape.width = width
        shape.height = height


def keep_52_heading_with_figure(doc: Document) -> None:
    intro_text = (
        "为配合平均指标结果，下面选取 X 射线、焊缝和纹理图像三类样例进行观察。"
        "这些样例的边缘或纹理较明显，更容易看出不同伪彩色方法的显示差异。"
    )
    paragraphs = body_paragraphs(doc)
    heading_idx = next((i for i, p in enumerate(paragraphs) if p.text.strip() == "5.2 典型图像对比"), None)
    for idx, paragraph in enumerate(list(paragraphs)):
        if paragraph.text.strip() == intro_text and idx != (heading_idx + 1 if heading_idx is not None else -1):
            remove_paragraph(paragraph)

    for paragraph in body_paragraphs(doc):
        if paragraph.text.strip() == "5.2 典型图像对比":
            paragraph.style = doc.styles["Heading 2"]
            style_paragraph(paragraph, east="\u9ed1\u4f53", size=12, bold=True,
                            align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=5, after=4)
            pf = paragraph.paragraph_format
            pf.page_break_before = False
            pf.keep_with_next = True
            pf.space_after = Pt(4)
            paragraphs = body_paragraphs(doc)
            idx = next((i for i, p in enumerate(paragraphs) if p._p is paragraph._p), None)
            next_text = paragraphs[idx + 1].text.strip() if idx is not None and idx + 1 < len(paragraphs) else ""
            if next_text != intro_text:
                added = add_paragraph_after(paragraph, intro_text)
                style_paragraph(added, east="\u5b8b\u4f53", size=10.5, bold=False,
                                align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
                added.paragraph_format.keep_with_next = True
            else:
                paragraphs[idx + 1].paragraph_format.keep_with_next = True
            break


def ensure_fig3_explanation(doc: Document) -> None:
    explanation = (
        "图3 给出 σ、α 和 λ 的小范围参数观察。σ 较小时，背景层过于接近原图，细节增强不明显；"
        "σ 或 α 过大时，边缘和噪声都更容易被拉强。λ 越接近 1，亮度越接近原灰度图；"
        "λ 较低时，增强后的灰度会更多进入亮度通道，局部对比更强，但明暗关系也更容易偏离原图。"
        "综合多类样例观察，后续实验固定为 σ = 2.0、α = 2.0、λ = 0.9。"
    )
    for paragraph in list(body_paragraphs(doc)):
        if paragraph.text.strip().startswith("图3 给出"):
            remove_paragraph(paragraph)
    for paragraph in body_paragraphs(doc):
        if paragraph.text.strip() == "图3 参数选择对输出效果的影响":
            added = add_paragraph_after(paragraph, explanation)
            style_paragraph(added, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            return


def add_space_after_tables(doc: Document) -> None:
    body = doc._body._element
    children = list(body)
    for idx, child in enumerate(children[:-1]):
        if child.tag != qn("w:tbl"):
            continue
        for next_child in children[idx + 1:]:
            if next_child.tag != qn("w:p"):
                continue
            paragraph = Paragraph(next_child, doc._body)
            if not paragraph.text.strip():
                continue
            paragraph.paragraph_format.space_before = Pt(6)
            break


def merge_discussion_and_conclusion(doc: Document) -> None:
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if text == "6 分析与讨论":
            set_text_hard(paragraph, "6 结论")
            paragraph.style = doc.styles["Heading 1"]
            style_paragraph(paragraph, east="\u9ed1\u4f53", size=14, bold=True,
                            align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=8, after=4)
        elif text.startswith("7 总结"):
            remove_paragraph(paragraph)


def move_kodak_overview_to_discussion(doc: Document) -> None:
    for paragraph in list(body_paragraphs(doc)):
        if paragraph.text.strip().startswith("图7 给出 Kodak"):
            remove_paragraph(paragraph)
    remove_caption_and_previous_drawing(doc, "图7")

    anchor = None
    for paragraph in body_paragraphs(doc):
        if paragraph.text.strip().startswith("这组结果与主实验的趋势基本一致"):
            anchor = paragraph
            break
    if anchor is None:
        for paragraph in body_paragraphs(doc):
            if paragraph.text.strip().startswith("Kodak 补充验证结果显示"):
                anchor = paragraph
                break
    if anchor is None:
        return

    image_path = ROOT / "figs" / "fig7_kodak_overview.png"
    if not image_path.exists():
        return

    image_para = add_paragraph_after(anchor, "")
    image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_para.paragraph_format.first_line_indent = None
    image_para.paragraph_format.space_before = Pt(6)
    image_para.paragraph_format.space_after = Pt(3)
    image_para.add_run().add_picture(str(image_path), width=Cm(15.8))

    caption = add_paragraph_after(image_para, "图7 补充验证图像输出总览")
    style_paragraph(caption, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                    bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=1, after=0)
    disable_east_asian_auto_spacing(caption)

    note = add_paragraph_after(
        caption,
        "图7 给出 Kodak 01-24 在固定参数下的 DE-Cividis 输出。Kodak 图像内容更接近自然照片，"
        "用于补充观察固定参数在另一组图像上的整体表现；相关数值比较见表2。"
    )
    style_paragraph(note, east="\u5b8b\u4f53", size=10.5, bold=False,
                    align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)


def polish_references(doc: Document) -> None:
    in_refs = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text == "\u53c2\u8003\u6587\u732e":
            in_refs = True
            style_paragraph(paragraph, east="\u9ed1\u4f53", size=14, bold=True,
                            align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=8, after=4)
            continue
        if in_refs and text:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.line_spacing = 1.05
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.left_indent = Cm(0.62)
            paragraph.paragraph_format.first_line_indent = Cm(-0.62)
            for run in paragraph.runs:
                if run.text:
                    set_run_font(run, east="\u5b8b\u4f53", latin="Times New Roman", size=8.5, bold=False, italic=False)


def trim_trailing_empty_paragraphs(doc: Document) -> None:
    for paragraph in reversed(doc.paragraphs):
        if paragraph.text.strip() or "w:drawing" in paragraph._p.xml:
            break
        remove_paragraph(paragraph)


def remove_adjacent_duplicate_paragraphs(doc: Document) -> None:
    previous = None
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text and text == previous and "w:drawing" not in paragraph._p.xml:
            remove_paragraph(paragraph)
            continue
        previous = text if text else None


def apply_no_ksc_final_rewrite(doc: Document) -> None:
    replacements = {
        "主要完成三项工作。": (
            "主要完成三项工作。第一，组合反锐化细节增强、Cividis 色图和 YCbCr 亮度回写，"
            "形成一个面向红绿色觉障碍友好显示的灰度伪彩色流程。第二，在 27 幅主评估图像和 "
            "24 幅 Kodak 补充图像上，与密度分层、三相位正弦变换、Jet、Hot、Turbo、Viridis 和 Cividis 进行比较。"
            "第三，结合色差、结构相似度、平均梯度、信息熵和典型图像分析方法效果。"
        ),
        "本文主要完成三项工作。": (
            "主要完成三项工作。第一，组合反锐化细节增强、Cividis 色图和 YCbCr 亮度回写，"
            "形成一个面向红绿色觉障碍友好显示的灰度伪彩色流程。第二，在 27 幅主评估图像和 "
            "24 幅 Kodak 补充图像上，与密度分层、三相位正弦变换、Jet、Hot、Turbo、Viridis 和 Cividis 进行比较。"
            "第三，结合色差、结构相似度、平均梯度、信息熵和典型图像分析方法效果。"
        ),
        "实验使用五类指标评价结果": (
            "实验使用五类指标评价结果。CIEDE2000 色差（ΔE2000）[14]用于观察输出图与原灰度结构之间的感知差异，"
            "数值越小表示改动越保守；Y-SSIM（亮度通道结构相似度）和 RGB-SSIM（RGB 三通道结构相似度）[15]"
            "用于观察结构保持情况；AvgGrad（平均梯度）用 Sobel 梯度描述色觉障碍模拟后的边缘和纹理强度；"
            "Entropy（信息熵）用于描述输出颜色变化的丰富程度。"
        ),
        "本文使用五类指标评价结果": (
            "实验使用五类指标评价结果。CIEDE2000 色差（ΔE2000）[14]用于观察输出图与原灰度结构之间的感知差异，"
            "数值越小表示改动越保守；Y-SSIM（亮度通道结构相似度）和 RGB-SSIM（RGB 三通道结构相似度）[15]"
            "用于观察结构保持情况；AvgGrad（平均梯度）用 Sobel 梯度描述色觉障碍模拟后的边缘和纹理强度；"
            "Entropy（信息熵）用于描述输出颜色变化的丰富程度。"
        ),
        "图3 展示了 KSCgain": (
            "图3 给出不同参数下的输出效果。σ 较小时，背景层过于接近原图，细节增强不明显；"
            "σ 或 α 过大时，边缘和噪声都更容易被拉强。综合多类样例观察，后续实验固定为 σ=2.0、α=2.0。"
        ),
        "在 KSCgain 上": (
            "AvgGrad 方面，DE-Cividis 相比固定 Cividis 有一定提升，说明在色觉障碍模拟后的亮度图中，"
            "边缘和纹理变化更明显。密度分层、三相位正弦和 Jet 也可能产生较强变化，但色差和结构稳定性通常不如 Cividis 系方法。"
        ),
        "EdgeGain 方面": (
            "AvgGrad 方面，DE-Cividis 的数值高于固定 Cividis，说明细节增强后边缘和纹理更容易形成局部变化。"
            "不过平均梯度也可能受到噪声影响，因此不能单独作为优劣判断依据。"
        ),
        "Entropy 方面": (
            "Entropy 方面，DE-Cividis 高于固定 Cividis，说明输出包含更丰富的颜色变化。"
            "颜色信息量增加本身不一定代表效果更好，因此仍需结合 ΔE2000、SSIM、AvgGrad 和可视化结果综合判断。"
        ),
        "表2 表明": (
            "表2 表明，在 Kodak 图像上，固定 Cividis 仍然是较保守的基线，DE-Cividis 则在保持亮度结构的同时提高 AvgGrad。"
            "由于 Kodak 图像含有更多自然场景纹理，结果主要作为固定参数的补充观察。"
        ),
        "从指标上看，固定 Cividis": (
            "从指标上看，固定 Cividis 的 ΔE2000 较低，说明它与原灰度亮度变化更接近。Jet、Hot 和 Turbo 颜色更强，"
            "可能增加局部可见差异，也更容易带来较大色差。DE-Cividis 的位置介于两者之间："
            "它增加少量色差，换取更明显的边缘和纹理提示，同时保持较高的亮度结构相似度。"
        ),
        "实验结果表明": (
            "实验结果表明，DE-Cividis 在亮度结构保持和边缘纹理提示之间取得了较好的折中。"
            "与密度分层、三相位正弦变换以及 Jet、Hot、Turbo 等色图相比，它的颜色变化更克制；"
            "与 Viridis、Cividis 相比，它对边缘和纹理的提示更明显。在本实验测试的图像范围内，"
            "DE-Cividis 可作为一种轻量的灰度伪彩色显示方案。"
        ),
        "多个指标需要一起解释": (
            "多个指标需要一起解释。色差低说明输出保守，但不一定更醒目；平均梯度高说明边缘和纹理变化更强，"
            "但也可能来自噪声或伪边界。因此，实验把色差、结构相似度、平均梯度、信息熵和图像对比放在一起看，"
            "不把某一个指标作为唯一目标。"
        ),
        "第一，本文使用色觉障碍模拟和客观指标进行评价": (
            "第一，本文使用色觉障碍模拟和客观指标进行评价，没有招募真实色觉障碍用户完成任务实验。"
            "后续工作可以设计边界识别、异常定位或强度排序任务，让真实用户参与评价，"
            "从而检验客观指标与主观体验之间的一致性。"
        ),
    }
    remove_prefixes = (
        "为了更贴近本文关心的边缘和纹理区域，实验中还加入了一个自定义辅助量 KSCgain",
        "除上述指标外，本文还计算 KSCgain",
        "KSCgain 不是用户实验",
        "KSCgain 的作用是",
        "该指标只作为辅助观察量",
        "KSCgain 提供的是一个辅助观察角度",
    )
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text.startswith(remove_prefixes):
            remove_paragraph(paragraph)
            continue
        for prefix, replacement in replacements.items():
            if text.startswith(prefix):
                set_text_hard(paragraph, replacement)
                break

    direction_note = (
        "表头中的 ↑ 和 ↓ 只表示指标方向：↑ 表示数值越大通常越明显或越接近目标，"
        "↓ 表示数值越小越保守；它们不是指相对于某一种方法的升降。"
    )
    if not any(p.text.strip() == direction_note for p in body_paragraphs(doc)):
        for paragraph in body_paragraphs(doc):
            if paragraph.text.strip().startswith("实验使用五类指标评价结果"):
                added = add_paragraph_after(paragraph, direction_note)
                style_paragraph(added, east="\u5b8b\u4f53", size=10.5, bold=False,
                                align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
                break

    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if "KSCgain" in text:
            remove_paragraph(paragraph)


def replace_final_figure_captions(doc: Document) -> None:
    captions = {
        "图1 DE-Cividis 方法总体流程": "图1 细节增强伪彩色方法总体流程",
        "图3 KSCgain 辅助指标中的边缘和局部对比权重": "图3 参数选择对输出效果的影响",
        "图3 关键结构增益辅助指标中的边缘和局部对比权重": "图3 参数选择对输出效果的影响",
        "图3 参数变化对输出效果的影响": "图3 参数选择对输出效果的影响",
        "图3 参数变化对输出效果的影响（λ = 0.9）": "图3 参数选择对输出效果的影响",
        "图3 σ、α 和 λ 参数变化对输出效果的影响": "图3 参数选择对输出效果的影响",
        "图4 最终评估集 27 幅图像的 DE-Cividis 输出总览": "图4 最终评估集 27 幅图像输出总览",
        "图6 DE-Cividis 消融实验可视化结果": "图6 方法消融实验可视化结果",
        "图7 Kodak 01-24 补充验证图像的 DE-Cividis 输出总览": "图7 补充验证图像输出总览",
    }
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if text in captions:
            set_text_hard(paragraph, captions[text])


def final_caption_and_duplicate_cleanup(doc: Document) -> None:
    final_captions = {
        "\u56fe1 \u7ec6\u8282\u589e\u5f3a\u4f2a\u5f69\u8272\u65b9\u6cd5\u603b\u4f53\u6d41\u7a0b",
        "\u56fe2 \u4e0d\u540c\u8272\u56fe\u5728\u6b63\u5e38\u89c6\u89c9\u4e0e\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u7684\u5dee\u5f02",
        "\u56fe3 \u53c2\u6570\u9009\u62e9\u5bf9\u8f93\u51fa\u6548\u679c\u7684\u5f71\u54cd",
        "\u56fe4 \u6700\u7ec8\u8bc4\u4f30\u96c6 27 \u5e45\u56fe\u50cf\u8f93\u51fa\u603b\u89c8",
        "\u56fe5 X \u5c04\u7ebf\u3001\u710a\u7f1d\u548c\u7eb9\u7406\u56fe\u50cf\u573a\u666f\u5bf9\u6bd4",
        "\u56fe6 \u65b9\u6cd5\u6d88\u878d\u5b9e\u9a8c\u53ef\u89c6\u5316\u7ed3\u679c",
        "\u56fe7 \u8865\u5145\u9a8c\u8bc1\u56fe\u50cf\u8f93\u51fa\u603b\u89c8",
        "\u56fe8 \u590d\u6742\u7eb9\u7406\u548c\u566a\u58f0\u573a\u666f\u4e0b\u7684\u5c40\u9650\u6027\u5206\u6790",
    }
    seen_captions = set()
    fig2_explained = False
    avggrad_seen = False

    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text in final_captions:
            if text in seen_captions:
                paragraphs = body_paragraphs(doc)
                idx = next((i for i, p in enumerate(paragraphs) if p._p is paragraph._p), None)
                if idx is not None and idx > 0 and "w:drawing" in paragraphs[idx - 1]._p.xml:
                    remove_paragraph(paragraphs[idx - 1])
                remove_paragraph(paragraph)
                continue
            seen_captions.add(text)
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                            bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                            first_line=False, before=1, after=0)
            disable_east_asian_auto_spacing(paragraph)
            continue

        if text.startswith("\u7b97\u6cd51 "):
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=10.5,
                            bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                            first_line=False, before=3, after=3)
            disable_east_asian_auto_spacing(paragraph)

        if text.startswith("\u56fe2 \u5bf9\u6bd4\u4e86"):
            if fig2_explained:
                remove_paragraph(paragraph)
                continue
            fig2_explained = True
            set_text_hard(
                paragraph,
                "\u56fe2 \u5bf9\u6bd4\u4e86 Jet\u3001Turbo\u3001Viridis \u548c Cividis "
                "\u5728\u6b63\u5e38\u89c6\u89c9\u3001protan\uff08\u7ea2\u8272\u89c9\u7f3a\u9677\uff09\u548c "
                "deutan\uff08\u7eff\u8272\u89c9\u7f3a\u9677\uff09\u6a21\u62df\u4e0b\u7684\u663e\u793a\u6548\u679c\u3002Jet \u548c Turbo "
                "\u7684\u989c\u8272\u533a\u5206\u5ea6\u8f83\u5f3a\uff0c\u4f46\u90e8\u5206\u533a\u57df\u5728\u8272\u89c9\u969c\u788d\u6a21\u62df\u4e0b\u4ecd\u5bb9\u6613\u63a5\u8fd1\uff1b"
                "Viridis \u548c Cividis \u7684\u8fc7\u6e21\u66f4\u5e73\u7a33\uff0c\u5176\u4e2d Cividis \u66f4\u9002\u5408\u4f5c\u4e3a"
                "\u7ea2\u7eff\u8272\u89c9\u969c\u788d\u53cb\u597d\u7684\u57fa\u7840\u8272\u56fe\u3002"
            )
            continue

        if text.startswith("AvgGrad \u65b9\u9762"):
            if avggrad_seen:
                remove_paragraph(paragraph)
                continue
            avggrad_seen = True
            set_text_hard(
                paragraph,
                "AvgGrad \u65b9\u9762\uff0cDE-Cividis \u7684\u6570\u503c\u9ad8\u4e8e\u56fa\u5b9a Cividis\uff0c"
                "\u8bf4\u660e\u7ec6\u8282\u589e\u5f3a\u540e\u8fb9\u7f18\u548c\u7eb9\u7406\u66f4\u5bb9\u6613\u5f62\u6210\u5c40\u90e8\u53d8\u5316\u3002"
                "\u5bc6\u5ea6\u5206\u5c42\u3001\u4e09\u76f8\u4f4d\u6b63\u5f26\u548c Jet \u4e5f\u53ef\u80fd\u4ea7\u751f\u8f83\u5f3a\u53d8\u5316\uff0c"
                "\u4f46\u540c\u65f6\u4f1a\u5e26\u6765\u66f4\u5927\u8272\u5dee\u6216\u66f4\u4e0d\u7a33\u5b9a\u7684\u989c\u8272\u5c42\u6b21\uff0c"
                "\u6240\u4ee5\u8fd9\u4e00\u9879\u9700\u8981\u4e0e \u0394E2000\u3001SSIM \u548c\u56fe\u50cf\u6548\u679c\u4e00\u8d77\u770b\u3002"
            )


def enforce_caption_styles_final(doc: Document) -> None:
    """Final guard so later text rewrites cannot strip caption fonts."""
    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if is_figure_caption(text):
            compact_caption_number(paragraph)
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                            bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                            first_line=False, before=1, after=0)
            disable_east_asian_auto_spacing(paragraph)
        elif is_table_caption(text) or re.match(r"^\u8868[0-9]+\s", text):
            compact_caption_number(paragraph)
            style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                            bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                            first_line=False, before=3, after=1)
            disable_east_asian_auto_spacing(paragraph)


def reduce_repeated_citation_noise(doc: Document) -> None:
    replacements = {
        "[3,5-7]": "",
    }
    removable_after_first = {"[1-2]", "[3]", "[5-6]", "[8-9]"}
    seen = set()
    citation_re = re.compile(r"\[\d+(?:[-,]\d+)*\]")

    for paragraph in body_paragraphs(doc):
        text = paragraph.text
        if not text.strip() or text.strip().startswith("["):
            continue
        updated = text
        for old, new in replacements.items():
            if old in updated:
                updated = updated.replace(old, new)
        if updated.startswith("\u7ecf\u5178\u4f2a\u5f69\u8272\u65b9\u6cd5\u901a\u5e38"):
            updated = updated.replace("[7]", "")
        if updated.startswith("\u6240\u6709\u5bf9\u6bd4\u65b9\u6cd5\u5747"):
            updated = updated.replace("[8-9]", "")
        for citation in citation_re.findall(updated):
            if citation in removable_after_first and citation in seen:
                updated = updated.replace(citation, "")
            else:
                seen.add(citation)
        if updated != text:
            set_text_hard(paragraph, updated)


def final_reference_strategy_cleanup(doc: Document) -> None:
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()

        if text.startswith("灰度图像是数字图像处理中最基础也最常见的数据形式"):
            set_text_hard(paragraph, text.replace("[1-2]", ""))
            continue

        if text.startswith("伪彩色处理的目的不是单纯增强图像鲜艳程度"):
            set_text_hard(paragraph, text.replace("[3]", ""))
            continue

        if text.startswith("色觉障碍并不是少数到可以忽略的问题"):
            set_text_hard(paragraph, text.replace("[4]", "[1]"))
            continue

        if text.startswith("现有较成熟的色觉障碍友好色图"):
            set_text_hard(paragraph, text.replace("[5-6]", ""))
            continue

        if text.startswith("经典伪彩色方法通常可以分为三类"):
            set_text_hard(
                paragraph,
                "经典伪彩色方法通常可以分为三类。第一类是密度分层，即把灰度区间划分为若干段，"
                "每一段赋予一种固定颜色[2-3]。这种方法易实现、易解释，适合突出特定灰度范围，"
                "但颜色变化不连续，分段边界可能被误认为真实结构。第二类是解析函数映射，例如用三相位正弦函数"
                "分别生成 R、G、B 通道。这种方法连续性较好，但频率、相位等参数依赖人工设定。"
                "第三类是查表色图，如 Jet、Hot、Turbo、Viridis、Cividis 等。"
            )
            continue

        if text.startswith("Jet 色图曾长期作为科学绘图中的事实标准"):
            set_text_hard(
                paragraph,
                "Jet 色图曾长期作为科学绘图中的事实标准，因为它看起来颜色丰富、对比强烈。"
                "但是大量研究指出，Jet 的亮度并不单调，颜色变化也不感知均匀[4]。"
                "在某些灰度区间，Jet 会人为制造很强的视觉边界；在另一些区间，真实的灰度变化又可能被压缩。"
                "这种不稳定性在正常视觉下已经有问题，在红绿色觉障碍视角下更严重。"
            )
            continue

        if text.startswith("Hot 色图相对单调"):
            set_text_hard(
                paragraph,
                "Hot 色图相对单调，结构保真度通常比 Jet 更好，但颜色信息量较低，且在高亮区域容易饱和。"
                "Cividis 在设计时进一步考虑了色觉障碍下的可读性[5]，Viridis 则是常用的感知均匀色图之一[6]。"
                "Turbo 保留了彩虹色图较强的视觉区分度，同时改善了 Jet 的部分跳变问题[7]。"
                "因此，Cividis 被选作基础色图。"
            )
            continue

        if text.startswith("实验使用五类指标评价结果"):
            base = re.sub(r"(?:后文)?表格中，ΔE2000 用 ↓ 标记.*$", "", text).strip()
            base = re.sub(r"表头中的.*$", "", base).strip()
            if "表格中，ΔE2000 用 ↓ 标记" not in base:
                base += (
                    "表格中，ΔE2000 用 ↓ 标记，表示数值越小改动越保守；"
                    "Y-SSIM 和 RGB-SSIM 用 ↑ 标记，表示结构保持越好；"
                    "AvgGrad 和 Entropy 也用 ↑ 标记，分别表示局部变化强度和颜色信息量更高。"
                )
            set_text_hard(paragraph, base)
            continue

        if text.startswith("表头中的"):
            remove_paragraph(paragraph)


def _paragraph_text_from_element(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def _insert_paragraph_before_table(table, text: str) -> Paragraph:
    p = OxmlElement("w:p")
    table._tbl.addprevious(p)
    paragraph = Paragraph(p, table._parent)
    set_text_hard(paragraph, text)
    return paragraph


def ensure_table_captions(doc: Document) -> None:
    captions = [
        "\u88681 \u4e3b\u8bc4\u4f30\u96c6 27 \u5e45\u56fe\u50cf\u7684\u5e73\u5747\u7ed3\u679c\uff08protan/deutan \u5e73\u5747\uff09",
        "\u88682 Kodak \u8865\u5145\u9a8c\u8bc1\u96c6\u7684\u5e73\u5747\u7ed3\u679c\uff08protan/deutan \u5e73\u5747\uff09",
    ]
    for idx, table in enumerate(doc.tables[:2]):
        caption_text = captions[idx]
        previous = table._tbl.getprevious()
        if previous is not None and previous.tag == qn("w:p") and _paragraph_text_from_element(previous).strip().startswith(caption_text[:2]):
            paragraph = Paragraph(previous, table._parent)
            set_text_hard(paragraph, caption_text)
        else:
            paragraph = _insert_paragraph_before_table(table, caption_text)
        style_paragraph(paragraph, east="\u5b8b\u4f53", latin="Times New Roman", size=9,
                        bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                        first_line=False, before=3, after=1)
        disable_east_asian_auto_spacing(paragraph)

    analysis_rewrites = {
        "\u88681 \u6c47\u603b\u4e86 27 \u5e45\u4e3b\u8bc4\u4f30\u56fe\u50cf\u4e0a\u7684\u5e73\u5747\u7ed3\u679c\u3002": (
            "\u4e3b\u8bc4\u4f30\u96c6\u7ed3\u679c\u663e\u793a\uff0c"
        ),
        "\u88682 \u8868\u660e\uff0c\u5728 Kodak \u56fe\u50cf\u4e0a\uff0c": (
            "Kodak \u8865\u5145\u9a8c\u8bc1\u7ed3\u679c\u663e\u793a\uff0c"
        ),
    }
    for paragraph in body_paragraphs(doc):
        text = paragraph.text
        for prefix, replacement in analysis_rewrites.items():
            if text.startswith(prefix):
                set_text_hard(paragraph, replacement + text[len(prefix):])
                break


def final_undergrad_cleanup(doc: Document) -> None:
    """Final paper-facing cleanup after legacy recovery rewrites.

    The recovered script contains several old replacement passes.  This last
    pass keeps the final DOCX aligned with the current method and metrics.
    """

    replacements = [
        (
            "摘要：灰度图像伪彩色处理常用于医学影像",
            "摘要：灰度图像伪彩色处理常用于医学影像、工业检测、地理可视化等场景，但以 Jet 为代表的部分彩虹色图依赖明显颜色差异，"
            "Hot 等色图也可能存在高亮饱和或颜色信息量不足的问题，色觉障碍用户容易丢失关键信息。固定 Cividis 色图具有较好的色觉障碍友好性，"
            "但它只按灰度值进行全局映射，不能主动突出边缘、裂缝和纹理等局部结构。针对这一问题，本文提出 DE-Cividis 方法：先用高斯平滑获得背景层，"
            "再以反锐化方式增强细节层，将增强后的灰度送入 Cividis 色图，最后在 YCbCr 空间进行以原灰度为主的亮度回写。该方法流程较轻量，"
            "参数较少，便于解释。27 幅最终评估图像和 24 幅额外 Kodak 图像上的实验表明，DE-Cividis 在保持较高亮度结构相似度的同时，"
            "有助于提升色觉障碍模拟视角下边缘、裂缝和纹理等局部结构的可见差异，可作为一种轻量的灰度伪彩色辅助显示方案。"
        ),
        (
            "色觉障碍并不是少数到可以忽略的问题",
            "红绿色觉异常在人群中并不少见。公开资料通常认为，约 8% 的男性和 0.4% 的女性存在不同程度的红绿色觉异常[1]。"
            "在医学诊断、工业检测、气象预警和地图阅读等场景中，颜色区分会直接影响信息获取。"
            "因此，伪彩色方法需要兼顾不同视觉条件下的信息可读性。"
        ),
        (
            "一种直观方案是完全不用彩色",
            "灰度显示有利于保持原始结构，但无法利用颜色通道提供额外区分。"
            "实际问题不是在灰图和彩色图之间二选一，而是能否设计一种上色方式，"
            "在不破坏原灰度结构的前提下，让红绿色觉障碍用户也能获得额外的局部结构提示。"
        ),
        (
            "Abstract: Pseudocolor mapping helps reveal structures",
            "Abstract: Pseudocolor mapping helps reveal structures in grayscale images, but rainbow colormaps such as Jet rely on strong color transitions, "
            "and Hot may suffer from saturation or limited color variation. This paper proposes DE-Cividis, a lightweight method that enhances a detail layer, "
            "maps the result with Cividis, and applies luminance write-back dominated by the original grayscale image in YCbCr space. "
            "Experiments on 27 evaluation images and 24 additional Kodak images show that DE-Cividis improves the visibility of local structures under simulated color vision deficiency "
            "while maintaining luminance structure, making it a lightweight auxiliary display method for grayscale pseudocolor visualization."
        ),
        (
            "经典伪彩色方法通常可以分为三类",
            "经典伪彩色增强通常包括密度分层、解析函数映射和查表色图等形式[2-3]。"
            "密度分层把灰度区间划分为若干段，每一段赋予一种固定颜色，适合突出特定灰度范围，"
            "但颜色变化不连续，分段边界也可能被误认为真实结构。解析函数映射则用连续函数把灰度值转换为 R、G、B 通道，"
            "后续实验中的三相位正弦变换即作为这类方法的代表。查表色图通过预设色表完成灰度到颜色的映射，"
            "常见例子包括 Jet、Hot、Turbo、Viridis 和 Cividis。"
        ),
        (
            "Jet 色图曾长期作为科学绘图中的事实标准",
            "Jet 色图曾长期作为科学绘图中的常用色图，因为它颜色丰富、对比强烈。"
            "但已有研究指出，Jet 的亮度变化并不均匀，容易形成伪边界或掩盖真实变化[4]。"
            "这种不稳定性在正常视觉下已经会影响读图，在红绿色觉障碍视角下更容易造成混淆。"
        ),
        (
            "Hot 色图相对单调",
            "Hot 色图相对单调，结构保真度通常比 Jet 更好，但颜色信息量较低，且在高亮区域容易饱和。"
            "Cividis 在设计时进一步考虑了色觉障碍下的可读性[5]，Viridis 则是常用的感知均匀色图之一[6]。"
            "Turbo 保留了彩虹色图较强的视觉区分度，同时改善了 Jet 的部分跳变问题[7]。"
            "综合本文的红绿色觉障碍友好目标和亮度结构保持需求，后续方法以 Cividis 作为基础色图。"
        ),
        (
            "现有较成熟的色觉障碍友好色图包括 Viridis 和 Cividis",
            "现有较成熟的色觉障碍友好色图包括 Viridis 和 Cividis。它们比 Jet 更平滑、更稳定，也较少依赖红绿对比。"
            "尤其 Cividis 在设计时考虑了色觉障碍视角，适合作为灰度伪彩色的安全基底。"
            "然而固定色图只根据灰度值进行一维映射：灰度差相同，颜色差就大致相同。"
            "固定色图不利用空间邻域信息，难以区分同一灰度值处在边缘、裂纹还是背景区域。"
            "因此，在科学可视化任务中，固定 Cividis 虽然稳定，却不一定足够突出局部结构。"
        ),
        (
            "色觉障碍图像处理通常包括模拟和补偿两类任务",
            "色觉障碍图像处理通常包括模拟和补偿两类任务。模拟的目标是估计色觉障碍用户看到的图像，"
            "例如 Brettel、Viénot 和 Mollon 提出的模型在 LMS 锥体响应空间中进行投影，"
            "能够模拟 protan（红色觉缺陷）、deutan（绿色觉缺陷）和 tritan（蓝黄色觉缺陷）等情况[8]；"
            "Machado 等提出的生理模型也常用于色觉障碍模拟[9]。"
        ),
        (
            "Daltonization 是典型的色觉补偿思路",
            "色觉补偿和重着色方法会重新调整已有 RGB 图像的颜色，使色觉障碍用户更容易区分原本混淆的区域。"
            "例如 Kuhn 等提出自然性保持的图像重着色方法，在增强可分辨性的同时尽量保持图像自然性[10]。"
            "这类方法主要面向已有 RGB 彩色图像；本文关注的是灰度图到伪彩色图的生成。"
        ),
        (
            "补偿方法则会重新调整颜色",
            "色觉补偿和重着色方法会重新调整已有 RGB 图像的颜色，使色觉障碍用户更容易区分原本混淆的区域。"
            "例如 Kuhn 等提出自然性保持的图像重着色方法，在增强可分辨性的同时尽量保持图像自然性[10]。"
            "这类方法主要面向已有 RGB 彩色图像；本文关注的是灰度图到伪彩色图的生成。"
        ),
        (
            "也有一些更复杂的图像重着色方法",
            "也有一些更复杂的图像重着色方法[11]，可以处理更一般的彩色图像，但通常需要更多参数或优化过程。"
            "这些研究为色觉障碍友好显示提供了参考；本文的实验对象则限定在灰度强度图像的轻量伪彩色转换。"
        ),
        (
            "设输入灰度图像为 I",
            "设输入灰度图像为 I，大小为 H×W，像素值归一化到 [0, 1]。伪彩色方法的输出记为 C，大小为 H×W×3。"
            "传统查表方法可以写为 C(x, y) = LUT(I(x, y))，其中 LUT 是固定色图。本文关注的问题是："
            "在尽量保持原灰度明暗关系的同时，让红绿色觉障碍模拟视角下的局部结构更容易观察。"
        ),
        (
            "这里所说的关键结构",
            "本文所称局部结构主要指边缘、纹理、强度突变和细节层。"
            "例如，X 射线图像中的缺陷边界、降雨图中的强度变化带、焊缝图中的裂纹或过渡区域，"
            "都属于这类需要优先保留和提示的信息。"
        ),
        (
            "这里的结构主要指",
            "本文所称局部结构主要指边缘、纹理、强度突变和细节层。"
            "例如，X 射线图像中的缺陷边界、降雨图中的强度变化带、焊缝图中的裂纹或过渡区域，"
            "都属于这类需要优先保留和提示的信息。"
        ),
        (
            "这里的局部结构主要包括",
            "本文所称局部结构主要指边缘、纹理、强度突变和细节层。"
            "例如，X 射线图像中的缺陷边界、降雨图中的强度变化带、焊缝图中的裂纹或过渡区域，"
            "都属于这类需要优先保留和提示的信息。"
        ),
        (
            "因此，该方法有三个约束",
            "该方法有三个设计点。第一，采用 Cividis 作为颜色基底，减少对红绿对比的依赖。"
            "第二，在进入色图前先做反锐化细节增强，使边缘和纹理区域获得更大的灰度差。"
            "第三，在 YCbCr 空间进行以原灰度为主的亮度回写，令 Y* = λI + (1 − λ)I′，其中 λ = 0.9。"
            "这样可以保留主要明暗层次，同时留下少量增强后的局部对比。"
        ),
        (
            "DE-Cividis 的第一步是估计背景层",
            "DE-Cividis 的第一步是估计背景层。使用高斯滤波得到 B，即 B = Gσ * I，其中 σ = 2.0。"
            "高斯滤波会抑制高频细节，保留图像的缓慢变化部分，因此 B 可以看作局部背景或低频照明层。"
            "第二步计算细节层 D = I − B。D 中包含边缘、纹理和局部突变，也可能包含少量噪声。"
        ),
        (
            "第四步是亮度回写",
            "第四步是亮度回写。直接把 I′ 映射成 Cividis 会同时改变颜色和亮度。"
            "为控制这种偏离，将初始伪彩色图转换到 YCbCr 空间，并用 Y* = 0.9I + 0.1I′ 替换 Y 通道，"
            "再转换回 RGB。由于颜色空间反变换可能产生少量越界值，最终输出统一裁剪到 [0, 1]。"
        ),
        (
            "输入：归一化灰度图 I",
            "输入：归一化灰度图 I，参数 σ = 2.0，α = 2.0，λ = 0.9"
        ),
        (
            "5. 将 C0 转到 YCbCr 空间",
            "5. 将 C0 转到 YCbCr 空间，用 Y* = λI + (1 − λ)I′ 替换 Y 通道"
        ),
        (
            "实验使用五类指标评价结果",
            "实验使用五类指标评价结果。CIEDE2000 色差（ΔE2000）[14]用于观察输出图与原灰度参考之间的感知差异，"
            "数值越小表示改动越保守；Y-SSIM（亮度通道结构相似度）和 RGB-SSIM（RGB 三通道结构相似度）[15]"
            "用于观察结构保持情况；AvgGrad（平均梯度）用 Sobel 梯度描述色觉障碍模拟后的边缘和纹理强度；"
            "Entropy（信息熵）用于描述输出颜色变化的丰富程度，其中 RGB-SSIM 为 R、G、B 三个通道分别计算 SSIM 后取平均，"
            "Entropy 为输出 RGB 三通道信息熵之和。表中数值按 protan 和 deutan 两种模拟分别统计后取平均；"
            "其中 AvgGrad 使用模拟后的图像计算，ΔE2000、Y-SSIM 和 RGB-SSIM 以原灰度图为参考，Entropy 计算输出图本身。"
            "表头中的 ↓ 表示数值越小越保守，↑ 表示数值越大通常越明显或越接近目标。"
        ),
        (
            "图3 给出不同参数下的输出效果",
            "图3 给出不同参数下的输出效果。σ 较小时，背景层过于接近原图，细节增强不明显；"
            "σ 或 α 过大时，边缘和噪声都更容易被拉强。λ 越接近 1，亮度越接近原灰度图；λ 过低时，"
            "输出明暗关系会更明显地偏离原图。综合多类样例观察，后续实验固定为 σ = 2.0、α = 2.0、λ = 0.9。"
        ),
        (
            "参数选择上，本文在候选阶段",
            "参数选择上，候选阶段对 σ、α 和 λ 做了小范围观察。σ 决定高斯背景层的平滑尺度，"
            "α 决定细节层回加的强度，λ 决定亮度回写中原灰度所占比例。σ 或 α 较小时，弱边缘和细纹理的变化不明显；"
            "取值较大时，噪声和局部过冲更容易出现。λ 太接近 1 时输出较保守，λ 太低时亮度结构偏离更明显。"
            "为保持横向比较一致，候选观察后统一固定方法列表和参数；后续统计均采用 σ = 2.0、α = 2.0、λ = 0.9。"
        ),
        (
            "AvgGrad 方面",
            "AvgGrad 方面，DE-Cividis 高于固定 Cividis，说明细节增强和亮度混合回写后，"
            "色觉障碍模拟视角下的边缘和纹理变化更明显。密度分层、三相位正弦和 Jet 也可能产生较强变化，"
            "但往往伴随更大的色差或更不稳定的颜色层次。"
        ),
        (
            "Entropy 方面",
            "Entropy 方面，DE-Cividis 高于固定 Cividis，说明输出包含更丰富的颜色变化。"
            "这项指标只反映颜色信息量，仍需结合色差、结构相似度和可视化结果一起理解。"
        ),
        (
            "图5 选取了 X 射线",
            "图5 选取了 X 射线、焊缝和纹理图像三类样例。Jet 和 Turbo 的色彩变化更强，"
            "但容易带来较夸张的颜色层次；Viridis 和 Cividis 更平稳，不过弱边缘不一定突出。"
            "DE-Cividis 在 Cividis 的基础上增强细节层，因此裂纹、边界和细纹理位置更容易观察。"
        ),
        (
            "图6 展示了 DE-Cividis 的消融结果",
            "图6 展示了 DE-Cividis 的消融结果。去掉细节增强后，边界位置的颜色差异变弱；"
            "去掉亮度回写后，局部对比虽然更强，但明暗关系更容易偏离原灰度图；改用 Jet 后颜色更鲜艳，"
            "但在色觉障碍模拟下不如 Cividis 稳定。"
        ),
        (
            "因此，本文保留细节增强",
            "因此，方法保留细节增强、Cividis 和亮度回写三个步骤。细节增强负责拉开结构区域，"
            "Cividis 提供相对色觉障碍友好的颜色基底，亮度回写用于约束原始灰度的明暗层次。"
        ),
        (
            "Kodak 补充验证结果显示",
            "Kodak 补充验证结果显示，固定 Cividis 仍然是较保守的基线，DE-Cividis 在保持较高亮度结构相似度的同时，"
            "相对固定 Cividis 提高了 AvgGrad。"
            "由于 Kodak 图像含有更多自然场景纹理，这部分结果主要作为固定参数下的补充观察。"
        ),
        (
            "这组结果与主实验的趋势基本一致",
            "这组结果与主实验的趋势基本一致：DE-Cividis 相比固定 Cividis 提供了更明显的局部结构变化。"
            "所有参数在补充验证前已经固定为 σ = 2.0、α = 2.0、λ = 0.9。"
        ),
        (
            "DE-Cividis 的主要作用不是扩大整幅图像的总体色彩跨度",
            "DE-Cividis 的主要作用是在保持灰度亮度结构的前提下，提高边缘、纹理和局部突变区域的可见差异。"
            "固定 Cividis 的优点是保守稳定，但所有像素只按灰度值查表，无法区分同一灰度值在背景区域和结构区域中的不同作用。"
            "DE-Cividis 通过细节层 D = I − B 对局部变化进行增强，使结构区域在进入 Cividis 前获得更大的灰度差。"
        ),
        (
            "从指标上看，固定 Cividis",
            "从指标上看，固定 Cividis 的 ΔE2000 较低，说明它与原灰度参考更接近。Jet、Hot 和 Turbo 颜色更强，"
            "可能增加局部可见差异，也更容易带来较大色差。DE-Cividis 的位置介于两者之间："
            "它增加少量色差和颜色变化，换取更明显的边缘和纹理提示，同时保持较高的亮度结构相似度。"
        ),
        (
            "DE-Cividis 的处理链较短",
            "DE-Cividis 的处理链较短。它先增强细节，再用 Cividis 上色，最后进行以原灰度为主的亮度回写。"
            "三个步骤分别对应结构、颜色和亮度三个问题，模块关系比较清楚。"
        ),
        (
            "参数方面，σ 决定背景层平滑尺度",
            "参数方面，σ 决定背景层平滑尺度，α 决定细节层回加强度，λ 决定亮度回写中原灰度所占比例。"
            "三者固定为 2.0、2.0 和 0.9，是在多类样例上观察后的折中：取值过小，弱边缘变化不明显；"
            "增强或亮度混合过强，则噪声、局部过冲和亮度偏移更容易出现。"
        ),
        (
            "如果输入本身是普通 RGB 彩色照片",
            "对于普通 RGB 彩色照片，颜色本身往往已经带有语义信息，通常更适合采用面向彩色图像的重着色或 Daltonization 方法。"
        ),
        (
            "参数固定后，程序使用同一组",
            "参数固定后，程序使用同一组 σ = 2.0、α = 2.0、λ = 0.9 处理所有最终评估图像和额外 Kodak 图像。"
            "表格中的指标和正文中的图像展示都来自这组固定参数。"
        ),
        (
            "亮度回写是实现中的关键步骤",
            "亮度回写是实现中的关键步骤。若直接输出 Cividis(I′)，图像亮度会同时受到细节增强和色图亮度变化影响，"
            "可能偏离原始灰度结构。将初始伪彩色结果转换到 YCbCr 空间，并用 Y* = 0.9I + 0.1I′ 替换 Y 通道，"
            "可以让亮度结构主要受原图约束，同时保留少量增强后的局部对比。转回 RGB 后，所有通道统一裁剪到 [0, 1]。"
        ),
        (
            "围绕红绿色觉障碍用户的灰度图像伪彩色可读性问题",
            "围绕红绿色觉障碍用户的灰度图像伪彩色可读性问题，提出了 DE-Cividis 方法。"
            "该方法通过高斯背景层估计和反锐化细节增强，使边缘、纹理和局部突变区域在进入色图前获得更大的灰度差；"
            "随后使用 Cividis 提供相对色觉障碍友好的颜色基底；最后在 YCbCr 空间进行以原灰度为主的亮度回写，以保持主要明暗结构。"
            "整体流程较简单，参数较少，便于解释。"
        ),
        (
            "实验结果表明，DE-Cividis",
            "实验结果表明，DE-Cividis 在亮度结构保持和边缘纹理提示之间取得了较好的折中。"
            "与密度分层、三相位正弦变换以及 Jet、Hot、Turbo 等色图相比，它的颜色变化更克制；"
            "与 Viridis、Cividis 相比，它对边缘和纹理的提示更明显。在本实验测试的图像范围内，"
            "DE-Cividis 可作为一种轻量的灰度伪彩色显示方案。"
        ),
    ]

    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if not text:
            continue
        if text.startswith("表头中的"):
            remove_paragraph(paragraph)
            continue
        if "KSCgain" in text or "EdgeGain" in text:
            remove_paragraph(paragraph)
            continue
        for prefix, replacement in replacements:
            if text.startswith(prefix):
                set_text_hard(paragraph, replacement)
                break

    paragraphs = body_paragraphs(doc)
    start = None
    end = None
    for i, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if text == "3.3 候选方法比较":
            start = i
        elif start is not None and text.startswith("4 实验设计"):
            end = i
            break
    if start is not None and end is not None:
        for paragraph in paragraphs[start:end]:
            remove_paragraph(paragraph)

    for paragraph in body_paragraphs(doc):
        text = paragraph.text.strip()
        if text.startswith("表1 "):
            set_text_hard(paragraph, "表1 主评估集 27 幅图像的平均结果（protan/deutan 平均）")
        elif text.startswith("表2 "):
            set_text_hard(paragraph, "表2 Kodak 补充验证集的平均结果（protan/deutan 平均）")


def final_submission_cleanup(doc: Document) -> None:
    """Last pass for the submission draft.

    Keep this pass small and paper-facing: it only fixes wording/structure that
    must survive all legacy recovery rewrites above.
    """

    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        has_page_break = any(br.get(qn("w:type")) == "page" for br in paragraph._p.iter(qn("w:br")))
        if not text and has_page_break:
            remove_paragraph(paragraph)

    all_text = "\n".join(p.text.strip() for p in body_paragraphs(doc))

    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if text.startswith("综合固定色图的稳定性和结构增强需求后"):
            set_text_hard(
                paragraph,
                "综合固定色图的稳定性和结构增强需求后，本文采用 DE-Cividis，即 Detail-Enhanced Cividis。"
                "该方法不重新设计色图，而是在进入 Cividis 前增强灰度细节层，再通过亮度回写保持原灰度明暗结构。"
                "后文将 DE-Cividis 与经典伪彩色方法和公开色图进行比较。"
            )
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if text.startswith("多尺度反锐化可以把不同尺度"):
            set_text_hard(
                paragraph,
                "多尺度反锐化可以把不同尺度的细节层加权融合，用来同时增强细边缘和中尺度纹理。"
                "本文采用单尺度反锐化作为实现，主要考虑参数数量、结果稳定性和方法解释的清晰度。"
            )
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if text.startswith("细节增强是图像处理中的经典主题"):
            set_text_hard(
                paragraph,
                "细节增强是图像处理中的经典主题。反锐化掩模的思想很直接：先用低通滤波得到平滑背景，"
                "再用原图减去背景得到细节层，最后把细节层加回原图。它的优点是计算简单、参数少、效果稳定；"
                "缺点是增强过强时可能放大噪声或产生过冲。本文将反锐化作为 DE-Cividis 的结构增强模块。"
            )
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if text.startswith("图1 ") and "总体流程" in text and "亮度回写后的结果比直接 Cividis 映射" not in all_text:
            added = add_paragraph_after(
                paragraph,
                "图1中，亮度回写后的结果比直接 Cividis 映射更接近原灰度明暗关系。"
                "亮度回写用于保留少量细节增强效果，同时减小色图亮度变化对原始明暗关系的影响。",
            )
            style_paragraph(added, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if text.startswith("图1中，亮度回写后的结果比直接 Cividis 映射"):
            set_text_hard(
                paragraph,
                "图1中，亮度回写后的结果比直接 Cividis 映射更接近原灰度明暗关系。"
                "亮度回写用于保留少量细节增强效果，同时减小色图亮度变化对原始明暗关系的影响。"
            )
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if (
            text.startswith("经典伪彩色增强通常包括")
            or text.startswith("经典伪彩色方法通常可以分为三类")
            or text.startswith("常见灰度伪彩色方法可以从")
        ):
            set_text_hard(
                paragraph,
                "常见灰度伪彩色方法可以从“灰度到颜色的映射关系”来理解。"
                "其中，密度分层把灰度区间划分为若干段，并为每一段赋予固定颜色[2-3]。"
                "这种方法易实现、易解释，适合突出特定灰度范围，但颜色变化不连续，分段边界也可能被误认为真实结构。"
                "另一类方法直接用函数关系生成 R、G、B 三个通道，例如通过相位不同的正弦函数把灰度变化转成连续的颜色变化。"
                "这类方法的过渡较平滑，但频率、相位等参数会影响显示效果。"
                "查表色图则通过预设色表完成灰度到颜色的映射，常见例子包括 Jet、Hot、Turbo、Viridis 和 Cividis。"
            )
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if text == "因此，本文将 DE-Cividis 与经典伪彩色方法及常用公开色图进行比较。":
            set_text_hard(paragraph, "因此，DE-Cividis 将与经典伪彩色方法及常用公开色图进行比较。")
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

        if text.startswith("图1 展示了 DE-Cividis"):
            set_text_hard(
                paragraph,
                text + " 其中亮度回写用于保留少量细节增强效果，同时减小色图亮度变化对原始明暗关系的影响。"
            )
            style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)
            continue

    seen_fig1_note = False
    for paragraph in list(body_paragraphs(doc)):
        text = paragraph.text.strip()
        if "图1中，亮度回写后的结果比直接 Cividis 映射" in text:
            if seen_fig1_note:
                remove_paragraph(paragraph)
            else:
                seen_fig1_note = True

    paragraphs = body_paragraphs(doc)
    start = None
    end = None
    for i, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if text in {"6 分析与讨论", "6 讨论与总结", "6 结论"}:
            start = i
        elif start is not None and text == "参考文献":
            end = i
            break

    if start is None or end is None:
        return

    chapter_heading = paragraphs[start]
    set_text_hard(chapter_heading, "6 结论")
    chapter_heading.style = doc.styles["Heading 1"]
    style_paragraph(chapter_heading, east="\u9ed1\u4f53", size=14, bold=True,
                    align=WD_ALIGN_PARAGRAPH.LEFT, first_line=False, before=8, after=4)
    for paragraph in paragraphs[start + 1:end]:
        remove_paragraph(paragraph)

    new_items = [
        (
            "针对灰度图像在红绿色觉障碍视角下不易区分局部结构的问题，本文设计了 DE-Cividis 伪彩色方法。"
            "该方法先用高斯滤波估计背景层，再通过反锐化增强边缘、纹理和局部突变；随后使用 Cividis 完成伪彩色映射，"
            "并在 YCbCr 空间进行亮度回写，使输出图尽量保留原灰度图的主要明暗关系。整个流程只包含细节增强、色图映射和亮度约束三个步骤，"
            "实现较简单，参数也比较少。",
            "body",
        ),
        (
            "在 27 幅主评估图像和 Kodak 补充图像上的实验表明，DE-Cividis 相比固定 Cividis 能提供更明显的边缘和纹理提示，"
            "同时没有像 Jet、Turbo 等色图那样产生过强的颜色跳变。结合可视化结果和各项指标来看，"
            "该方法在局部结构增强和原灰度明暗保持之间取得了较好的折中，适合作为一种轻量的灰度伪彩色显示方案。",
            "body",
        ),
        (
            "从适用范围看，DE-Cividis 更适合医学灰度图、工业检测图、遥感或气象强度图等单通道强度图的辅助显示。"
            "对于普通 RGB 彩色照片，原有颜色已经包含较多语义信息，更适合采用面向彩色图像的重着色或 Daltonization 方法。",
            "body",
        ),
        (
            "本文仍有不足。实验主要基于色觉障碍模拟和客观指标，没有进行真实色觉障碍用户的任务测试。"
            "此外，参数 σ = 2.0、α = 2.0、λ = 0.9 是在多类样例上观察得到的折中取值，面对噪声较强或纹理特别复杂的图像时，"
            "仍可能需要调整增强强度。",
            "body",
        ),
        (
            "后续工作可以从两个方向继续改进：一是加入简单的去噪或自适应参数选择，减少噪声被同步增强的问题；"
            "二是设计边界识别、异常定位等小任务，邀请真实用户参与评价，使方法效果不只停留在模拟结果和数值指标上。",
            "body",
        ),
    ]

    for text, kind in reversed(new_items):
        paragraph = add_paragraph_after(chapter_heading, text)
        style_paragraph(paragraph, east="\u5b8b\u4f53", size=10.5, bold=False,
                        align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=True, before=0, after=0)


def replace_embedded_images(docx_path: Path) -> None:
    figure_order = [
        ROOT / "figs" / "fig1_pipeline.png",
        ROOT / "figs" / "fig2_colormap_cvd.png",
        ROOT / "figs" / "fig3_saliency.png",
        ROOT / "figs" / "fig4_eval_overview.png",
        ROOT / "figs" / "fig5_typical_cases.png",
        ROOT / "figs" / "fig6_ablation.png",
        ROOT / "figs" / "fig7_kodak_overview.png",
    ]
    missing = [str(path) for path in figure_order if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing figure files: " + "; ".join(missing))

    tmp = docx_path.with_name(docx_path.stem + "_tmp_image_replace.docx")
    with zipfile.ZipFile(docx_path, "r") as zin:
        document_xml = zin.read("word/document.xml").decode("utf-8")
        rels_xml = zin.read("word/_rels/document.xml.rels").decode("utf-8")
        embeds = re.findall(r'r:embed="(rId\d+)"', document_xml)
        rel_targets = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="media/([^"]+)"', rels_xml))
        image_map = {}
        for figure_path, rid in zip(figure_order, embeds):
            target = rel_targets.get(rid)
            if target:
                image_map[f"word/media/{target}"] = figure_path

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in image_map:
                    zout.writestr(item, image_map[item.filename].read_bytes())
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(str(tmp), str(docx_path))


def main() -> None:
    if not DOCX.exists():
        raise FileNotFoundError(DOCX)

    doc = Document(DOCX)
    fix_figure_sequence(doc)
    finalize_figure_xml(doc)
    normalize_section_titles(doc)
    polish_wording(doc)
    split_fig2_fig3_explanations(doc)
    add_limitations_intro(doc)
    revise_current_draft_text(doc)
    add_course_knowledge_sentence(doc)
    split_fig2_fig3_explanations(doc)
    insert_evaluation_overview(doc)
    insert_kodak_overview(doc)
    add_figure_overview_explanations(doc)
    apply_final_baseline_rewrite(doc)
    apply_no_ksc_final_rewrite(doc)
    remove_adjacent_duplicate_paragraphs(doc)
    replace_final_figure_captions(doc)
    ensure_section_headings(doc)
    apply_page_and_styles(doc)
    polish_algorithm_block(doc)
    update_metric_tables_v2(doc)
    polish_tables(doc)
    finalize_metric_table_labels(doc)
    add_page_number_footer(doc)
    resize_images(doc)
    final_caption_and_duplicate_cleanup(doc)
    ensure_table_captions(doc)
    reduce_repeated_citation_noise(doc)
    final_reference_strategy_cleanup(doc)
    final_undergrad_cleanup(doc)
    final_submission_cleanup(doc)
    merge_discussion_and_conclusion(doc)
    move_kodak_overview_to_discussion(doc)
    keep_52_heading_with_figure(doc)
    ensure_fig3_explanation(doc)
    add_space_after_tables(doc)
    resize_images(doc)
    polish_references(doc)
    apply_reference_crosslinks(doc)
    enforce_caption_styles_final(doc)
    trim_trailing_empty_paragraphs(doc)
    doc.save(DOCX)
    replace_embedded_images(DOCX)
    print(f"saved: {DOCX}")


if __name__ == "__main__":
    main()
