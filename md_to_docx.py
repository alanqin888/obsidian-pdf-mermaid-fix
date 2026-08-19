#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import io
import json
import re
import urllib.request
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


PRIMARY = "0C1A32"
ACCENT = "FF5722"
LIGHT_BG = "F6F8FB"
TABLE_HEADER_BG = "0C1A32"
TABLE_BORDER = "D8DEE9"
BODY_FONT = "PingFang SC"
WEST_FONT = "Arial"


def set_run_font(run, size: int | float | None = None, bold: bool | None = None, color: str | None = None):
    run.font.name = WEST_FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def set_paragraph_font(paragraph, size: int | float = 10.5, color: str = "1F2937"):
    for run in paragraph.runs:
        set_run_font(run, size=size, color=color)


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=120, start=120, bottom=120, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), TABLE_BORDER)


def clear_cell(cell):
    for paragraph in cell.paragraphs:
        paragraph.clear()


def clean_inline(text: str) -> str:
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\{width=[^}]+\}", "", text)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()


def parse_inline_tokens(text: str) -> list[tuple[str, bool, bool]]:
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(p|div|span)[^>]*>", "", text, flags=re.IGNORECASE)
    text = text.replace("<strong>", "**").replace("</strong>", "**")
    text = re.sub(r"\{width=[^}]+\}", "", text)
    
    pattern = r"(\*\*.+?\*\*|`[^`]+`)"
    tokens = re.split(pattern, text)
    result = []
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            result.append((token[2:-2], True, False))
        elif token.startswith("`") and token.endswith("`") and len(token) >= 2:
            result.append((token[1:-1], False, True))
        else:
            result.append((token, False, False))
    return result


def add_formatted_text(paragraph, text: str, default_size=10.5, default_color="1F2937", force_bold=False):
    tokens = parse_inline_tokens(text)
    for content, is_bold, is_code in tokens:
        lines = content.split("\n")
        for i, line_text in enumerate(lines):
            if i > 0:
                paragraph.add_run().add_break()
            if not line_text:
                continue
            run = paragraph.add_run(line_text)
            if is_code:
                run.font.name = "Courier New"
                set_run_font(run, size=default_size - 0.5, bold=force_bold, color="334155")
            else:
                set_run_font(run, size=default_size, bold=is_bold or force_bold, color=default_color)


def write_cell(cell, text: str, header: bool = False, zebra: bool = False):
    clear_cell(cell)
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    add_formatted_text(paragraph, text, default_size=9.5, default_color="FFFFFF" if header else "1F2937", force_bold=header)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell, top=120, start=140, bottom=120, end=140)
    
    if header:
        set_cell_shading(cell, TABLE_HEADER_BG)
    elif zebra:
        set_cell_shading(cell, "F8FAFC")
    else:
        set_cell_shading(cell, "FFFFFF")


def add_paragraph_with_inline(paragraph, text: str, size=10.5, color="1F2937", bold=False):
    add_formatted_text(paragraph, text, default_size=size, default_color=color, force_bold=bold)


def set_paragraph_callout_style(paragraph, fill_color="F8FAFC", border_color="0C1A32"):
    pPr = paragraph._p.get_or_add_pPr()
    
    # 背景着色
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    pPr.append(shd)
    
    # 左侧加粗边框
    pBdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '24') # 3pt 边框宽度
    left.set(qn('w:space'), '12')
    left.set(qn('w:color'), border_color)
    pBdr.append(left)
    pPr.append(pBdr)


def add_quote_block(doc: Document, quote_lines: list[str], is_cover: bool = False):
    if not quote_lines:
        return
    
    # 检查是否全部或大部分为 Key-Value 格式 (例如 "**文档定位**：xxx" 或 "**适用周期**：xxx")
    kv_pairs = []
    is_kv = True
    for line in quote_lines:
        m = re.match(r"^\*{0,2}([\u4e00-\u9fa5a-zA-Z0-9_（）\s]{2,12})\*{0,2}[：:]\s*(.+)$", line.strip())
        if m:
            key = clean_inline(m.group(1))
            val = m.group(2).strip()
            kv_pairs.append((key, val))
        else:
            if kv_pairs and (line.strip().startswith("1.") or line.strip().startswith("2.") or line.strip().startswith("-") or line.strip().startswith("•")):
                last_k, last_v = kv_pairs[-1]
                kv_pairs[-1] = (last_k, last_v + "\n" + line.strip())
            else:
                is_kv = False
                break
                
    if is_kv and len(kv_pairs) >= 1:
        # 如果是封面元数据表，前面加入优雅的下沉间距
        if is_cover:
            spacer = doc.add_paragraph()
            set_paragraph_spacing(spacer, before=36, after=0)
            
        table = doc.add_table(rows=len(kv_pairs), cols=2)
        table.autofit = False
        set_table_borders(table)
        
        for row_idx, (k, v) in enumerate(kv_pairs):
            row = table.rows[row_idx]
            
            # 左单元格
            c0 = row.cells[0]
            c0.width = Cm(3.2)
            clear_cell(c0)
            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run0 = p0.add_run(k)
            set_run_font(run0, size=10, bold=True, color=PRIMARY)
            c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(c0, top=120, start=120, bottom=120, end=120)
            set_cell_shading(c0, "F1F5F9")
            
            # 右单元格
            c1 = row.cells[1]
            c1.width = Cm(12.6)
            clear_cell(c1)
            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
            add_formatted_text(p1, v, default_size=10, default_color="1F2937")
            c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(c1, top=120, start=140, bottom=120, end=140)
            set_cell_shading(c1, "FFFFFF")
            
        if is_cover:
            # 封面结束，自动插入标准分页符，进入正文第 2 页！
            doc.add_page_break()
        else:
            p_sp = doc.add_paragraph()
            set_paragraph_spacing(p_sp, after=10)
    else:
        # 自由文本摘要：生成单格精致摘要边框表格
        if is_cover:
            spacer = doc.add_paragraph()
            set_paragraph_spacing(spacer, before=36, after=0)
            
        table = doc.add_table(rows=1, cols=1)
        table.autofit = True
        set_table_borders(table)
        cell = table.rows[0].cells[0]
        clear_cell(cell)
        set_cell_margins(cell, top=140, start=160, bottom=140, end=160)
        set_cell_shading(cell, "F8FAFC")
        
        for idx, line in enumerate(quote_lines):
            p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
            set_paragraph_spacing(p, before=2 if idx > 0 else 0, after=2, line=1.35)
            add_formatted_text(p, line, default_size=10, default_color="1F2937")
            
        if is_cover:
            doc.add_page_break()
        else:
            p_sp = doc.add_paragraph()
            set_paragraph_spacing(p_sp, after=10)


def configure_styles(doc: Document):
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    normal = doc.styles["Normal"]
    normal.font.name = WEST_FONT
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string("1F2937")

    for name in ("Heading 1", "Heading 2", "Heading 3"):
        style = doc.styles[name]
        style.font.name = WEST_FONT
        style._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
        style.font.color.rgb = RGBColor.from_string(PRIMARY)
        style.font.bold = True

    doc.styles["Heading 1"].font.size = Pt(18)
    doc.styles["Heading 2"].font.size = Pt(14)
    doc.styles["Heading 3"].font.size = Pt(12)


def set_paragraph_spacing(paragraph, before=0, after=6, line=1.35):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line


def add_heading(doc: Document, text: str, level: int, is_cover: bool = False):
    clean_title = clean_inline(text)
    
    if level == 1:
        # 大标题 (H1) 居中排版；作为封面大标题时下沉并加大字号
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(paragraph, before=60 if is_cover else 10, after=36 if is_cover else 16, line=1.2)
        run = paragraph.add_run(clean_title)
        set_run_font(run, size=24 if is_cover else 22, bold=True, color=PRIMARY)
        return

    # 二级与三级标题（标准规范/论文排版：简洁、大气、层次分明）
    paragraph = doc.add_heading(clean_title, level=min(level, 3))
    
    if level == 2:
        # 二级标题 (H2)
        set_paragraph_spacing(paragraph, before=18, after=6, line=1.25)
        for run in paragraph.runs:
            set_run_font(run, size=15, bold=True, color=PRIMARY)
    else:
        # 三级标题 (H3)
        set_paragraph_spacing(paragraph, before=12, after=4, line=1.2)
        for run in paragraph.runs:
            set_run_font(run, size=12.5, bold=True, color="1E293B")


def add_body_paragraph(doc: Document, text: str):
    paragraph = doc.add_paragraph()
    set_paragraph_spacing(paragraph, after=9, line=1.55)
    add_paragraph_with_inline(paragraph, text, size=11)


def add_bullet(doc: Document, text: str):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.left_indent = Cm(0.55)
    paragraph.paragraph_format.first_line_indent = Cm(-0.28)
    set_paragraph_spacing(paragraph, after=6, line=1.5)
    bullet = paragraph.add_run("• ")
    set_run_font(bullet, size=11, bold=False, color=ACCENT)
    add_formatted_text(paragraph, text, default_size=11, default_color="1F2937")


def add_numbered(doc: Document, number: str, text: str):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.left_indent = Cm(0.55)
    paragraph.paragraph_format.first_line_indent = Cm(-0.4)
    set_paragraph_spacing(paragraph, after=7, line=1.52)
    prefix = paragraph.add_run(f"{number}. ")
    set_run_font(prefix, size=11, bold=True, color=PRIMARY)
    add_formatted_text(paragraph, text, default_size=11, default_color="1F2937")


def get_image_size(image_bytes: bytes) -> tuple[int, int]:
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(image_bytes))
        return img.size
    except Exception:
        pass

    # 纯 Python 解析 PNG 头部 IHDR (offset 16..24)
    if image_bytes.startswith(b"\x89PNG\r\n\x1a\n") and len(image_bytes) >= 24:
        w = int.from_bytes(image_bytes[16:20], "big")
        h = int.from_bytes(image_bytes[20:24], "big")
        return w, h

    return 800, 600


def add_scaled_picture(run, image_bytes: bytes, max_width_cm: float = 15.8, max_height_cm: float = 18.0):
    w, h = get_image_size(image_bytes)
    if w <= 0 or h <= 0:
        w, h = 800, 600

    aspect_ratio = h / w

    # 默认按标准页宽 (15.8cm) 计算高度
    target_width_cm = max_width_cm
    target_height_cm = target_width_cm * aspect_ratio

    # 如果高度超过 A4 单页最大容纳高度 (18.0cm)，以最大高度按比例约束宽度
    if target_height_cm > max_height_cm:
        target_height_cm = max_height_cm
        target_width_cm = target_height_cm / aspect_ratio

    image_stream = io.BytesIO(image_bytes)
    run.add_picture(image_stream, width=Cm(target_width_cm), height=Cm(target_height_cm))


def parse_table(lines: list[str], start: int) -> tuple[list[list[str]], int]:
    rows = []
    index = start
    while index < len(lines):
        line = lines[index].strip()
        if not line.startswith("|") or not line.endswith("|"):
            break
        cells = [cell.strip() for cell in line.strip("|").split("|")]
        if all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells):
            index += 1
            continue
        rows.append(cells)
        index += 1
    return rows, index


def add_table(doc: Document, rows: list[list[str]]):
    if not rows:
        return
    width = max(len(row) for row in rows)
    table = doc.add_table(rows=len(rows), cols=width)
    table.autofit = True
    set_table_borders(table)
    
    for row_index, row in enumerate(rows):
        is_header = (row_index == 0)
        is_zebra = (row_index % 2 == 1 and not is_header)
        for col_index in range(width):
            text = row[col_index] if col_index < len(row) else ""
            write_cell(table.rows[row_index].cells[col_index], text, header=is_header, zebra=is_zebra)
            
    paragraph = doc.add_paragraph()
    set_paragraph_spacing(paragraph, after=10)


def add_image(doc: Document, md_path: Path, line: str):
    match = re.match(r"!\[(?P<alt>[^\]]*)\]\((?P<path>[^)]+)\)(?P<attrs>\{[^}]+\})?", line.strip())
    if not match:
        add_body_paragraph(doc, line)
        return
    alt = match.group("alt").strip()
    image_path = (md_path.parent / match.group("path")).resolve()
    if not image_path.exists():
        add_body_paragraph(doc, f"[图片缺失] {alt}：{image_path}")
        return
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(paragraph, before=8, after=4)
    run = paragraph.add_run()
    
    with open(image_path, "rb") as f:
        img_bytes = f.read()
    add_scaled_picture(run, img_bytes, max_width_cm=15.8, max_height_cm=18.0)
    
    if alt:
        caption = doc.add_paragraph()
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(caption, after=10)
        caption_run = caption.add_run(alt)
        set_run_font(caption_run, size=9, color="64748B")


def sanitize_mermaid_code(code: str) -> str:
    lines = code.splitlines()
    sanitized = []
    is_sequence = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("sequenceDiagram"):
            is_sequence = True
            sanitized.append(line)
            continue
        if is_sequence:
            m_actor = re.match(r"^(\s*actor\s+)([^\"'\s]+)(.*)$", line)
            if m_actor:
                line = f'{m_actor.group(1)}"{m_actor.group(2)}"{m_actor.group(3)}'
            m_part = re.match(r"^(\s*participant\s+[^\"'\s]+\s+as\s+)([^\"'].*)$", line)
            if m_part and not m_part.group(2).startswith('"'):
                line = f'{m_part.group(1)}"{m_part.group(2).strip()}"'
        sanitized.append(line)
    return "\n".join(sanitized)


import json

def add_mermaid_image(doc: Document, mermaid_code: str):
    png_data = None
    svg_data = None

    # 1. 抓取 SVG 矢量数据与 PNG 高清数据
    try:
        graph_dict = {"code": mermaid_code.strip(), "mermaid": {"theme": "default"}}
        base64_str = base64.b64encode(json.dumps(graph_dict).encode("utf-8")).decode("utf-8")
        
        svg_url = f"https://mermaid.ink/svg/{base64_str}"
        png_url = f"https://mermaid.ink/img/{base64_str}"

        req_svg = urllib.request.Request(svg_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req_svg, timeout=15) as resp:
            svg_data = resp.read()

        req_png = urllib.request.Request(png_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req_png, timeout=15) as resp:
            png_data = resp.read()
    except Exception as e:
        print(f"mermaid.ink fetch failed: {e}")

    # 备用 kroki
    if not svg_data or not png_data:
        try:
            req_kroki = urllib.request.Request(
                "https://kroki.io/mermaid/png",
                data=mermaid_code.strip().encode("utf-8"),
                headers={"Content-Type": "text/plain; charset=utf-8", "User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req_kroki, timeout=15) as resp:
                png_data = resp.read()
        except Exception as e:
            print(f"kroki fetch failed: {e}")

    # 绘制高清晰度流程图节点 (高度限制在 18.0cm 以内，绝对不会超过单页被截断)
    if png_data:
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(paragraph, before=10, after=10)
        run = paragraph.add_run()
        add_scaled_picture(run, png_data, max_width_cm=15.8, max_height_cm=18.0)

        # 附加矢量 SVG 结构
        if svg_data:
            try:
                from docx.opc.constants import RELATIONSHIP_TYPE
                from docx.opc.packuri import PackURI
                from docx.opc.part import Part
                from docx.oxml import parse_xml

                part_count = len(doc.part.package.parts)
                svg_pack_uri = PackURI(f"/word/media/image_svg_{part_count}.svg")
                svg_part = Part(svg_pack_uri, "image/svg+xml", svg_data, doc.part.package)
                svg_rId = doc.part.relate_to(svg_part, RELATIONSHIP_TYPE.IMAGE)

                blip = run._r.xpath(".//a:blip")[0]
                svg_blip_xml = f'<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{svg_rId}"/>'
                blip.append(parse_xml(svg_blip_xml))
            except Exception as svg_err:
                print(f"Failed to attach SVG blip: {svg_err}")
    else:
        add_quote(doc, f"[Mermaid 流程图代码]:\n{mermaid_code}")


def add_code_block(doc: Document, code_text: str):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.left_indent = Cm(0.5)
    paragraph.paragraph_format.right_indent = Cm(0.5)
    set_paragraph_spacing(paragraph, before=6, after=8, line=1.15)
    run = paragraph.add_run(code_text)
    set_run_font(run, size=9.5, color="334155")
    run.font.name = "Courier New"


def add_footer(doc: Document):
    footer = doc.sections[0].footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("爻鉴 · 数字内容资产整理工具链说明")
    set_run_font(run, size=8, color="64748B")


def convert(md_path: Path, output_path: Path):
    doc = Document()
    configure_styles(doc)
    add_footer(doc)

    lines = md_path.read_text(encoding="utf-8").splitlines()
    
    # 检测前部是否存在封面元数据结构 (H1 紧跟元数据属性，无论是否带 > 引用前缀)
    has_cover_structure = False
    first_h1_idx = -1
    first_h2_idx = 999999
    cover_meta_lines = []
    cover_meta_indices = set()
    
    for idx, raw in enumerate(lines):
        l = raw.strip()
        if not l:
            continue
        if re.match(r"^#\s+", l) and first_h1_idx == -1:
            first_h1_idx = idx
            continue
        if re.match(r"^##\s+", l):
            first_h2_idx = idx
            break
        if first_h1_idx != -1 and idx < first_h2_idx:
            if re.fullmatch(r"-{3,}", l):
                cover_meta_indices.add(idx)
                continue
            clean_l = l[2:].strip() if l.startswith("> ") else l
            m = re.match(r"^\*{0,2}([\u4e00-\u9fa5a-zA-Z0-9_（）\s]{2,12})\*{0,2}[：:]\s*(.+)$", clean_l)
            if m:
                cover_meta_lines.append(clean_l)
                cover_meta_indices.add(idx)
            else:
                break
                
    if first_h1_idx != -1 and len(cover_meta_lines) >= 1:
        has_cover_structure = True

    index = 0
    h1_rendered = False
    cover_table_rendered = False

    while index < len(lines):
        # 如果当前行属于封面元数据区域且不是 H1
        if has_cover_structure and index in cover_meta_indices:
            if not cover_table_rendered:
                add_quote_block(doc, cover_meta_lines, is_cover=True)
                cover_table_rendered = True
            index += 1
            continue

        raw = lines[index]
        line = raw.strip()
        if not line:
            index += 1
            continue

        if re.fullmatch(r"-{3,}", line):
            index += 1
            continue

        if line.startswith("|") and line.endswith("|"):
            rows, index = parse_table(lines, index)
            add_table(doc, rows)
            continue

        heading = re.match(r"^(#{1,6})\s+(.+)$", line)
        if heading:
            level = len(heading.group(1))
            text = heading.group(2)
            is_cover_h1 = (level == 1 and has_cover_structure and not h1_rendered)
            add_heading(doc, text, level, is_cover=is_cover_h1)
            if level == 1:
                h1_rendered = True
            index += 1
            continue

        if line.startswith("!["):
            add_image(doc, md_path, line)
            index += 1
            continue

        bullet = re.match(r"^[-*]\s+(.+)$", line)
        if bullet:
            add_bullet(doc, bullet.group(1))
            index += 1
            continue

        numbered = re.match(r"^(\d+)\.\s+(.+)$", line)
        if numbered:
            add_numbered(doc, numbered.group(1), numbered.group(2))
            index += 1
            continue

        if line.startswith("> "):
            quote_lines = []
            while index < len(lines) and lines[index].strip().startswith("> "):
                quote_lines.append(lines[index].strip()[2:].strip())
                index += 1
            add_quote_block(doc, quote_lines, is_cover=False)
            continue

        if line.startswith("```"):
            if line.startswith("```mermaid"):
                mermaid_lines = []
                index += 1
                while index < len(lines) and not lines[index].strip().startswith("```"):
                    mermaid_lines.append(lines[index])
                    index += 1
                if index < len(lines) and lines[index].strip().startswith("```"):
                    index += 1
                add_mermaid_image(doc, "\n".join(mermaid_lines))
                continue
            else:
                code_lines = []
                index += 1
                while index < len(lines) and not lines[index].strip().startswith("```"):
                    code_lines.append(lines[index])
                    index += 1
                if index < len(lines) and lines[index].strip().startswith("```"):
                    index += 1
                add_code_block(doc, "\n".join(code_lines))
                continue

        html_heading = re.match(r"^<h([1-6])>(.+?)</h\1>$", line)
        if html_heading:
            level = int(html_heading.group(1))
            text = html_heading.group(2)
            add_heading(doc, text, level)
            index += 1
            continue

        # 剥离简单的 <p> 标签包围
        if line.startswith("<p>") and line.endswith("</p>"):
            line = line[3:-4]
            
        # 如果是 HTML 注释，直接跳过
        if line.startswith("<!--") and line.endswith("-->"):
            index += 1
            continue

        add_body_paragraph(doc, line)
        index += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Convert project Markdown to a structured Word document.")
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    args = parser.parse_args()
    convert(args.input.resolve(), args.output.resolve())


if __name__ == "__main__":
    main()
