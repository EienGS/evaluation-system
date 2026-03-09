"""
将 .docx 文件转换为 Markdown 文本
依赖：python-docx
"""
from docx import Document
from docx.oxml.ns import qn
import io


def _is_list_paragraph(para):
    return para.style.name.startswith("List") or para._p.find(qn("w:numPr")) is not None


def _get_heading_level(para):
    name = para.style.name
    if name.startswith("Heading "):
        try:
            return int(name.split(" ")[1])
        except ValueError:
            pass
    return 0


def _table_to_md(table):
    rows = []
    for i, row in enumerate(table.rows):
        cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
        rows.append("| " + " | ".join(cells) + " |")
        if i == 0:
            rows.append("| " + " | ".join(["---"] * len(cells)) + " |")
    return "\n".join(rows)


def docx_to_markdown(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    lines = []

    # 遍历文档中的块级元素（段落 + 表格）
    for block in doc.element.body:
        tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag

        if tag == "p":
            # 包装为 Paragraph 对象
            from docx.text.paragraph import Paragraph
            para = Paragraph(block, doc)
            text = para.text.strip()
            if not text:
                lines.append("")
                continue

            level = _get_heading_level(para)
            if level:
                lines.append(f"{'#' * level} {text}")
            elif _is_list_paragraph(para):
                lines.append(f"- {text}")
            else:
                lines.append(text)

        elif tag == "tbl":
            from docx.table import Table
            table = Table(block, doc)
            lines.append("")
            lines.append(_table_to_md(table))
            lines.append("")

    return "\n".join(lines)
