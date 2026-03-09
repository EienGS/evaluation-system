"""
将 .docx / .doc 文件转换为 Markdown 文本
依赖：python-docx（docx），LibreOffice（doc→docx 转换）
"""
from docx import Document
from docx.oxml.ns import qn
import io
import os
import subprocess
import tempfile


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


def doc_to_markdown(file_bytes: bytes) -> str:
    """
    将 .doc 文件转换为 Markdown。
    先用 LibreOffice 将 .doc 转换为 .docx，再调用 docx_to_markdown。
    需要系统安装 LibreOffice（soffice 命令可用）。
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        doc_path = os.path.join(tmpdir, "input.doc")
        with open(doc_path, "wb") as f:
            f.write(file_bytes)

        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "docx", "--outdir", tmpdir, doc_path],
                check=True,
                timeout=60,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            raise RuntimeError(
                "doc 转换失败：请确保服务器已安装 LibreOffice（soffice 命令可用）。原始错误：" + str(e)
            )

        docx_path = os.path.join(tmpdir, "input.docx")
        if not os.path.exists(docx_path):
            raise RuntimeError("LibreOffice 转换后未生成 .docx 文件，请检查 LibreOffice 安装。")

        with open(docx_path, "rb") as f:
            docx_bytes = f.read()

    return docx_to_markdown(docx_bytes)


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
