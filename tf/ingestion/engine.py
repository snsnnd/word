"""接入引擎（见 docs/02 §4、docs/06 §2.2）。

ingest(docx_path) -> list[Block]

按文档顺序遍历正文（段落与表格交错），构建 Block 列表：
- 段落 → BlockKind.PARAGRAPH / IMAGE / FORMULA（视内容探测）
- 表格 → BlockKind.TABLE
- 分节符 → BlockKind.SECTION_BREAK（段内 sectPr 的段落）

只读，不修改文档对象模型。xml_element 用 lxml 引用，渲染层可直接原地改。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Optional

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

from tf.core.enums import BlockKind, BlockSource, VertAlign
from tf.core.idgen import make_block_id
from tf.core.models import AssetInfo, Block, RunInfo
from tf.ingestion.assets import (
    W,
    detect_paragraph_kind,
    has_footnote_ref,
    has_section_break,
    qn,
    table_dimensions,
)

log = logging.getLogger(__name__)


def ingest(docx_path: str | Path) -> list[Block]:
    """把 docx 转成块列表。"""
    path = Path(docx_path)
    document = docx.Document(str(path))
    body = document.element.body

    blocks: list[Block] = []
    index = 0
    section_idx = 0

    for child in body:
        tag = child.tag
        if tag == qn("w:p"):
            block = _build_paragraph_block(child, document, index, section_idx)
            if block is not None:
                blocks.append(block)
                index += 1
            if has_section_break(child):
                section_idx += 1
        elif tag == qn("w:tbl"):
            block = _build_table_block(child, document, index, section_idx)
            if block is not None:
                blocks.append(block)
                index += 1
        elif tag == qn("w:sdt"):
            # 内容控件：递归处理其中的段落（best-effort）
            log.warning("[ingest] 检测到 w:sdt 内容控件，已跳过（不影响其它块）。")
            continue
        elif tag == qn("w:sectPr"):
            # 文末分节符，标志最后一个 section 结束
            section_idx += 1
            continue
        else:
            continue

    return blocks


def _build_paragraph_block(
    p_element, document, index: int, section_idx: int
) -> Optional[Block]:
    paragraph = Paragraph(p_element, document)
    text = paragraph.text or ""
    kind, asset = detect_paragraph_kind(p_element)

    runs = _extract_runs(p_element)
    block_id = make_block_id(index, text, p_element)
    footnote = has_footnote_ref(p_element)

    return Block(
        block_id=block_id,
        index=index,
        kind=kind,
        protected=kind != BlockKind.PARAGRAPH,
        asset=asset,
        text=text,
        runs=runs,
        xml_element=p_element,
        parent_paragraph=paragraph,
        section_index=section_idx,
        semantic_tag=None,
        confidence=0.0,
        source=None,
    )


def _build_table_block(
    tbl_element, document, index: int, section_idx: int
) -> Optional[Block]:
    table = Table(tbl_element, document)
    rows, cols, has_header = table_dimensions(tbl_element)
    # 表格纯文本（逐行拼接），用于规则表/LLM 判断
    text = _table_text(tbl_element)
    asset = AssetInfo(
        kind=BlockKind.TABLE,
        rows=rows,
        cols=cols,
        has_header=has_header,
    )
    block_id = make_block_id(index, text, tbl_element)
    return Block(
        block_id=block_id,
        index=index,
        kind=BlockKind.TABLE,
        protected=True,
        asset=asset,
        text=text,
        runs=[],
        xml_element=tbl_element,
        parent_paragraph=None,
        section_index=section_idx,
        semantic_tag=None,
        confidence=0.0,
        source=None,
    )


def _table_text(tbl_element) -> str:
    """表格纯文本，逐行用 ' | ' 连接，行用 '\\n'。"""
    lines: list[str] = []
    for tr in tbl_element.findall(qn("w:tr")):
        cells: list[str] = []
        for tc in tr.findall(qn("w:tc")):
            cell_text = "".join(
                t.text or "" for t in tc.iter(qn("w:t"))
            )
            cells.append(cell_text.strip())
        lines.append(" | ".join(cells))
    return "\n".join(lines)


def _extract_runs(p_element) -> list[RunInfo]:
    """提取段落的 run 级信息。"""
    runs: list[RunInfo] = []
    for r in p_element.iter(qn("w:r")):
        text = "".join(t.text or "" for t in r.iter(qn("w:t")))
        if not text:
            # 可能是只有 drawing/object 的 run，仍保留一个空文本条目以便渲染
            has_visual = r.find(qn("w:drawing")) is not None or r.find(qn("w:object")) is not None
            if not has_visual:
                continue
        rpr = r.find(qn("w:rPr"))
        font_cn = font_en = None
        size_pt = bold = italic = None
        vert = None
        if rpr is not None:
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is not None:
                font_cn = rfonts.get(qn("w:eastAsia"))
                font_en = rfonts.get(qn("w:ascii"))
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                try:
                    size_pt = int(sz.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass
            if rpr.find(qn("w:b")) is not None:
                bold = True
            if rpr.find(qn("w:i")) is not None:
                italic = True
            va = rpr.find(qn("w:vertAlign"))
            if va is not None:
                val = va.get(qn("w:val"))
                if val == "superscript":
                    vert = VertAlign.SUPERSCRIPT
                elif val == "subscript":
                    vert = VertAlign.SUBSCRIPT
        runs.append(RunInfo(
            text=text,
            font_cn=font_cn,
            font_en=font_en,
            size_pt=size_pt,
            bold=bold,
            italic=italic,
            vert_align=vert,
        ))
    return runs