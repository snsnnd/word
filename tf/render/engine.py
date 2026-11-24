"""渲染引擎（M2，见 docs/05、docs/06 §2.3）。

`render(blocks, spec, out_path)`：把源文档解析出的块，**灌入母版副本**并套用样式。

流程（对齐 docs/05 §1）：
1. 载入母版为工作副本（源母版只读）；
2. 清空母版正文占位内容（保留样式 / 页眉页脚 / 分节 / 目录域）；
3. 逐块渲染：保护块（table/image/formula）原样搬入、普通段落套样式 + 字体分离；
4. 保存为 out_path（.docx）。

设计取舍：
- **直接格式化优先**：块上逐段/逐 run 写 `w:rFonts`/`w:sz`/`w:ind`/`w:jc`，
  保证跨 Word/WPS 一致，同时仍套用母版 Heading N 样式（存在时）以获得大纲级别。
- **样式继承**：字段级级联 —— 规则未显式给出的字段继承 `default`。
- **保护块不改内容**：表格 / 公式（OMML/OLE）只搬迁不 fonts；图片按 max_width 等比缩放。
"""
from __future__ import annotations

import logging
import shutil
import subprocess
import sys
import tempfile
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any, Optional

import docx
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

from tf.core.enums import Align, BlockKind, BlockType
from tf.core.models import (
    Block,
    FinalizeResult,
    RenderResult,
    Style,
    TemplateSpec,
)
from tf.render.ooxml import set_outline_level, set_para_format, set_run_font

log = logging.getLogger(__name__)

# 这些语义标签的标题/特殊块不首行缩进（渲染层策略，不改规则表）
_NO_INDENT_TAGS = frozenset(
    {
        BlockType.COVER_TITLE,
        BlockType.COVER_META,
        BlockType.ABSTRACT_TITLE,
        BlockType.KEYWORDS,
        BlockType.TOC,
        BlockType.HEADING_1,
        BlockType.HEADING_2,
        BlockType.HEADING_3,
        BlockType.HEADING_4,
        BlockType.CAPTION,
        BlockType.REFERENCE_TITLE,
        BlockType.APPENDIX_TITLE,
    }
)

_HEADING_TAGS = {
    BlockType.HEADING_1: 1,
    BlockType.HEADING_2: 2,
    BlockType.HEADING_3: 3,
    BlockType.HEADING_4: 4,
}


def render(
    blocks: list[Block],
    spec: TemplateSpec,
    out_path: str | Path,
) -> RenderResult:
    """把 blocks 灌入母版并套样式，输出到 out_path。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = _load_master(spec.master_path)
    _clear_body(doc)
    sectPr = _ensure_sectPr(doc)

    src_doc = _source_doc(blocks)
    stats = {"blocks": len(blocks), "rendered": 0, "skipped": 0, "warnings": []}

    for block in blocks:
        try:
            placed, ok, msgs = _place_block(doc, src_doc, sectPr, block, spec)
            # 块级告警只收进 stats，由调用方统一打印，避免与日志重复输出
            stats["warnings"].extend(msgs)
            if not ok:
                stats["skipped"] += 1
                continue
            if placed is not None:
                _apply_block(doc, placed, block, spec)
            stats["rendered"] += 1
        except Exception as exc:  # noqa: BLE001 - 单块失败不阻断整体渲染
            stats["skipped"] += 1
            stats["warnings"].append(f"块 {block.block_id} 渲染失败：{exc}")
            log.debug("[render] 块 %s 渲染失败：%s", block.block_id, exc)

    doc.save(str(out_path))
    return RenderResult(
        docx_path=str(out_path),
        pdf_path=None,
        qa_report={"stats": stats, "out": str(out_path)},
        engine="m2-direct",
    )


# ------------------------------------------------------------------ 母版装载
def _load_master(master_path: Optional[str | Path]) -> Any:
    """载入母版为工作副本；无母版时退回空白文档。"""
    if master_path:
        p = Path(master_path)
        if p.exists():
            return docx.Document(str(p))
        log.warning("[render] 母版不存在：%s，改用空白文档", p)
    return docx.Document()


def _clear_body(doc: Any) -> None:
    """清空正文，仅保留末尾 `w:sectPr`（分节/页面设置）。"""
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def _ensure_sectPr(doc: Any):
    """返回正文末尾的 `w:sectPr`，不存在则新建。"""
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    if sectPr is None:
        sectPr = body.makeelement(qn("w:sectPr"), {})
        body.append(sectPr)
    return sectPr


def _source_doc(blocks: list[Block]) -> Optional[Any]:
    """从块反查源 Document（用于图片关系/公式搬迁）。"""
    for b in blocks:
        para = b.parent_paragraph
        if para is not None and getattr(para, "_parent", None) is not None:
            return para._parent
    return None


# ------------------------------------------------------------------ 放置块
def _place_block(doc, src_doc, sectPr, block: Block, spec: TemplateSpec):
    """把块搬入工作副本。

    返回 `(paragraph_or_None, placed, warnings)`：
    - placed=False 表示未搬入（无源 XML 的空块、图片插入失败等），计入 skipped；
    - paragraph 非 None 时需要套样式，None 表示纯搬迁（如表格）；
    - warnings 为该块的告警文本，由调用方统一收集输出。
    """
    kind = block.kind

    if kind == BlockKind.TABLE:
        # 保护块：原样搬入，不改内容
        if block.xml_element is not None:
            sectPr.addprevious(deepcopy(block.xml_element))
            return None, True, []
        return None, False, []

    if kind == BlockKind.IMAGE:
        return _place_image(doc, src_doc, sectPr, block, spec)

    if kind == BlockKind.FORMULA:
        # 公式：保留 OMML 树，仅调整段落（居中/行距），不动 run 字体
        para = _place_paragraph(doc, sectPr, block)
        set_para_format(para, align=Align.CENTER)
        return para, True, []

    if kind in (BlockKind.SECTION_BREAK, BlockKind.TOC, BlockKind.FOOTNOTE_REF):
        if block.xml_element is not None:
            sectPr.addprevious(deepcopy(block.xml_element))
            return None, True, []
        return None, False, []

    # 普通段落（含标题/题注/参考文献/封面等，ingestion 均记为 PARAGRAPH）
    return _place_paragraph(doc, sectPr, block), True, []


def _place_paragraph(doc, sectPr, block: Block) -> Paragraph:
    """把段落块搬入工作副本并返回 Paragraph 包装。"""
    if block.xml_element is not None:
        new_el = deepcopy(block.xml_element)
        sectPr.addprevious(new_el)
        return Paragraph(new_el, doc)
    # 无源 XML（手工构造的块）：新建段落
    p = doc.add_paragraph(block.text)
    sectPr.addprevious(p._p)  # 移到 sectPr 之前
    return p


def _place_image(
    doc, src_doc, sectPr, block: Block, spec: TemplateSpec
) -> tuple[Paragraph, bool, list[str]]:
    """图片块：从源文档关系取 blob，重新插入工作副本（重建关系）。

    失败时返回 `placed=False` 并给出告警文本，由调用方计入 skipped。
    """
    p = doc.add_paragraph()
    sectPr.addprevious(p._p)

    rel_id = block.asset.rel_id if block.asset else None
    if rel_id and src_doc is not None:
        try:
            part = src_doc.part.related_parts[rel_id]
            run = p.add_run()
            shape = run.add_picture(BytesIO(part.blob))
            max_w = spec.rules.image_cfg.get("max_width_cm") if spec.rules else None
            if max_w and shape.width:
                limit = Cm(float(max_w))
                if shape.width > limit:
                    ratio = limit / shape.width
                    shape.width = int(shape.width * ratio)
                    shape.height = int(shape.height * ratio)
            return p, True, []
        except Exception as exc:  # noqa: BLE001
            return p, False, [f"图片块 {block.block_id} 插入失败：{exc}"]
    return p, False, [f"图片块 {block.block_id} 缺少关系 id，已跳过"]


# ------------------------------------------------------------------ 套样式
def _apply_block(doc, paragraph: Paragraph, block: Block, spec: TemplateSpec) -> None:
    """按块的语义标签套用样式（直接格式化）。"""
    tag = block.semantic_tag
    if tag is None or tag == BlockType.UNKNOWN:
        return
    if tag == BlockType.IMAGE:
        set_para_format(paragraph, align=_image_align(spec))
        return
    _apply_paragraph(doc, paragraph, block, spec, tag)


def _apply_paragraph(
    doc, paragraph: Paragraph, block: Block, spec: TemplateSpec, tag: BlockType
) -> None:
    """段级 + 字符级样式。"""
    style = spec.rules.resolve(tag) if spec.rules is not None else Style()

    # 悬挂缩进优先于首行缩进；规则同时给出时以悬挂为准
    first_line = style.first_line_indent_chars
    hanging = style.hanging_indent_chars
    if tag in _NO_INDENT_TAGS:
        first_line = None
        hanging = None
    if hanging is not None:
        first_line = None

    # 标题：套用 Heading N 样式（保留母版大纲级别），失败则回退纯直接格式化
    level = _HEADING_TAGS.get(tag)
    if level is not None:
        _set_heading_style(doc, paragraph, level)

    set_para_format(
        paragraph,
        align=style.align,
        line_spacing=style.line_spacing,
        first_line_indent_chars=first_line,
        hanging_indent_chars=hanging,
        space_before_pt=style.space_before_pt,
        space_after_pt=style.space_after_pt,
        font_size_pt=style.size_pt,
    )

    for run in paragraph.runs:
        set_run_font(
            run,
            font_cn=style.font_cn,
            font_en=style.font_en,
            size_pt=style.size_pt,
            bold=style.bold,
            italic=style.italic,
        )


def _set_heading_style(doc, paragraph: Paragraph, level: int) -> None:
    """套 Heading N 样式 + outlineLvl；母版缺该样式时静默降级。"""
    try:
        paragraph.style = doc.styles[f"Heading {level}"]
    except KeyError:
        pass
    set_outline_level(paragraph, level)


def _image_align(spec: TemplateSpec) -> Align:
    raw = spec.rules.image_cfg.get("alignment", "center") if spec.rules else "center"
    try:
        return Align(raw)
    except ValueError:
        return Align.CENTER


# ------------------------------------------------------------------ 收尾
def _detect_finalizer() -> str:
    """选择收尾引擎：word（Windows + Word）> libreoffice > none。

    只探测、不导入：`win32com` 在非 Windows 上不存在，`shutil.which` 用于找
    `soffice`/`libreoffice`。两者都不可用时降级为 none（docs/07 降级矩阵）。
    """
    if sys.platform.startswith("win"):
        try:  # pragma: no cover - 仅 Windows 且装了 Word 时命中
            import win32com.client  # noqa: F401
        except Exception:  # noqa: BLE001
            pass
        else:
            return "word"
    for name in ("soffice", "libreoffice"):
        if shutil.which(name):
            return "libreoffice"
    return "none"


def _finalize_with_word(docx_path: Path, spec: TemplateSpec) -> FinalizeResult:
    """Windows + Word COM：更新目录域/交叉引用，另存 PDF。"""
    import win32com.client  # 仅 Windows 有

    result = FinalizeResult(engine="word")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path.resolve()))
        # 1) 更新所有域（交叉引用、SEQ 等）
        doc.Fields.Update()
        # 2) 更新目录（TOC 是独立集合，Fields.Update 覆盖不到）
        for i in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(i).Update()
        # 3) 重新分页，保证目录页码与实际一致
        doc.Repaginate()
        pdf_path = docx_path.with_suffix(".pdf")
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_path.resolve()),
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,
        )
        result.toc_updated = True
        result.fields_updated = True
        result.pdf_path = str(pdf_path)
        result.docx_path = str(docx_path)
    except Exception as exc:  # noqa: BLE001
        result.degraded = True
        result.warnings.append(f"Word 收尾失败：{exc}")
    finally:
        try:
            if doc is not None:
                doc.Close(Saving=False)
            word.Quit()
        except Exception:  # noqa: BLE001
            pass
    return result


def _finalize_with_libreoffice(docx_path: Path, spec: TemplateSpec) -> FinalizeResult:
    """Linux：LibreOffice 无头导出 PDF。

    LO 命令行不更新 TOC 域，因此 `toc_updated=False`，并提示用户在 Word 里
    按 Ctrl+A → F9（或手动在 LO 里更新索引）后重新导出。
    """
    result = FinalizeResult(engine="libreoffice")
    pdf_path = docx_path.with_suffix(".pdf")
    # 独立 UserInstallation：避免与常驻 soffice 进程抢同一个用户配置（会静默失败）。
    # 放在系统临时目录，不在用户产物目录里留下 .lo_profile 之类的垃圾。
    profile = Path(tempfile.mkdtemp(prefix="tf_lo_"))
    cmd = [
        _find_soffice(),
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--headless", "--norestore", "--invisible",
        "--convert-to", "pdf",
        "--outdir", str(docx_path.parent.resolve()),
        str(docx_path.resolve()),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    except Exception as exc:  # noqa: BLE001
        result.degraded = True
        result.warnings.append(f"LibreOffice 导出失败：{exc}")
        return result
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if proc.returncode != 0 or not pdf_path.exists():
        result.degraded = True
        detail = (proc.stderr or proc.stdout or "").strip().splitlines()
        result.warnings.append(
            "LibreOffice 导出失败：" + (detail[-1] if detail else f"退出码 {proc.returncode}")
        )
        return result
    result.pdf_path = str(pdf_path)
    result.docx_path = str(docx_path)
    result.warnings.append(
        "目录/交叉引用未刷新（LibreOffice 命令行不更新域）。请在 Word 中按 Ctrl+A → F9 后重新导出。"
    )
    return result


def _find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    raise RuntimeError("未找到 LibreOffice 可执行文件")


def finalize(docx_path: str | Path, spec: TemplateSpec) -> FinalizeResult:
    """Word 原生收尾（M3）：Word > LibreOffice > none，逐级降级（docs/07 降级矩阵）。"""
    path = Path(docx_path)
    engine = _detect_finalizer()
    if engine == "word":
        return _finalize_with_word(path, spec)
    if engine == "libreoffice":
        return _finalize_with_libreoffice(path, spec)
    result = FinalizeResult(engine="none")
    result.degraded = True
    result.docx_path = str(path) if path.exists() else None
    result.warnings.append(
        "未检测到 Word / LibreOffice，目录/交叉引用未刷新。请在 Word 中按 Ctrl+A → F9 更新域后导出 PDF。"
    )
    return result


class RenderEngine:
    """协议实现（docs/05 §8）。"""

    def render(self, blocks: list[Block], spec: TemplateSpec, out_path) -> RenderResult:
        return render(blocks, spec, out_path)

    def finalize(self, docx_path, spec: TemplateSpec) -> FinalizeResult:
        return finalize(docx_path, spec)