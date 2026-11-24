"""渲染层测试（docs/07 §2 测试清单 · test_render）。

覆盖：
- 规则级联（字段级继承，docs/06 §2.2）
- 灌母版：保留页眉 / 清空占位 / 搬运源内容
- 套样式：正文首行缩进、标题大纲级别、图表标题不缩进、参考文献悬挂缩进
- 保护块：图片居中且原样搬入、表格内容不变
- RenderEngine 协议与 finalize 降级
"""
from __future__ import annotations

import base64
import shutil
import zipfile
from io import BytesIO
from pathlib import Path

import docx
import pytest
from docx.oxml.ns import qn
from docx.shared import Cm

from tf.core.enums import Align, BlockKind, BlockSource, BlockType
from tf.core.models import AssetInfo, Block, RenderResult, Style, TemplateSpec
from tf.parser.regex_parser import classify
from tf.ingestion.engine import ingest
from tf.render import render
from tf.render.engine import RenderEngine, finalize
from tf.template.rules import default_rules_path, load_rules

# 1×1 PNG（合法 PNG，python-docx 可解析）
PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQ"
    "AAAABJRU5ErkJggg=="
)

MASTER_HEADER = "MASTER-HEADER"
MASTER_PLACEHOLDER = "MASTER-PLACEHOLDER"

BODY_TEXT = "这是正文段落，用于验证首行缩进与中文字体是否按规则套用。"
CAPTION_TEXT = "图 1-1 系统总体框图"
REF_TITLE = "参考文献"
REF_ITEM = "[1] 张三. 论文题目[J]. 某期刊, 2020, 1(1): 1-10."


# ------------------------------------------------------------------ 工具
def _make_master(path: Path) -> Path:
    """构造母版：含 Heading 1 样式 + 页眉 + 占位正文。"""
    d = docx.Document()
    d.add_heading("Master Heading", level=1)
    d.add_paragraph(MASTER_PLACEHOLDER)
    sec = d.sections[0]
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = MASTER_HEADER
    d.save(str(path))
    return path


def _make_source(path: Path) -> Path:
    """构造源文档：章标题 / 二级标题 / 正文 / 图片 / 图题 / 参考文献。"""
    d = docx.Document()
    d.add_paragraph("第一章 绪论")
    d.add_paragraph("1.1 研究背景")
    d.add_paragraph(BODY_TEXT)
    d.add_paragraph().add_run().add_picture(BytesIO(PNG_BYTES), width=Cm(2))
    d.add_paragraph(CAPTION_TEXT)
    d.add_paragraph(REF_TITLE)
    d.add_paragraph(REF_ITEM)
    d.save(str(path))
    return path


def _spec(master: Path) -> TemplateSpec:
    rules = load_rules(default_rules_path())
    return TemplateSpec(
        name="test",
        version="0.1.0",
        master_path=str(master),
        rules=rules,
        style_snapshot={},
    )


def _blocks(source: Path) -> list[Block]:
    blocks = ingest(source)
    classify(blocks)
    return blocks


def _para_by_text(document, text: str):
    for p in document.paragraphs:
        if p.text.strip() == text.strip():
            return p
    raise AssertionError(f"输出中找不到段落：{text!r}")


def _val(element, attr: str):
    if element is None:
        return None
    node = element.find(qn(attr))
    return None if node is None else node.get(qn("w:val"))


def _pPr(paragraph):
    return paragraph._p.pPr


def _ind(paragraph):
    pPr = _pPr(paragraph)
    return None if pPr is None else pPr.find(qn("w:ind"))


def _run_props(paragraph):
    """返回首个 run 的 (eastAsia 字体, 半磅字号)。"""
    if not paragraph.runs:
        return (None, None)
    rPr = paragraph.runs[0]._r.find(qn("w:rPr"))
    if rPr is None:
        return (None, None)
    fonts = rPr.find(qn("w:rFonts"))
    sz = rPr.find(qn("w:sz"))
    return (
        None if fonts is None else fonts.get(qn("w:eastAsia")),
        None if sz is None else sz.get(qn("w:val")),
    )


# ------------------------------------------------------------------ 规则级联
def test_resolve_partial_inherits_from_default():
    """specific 只覆盖显式给出的字段，其余继承 default。"""
    rules = load_rules(default_rules_path())
    body = rules.resolve(BlockType.BODY)
    assert body.font_cn == "宋体"  # 继承 default
    assert body.line_spacing == 1.5  # 继承 default
    assert body.size_pt == 12.0


def test_resolve_partial_overrides_only_size():
    """只设 size 的规则不会清空 default 的字体与行距（回归：整块替换）。"""
    base = Style(font_cn="黑体", font_en="Times New Roman", size_pt=12.0, line_spacing=1.5)
    thin = Style(size_pt=9.0)
    rules = _rules_with(base, {BlockType.BODY: thin})
    body = rules.resolve(BlockType.BODY)
    assert body.size_pt == 9.0
    assert body.font_cn == "黑体"
    assert body.line_spacing == 1.5


def test_resolve_undefined_tag_returns_default():
    rules = load_rules(default_rules_path())
    assert rules.resolve(BlockType.UNKNOWN).size_pt == rules.default.size_pt


def _rules_with(default: Style, styles: dict):
    from tf.core.models import RuleSet

    return RuleSet(default=default, styles=styles)


# ------------------------------------------------------------------ 灌母版
def test_render_pours_into_master(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    out = tmp_path / "out.docx"

    result = render(_blocks(source), _spec(master), out)
    assert out.exists()

    d = docx.Document(str(out))
    text = "\n".join(p.text for p in d.paragraphs)
    # 源内容搬入
    assert "第一章 绪论" in text
    assert BODY_TEXT in text
    # 母版占位被清空
    assert MASTER_PLACEHOLDER not in text
    assert "Master Heading" not in text
    # 页眉保留
    assert any(MASTER_HEADER in p.text for p in d.sections[0].header.paragraphs)


def test_render_all_blocks_rendered(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    blocks = _blocks(source)
    result = render(blocks, _spec(master), tmp_path / "out.docx")
    stats = result.qa_report["stats"]
    assert stats["blocks"] == len(blocks)
    assert stats["rendered"] == len(blocks)
    assert stats["skipped"] == 0
    assert stats["warnings"] == []
    assert result.engine == "m2-direct"
    assert result.pdf_path is None


# ------------------------------------------------------------------ 套样式
def test_render_body_style(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    p = _para_by_text(docx.Document(str(tmp_path / "out.docx")), BODY_TEXT)
    pPr = _pPr(p)
    assert _val(pPr, "w:jc") == "both"
    assert _run_props(p) == ("宋体", "24")  # 12pt → 24 半磅
    ind = _ind(p)
    assert ind is not None and ind.get(qn("w:firstLineChars")) == "200"


def test_render_heading_outline_levels(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    d = docx.Document(str(tmp_path / "out.docx"))
    h1 = _para_by_text(d, "第一章 绪论")
    h2 = _para_by_text(d, "1.1 研究背景")
    assert _run_props(h1) == ("黑体", "32")  # 三号 16pt
    assert _val(_pPr(h1), "w:outlineLvl") == "0"
    assert _run_props(h2) == ("黑体", "28")  # 四号 14pt
    assert _val(_pPr(h2), "w:outlineLvl") == "1"


def test_render_caption_no_first_line_indent(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    p = _para_by_text(docx.Document(str(tmp_path / "out.docx")), CAPTION_TEXT)
    assert _val(_pPr(p), "w:jc") == "center"
    ind = _ind(p)
    # 图表标题不缩进（无 firstLineChars / hangingChars）
    assert ind is None or (
        ind.get(qn("w:firstLineChars")) is None and ind.get(qn("w:hangingChars")) is None
    )
    assert _run_props(p) == ("宋体", "21")  # 五号 10.5pt


def test_render_reference_hanging_indent(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    d = docx.Document(str(tmp_path / "out.docx"))
    p = _para_by_text(d, REF_ITEM)
    ind = _ind(p)
    assert ind is not None and ind.get(qn("w:hangingChars")) == "200"
    # 悬挂缩进与首行缩进互斥
    assert ind.get(qn("w:firstLineChars")) is None


# ------------------------------------------------------------------ 保护块
def test_render_image_centered_and_preserved(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    d = docx.Document(str(tmp_path / "out.docx"))
    assert len(d.inline_shapes) == 1  # 图片原样搬入
    img_para = next(p for p in d.paragraphs if p._p.findall(".//" + qn("w:drawing")))
    assert _val(_pPr(img_para), "w:jc") == "center"


def test_render_table_preserved(tmp_path: Path):
    master = _make_master(tmp_path / "master.docx")
    source = tmp_path / "src_table.docx"
    d = docx.Document()
    d.add_paragraph("第一章 绪论")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "表头A"
    t.cell(1, 1).text = "数据B"
    d.save(str(source))

    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    out = docx.Document(str(tmp_path / "out.docx"))
    assert len(out.tables) == 1
    assert out.tables[0].cell(0, 0).text == "表头A"
    assert out.tables[0].cell(1, 1).text == "数据B"


# ------------------------------------------------------------------ 协议/收尾
def test_render_engine_protocol(tmp_path: Path):
    engine = RenderEngine()
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    result = engine.render(_blocks(source), _spec(master), tmp_path / "out.docx")
    assert isinstance(result, RenderResult)
    assert Path(result.docx_path).exists()


def test_finalize_degrades_without_any_converter(tmp_path: Path, monkeypatch):
    """Word 与 LibreOffice 都没有时：降级为 none + 提示手动更新域。"""
    master = _make_master(tmp_path / "master.docx")
    from tf.render import engine as engine_mod

    monkeypatch.setattr(engine_mod, "_detect_finalizer", lambda: "none")
    result = finalize(tmp_path / "out.docx", _spec(master))
    assert result.engine == "none"
    assert result.degraded is True
    assert result.warnings  # 提示用户手动更新域


@pytest.mark.skipif(shutil.which("soffice") is None, reason="环境无 LibreOffice")
def test_finalize_uses_libreoffice_to_export_pdf(tmp_path: Path):
    """Linux 下走 LibreOffice 无头导出 PDF，并声明目录未刷新。"""
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    result = finalize(tmp_path / "out.docx", _spec(master))
    assert result.engine == "libreoffice"
    assert result.pdf_path and Path(result.pdf_path).exists()
    assert result.toc_updated is False
    assert any("未刷新" in w for w in result.warnings)


def test_render_without_master_falls_back_to_blank(tmp_path: Path):
    """母版缺失时不崩，退回空白文档渲染。"""
    source = _make_source(tmp_path / "src.docx")
    spec = TemplateSpec(name="t", version="0.1.0", master_path=None, rules=load_rules(default_rules_path()))
    out = tmp_path / "out.docx"
    result = render(_blocks(source), spec, out)
    assert out.exists()
    assert result.qa_report["stats"]["rendered"] == result.qa_report["stats"]["blocks"]


@pytest.mark.parametrize(
    "tag",
    [
        BlockType.HEADING_1,
        BlockType.CAPTION,
        BlockType.REFERENCE_TITLE,
        BlockType.COVER_TITLE,
    ],
)
def test_no_indent_tags_have_no_first_line_indent(tmp_path: Path, tag):
    """渲染层策略：标题/图表题不缩进（docs/06 §2.2 与规则表解耦）。"""
    from tf.render.engine import _NO_INDENT_TAGS

    assert tag in _NO_INDENT_TAGS


# ------------------------------------------------------------------ 快照 / 告警
def test_render_snapshot_document_xml(tmp_path: Path):
    """快照测试（docs/07 §4）：解包产出 docx，比对 document.xml / styles.xml 片段。"""
    master = _make_master(tmp_path / "master.docx")
    source = _make_source(tmp_path / "src.docx")
    render(_blocks(source), _spec(master), tmp_path / "out.docx")

    with zipfile.ZipFile(tmp_path / "out.docx") as zf:
        names = set(zf.namelist())
        assert "word/document.xml" in names
        xml = zf.read("word/document.xml").decode("utf-8")
        styles = zf.read("word/styles.xml").decode("utf-8")
        headers = [n for n in names if n.startswith("word/header")]
        header_xml = [zf.read(n).decode("utf-8", "ignore") for n in headers]

    # 正文首行缩进（200 = 2 字符）与西文字体分离
    assert 'w:firstLineChars="200"' in xml
    assert 'w:ascii="Times New Roman"' in xml
    # 标题居中 + 三号黑体（w:sz 为半磅，16pt → 32）
    assert 'w:jc w:val="center"' in xml
    assert 'w:eastAsia="黑体"' in xml
    assert 'w:sz w:val="32"' in xml
    # 母版页眉随母版保留
    assert headers
    assert any(MASTER_HEADER in h for h in header_xml)
    # 母版命名样式保留（策略 A：只使用母版自带 styleId，不注入新样式）
    assert 'w:styleId="Heading1"' in styles


def test_render_reports_image_block_without_rel_id(tmp_path: Path):
    """图片块缺关系 id：计入 skipped 并产出告警，不得伪装成已渲染。"""
    para = docx.Document().add_paragraph("占位")
    blocks = [
        Block(
            block_id="p0", index=0, kind=BlockKind.PARAGRAPH, text="占位",
            xml_element=para._p, parent_paragraph=para,
            semantic_tag=BlockType.BODY, source=BlockSource.REGEX, confidence=1.0,
        ),
        Block(
            block_id="p1", index=1, kind=BlockKind.IMAGE,
            asset=AssetInfo(kind=BlockKind.IMAGE, width_cm=6.0, rel_id=None),
            semantic_tag=BlockType.IMAGE,
        ),
    ]
    out = tmp_path / "out.docx"
    result = render(blocks, TemplateSpec(name="t", version="0.1.0", master_path=None,
                                        rules=load_rules(default_rules_path())), out)

    stats = result.qa_report["stats"]
    assert stats["rendered"] == 1
    assert stats["skipped"] == 1
    assert any("缺少关系 id" in w for w in stats["warnings"])