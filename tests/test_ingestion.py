"""tf.ingestion.engine 的结构化导入验收测试（task-7 任务书 §4）。

本文件覆盖「导入」这一层：
  1. 每个块都有稳定唯一的 block_id 和序号 index
  2. 段落块保留 runs（文本、粗体/斜体、中文字体字号）
  3. 图片、表格被识别并标记为 protected
  4. 表格行列数（用真实模板 source_template.docx 验证）
  5. 图片相对路径（rel_id）可读
  6. section_index 取值正确（用真实模板验证 0 与 1）

分工：正则判定路径（摘要 / 关键词 / 参考文献 / 章节标题 / 图题 / 表题 /
目录 / 附录 / 封面标签行 / 表格标题行）见
``tests/test_parser_regex.py``。本文件只负责导入层与「跨全库排序」不变量。
"""
import base64
import os
from io import BytesIO

import pytest
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from tf.core.enums import BlockKind
from tf.ingestion.engine import ingest

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
TEMPLATE = os.path.join(FIXTURES, "source_template.docx")
NOT_PASSED = object()  # getparent() 返回 None 时的哨兵


# ---------- 构造 fixture 的辅助函数 ----------

def _build(path, specs):
    """按 specs 生成文档。

    specs 每项可为：
      * str            —— 单 run 纯文本段落
      * list[dict]     —— 多 run 段落，dict 支持 text / bold / italic / font / size(磅)
      * dict           —— 单 run 段落（同上，省去外层列表）
    """
    d = Document()
    for spec in specs:
        p = d.add_paragraph()
        if isinstance(spec, str):
            p.add_run(spec)
            continue
        runs = [spec] if isinstance(spec, dict) else spec
        for rs in runs:
            r = p.add_run(rs["text"])
            if "bold" in rs:
                r.bold = rs["bold"]
            if "italic" in rs:
                r.italic = rs["italic"]
            if rs.get("font"):
                r.font.name = rs["font"]
                r._element.rPr.rFonts.set(qn("w:eastAsia"), rs["font"])
            if rs.get("size"):
                r.font.size = Pt(rs["size"])
    d.save(path)
    return path


def _params():
    return [
        pytest.param(
            "入门导言", 9, id="大纲第9段「入门导言」不是任何标题"
        ),
        pytest.param(
            "绪论", 14, id="大纲第14段「绪论」不在参考列表里（特判验证点）"
        ),
        pytest.param(
            "正文", 18, id="大纲第18段「正文」本身也不是标题"
        ),
        pytest.param(
            "系统需求",
            36,
            id="大纲第36段「系统需求」不在大纲/参考/已知之一（特判验证点）",
        ),
    ]


@pytest.fixture(scope="module")
def template_blocks():
    return ingest(TEMPLATE)


_PARAM_PATH = "/tmp/t_params.docx"


def _params_docx():
    """把 _params() 的四段大纲文字写成 docx，作为导入层测试的统一输入。

    _params() 返回的是 pytest.param 列表，不能直接喂给 ingest()（它要的是路径）。
    """
    return _build(_PARAM_PATH, [p.values[0] for p in _params()])


# ---------- 1. block_id 稳定且唯一 ----------

def test_block_id_非空():
    bs = _ingest(_params_docx())
    assert all(isinstance(b.block_id, str) and b.block_id for b in bs)


def test_block_id_在文档内唯一():
    bs = _ingest(_params_docx())
    ids = [b.block_id for b in bs]
    assert len(set(ids)) == len(ids)


def test_block_id_跨重复解析稳定():
    bs1 = _ingest(_params_docx())
    bs2 = _ingest(_params_docx())
    assert [b.block_id for b in bs1] == [b.block_id for b in bs2]


def test_block_id_可反解序号():
    bs = _ingest(_params_docx())
    from tf.core.idgen import parse_index

    assert all(parse_index(b.block_id) == b.index for b in bs)


def test_block_id_相同文本不同上下文不撞():
    """index 前缀本身保证不同序号块 ID 不同（抗碰撞的底线）。"""
    path = _build("/tmp/t_blockid.docx", ["同一段文字", "同一段文字", "同一段文字"])
    bs = _ingest(path)
    assert len({b.block_id for b in bs}) == 3
    assert [b.index for b in bs] == [0, 1, 2]


# ---------- 2. index 跨段落块与表格块连续递增 ----------

def test_index_跨段落与表格块连续():
    bs = _ingest(_params_docx())
    assert [b.index for b in bs] == list(range(len(bs)))


def test_index_表格块参与计数_不与段落块冲突():
    path = _build("/tmp/t_idxtbl.docx", ["第一段", "第二段"])
    d = Document(path)
    d.add_table(1, 2)
    d.add_paragraph("表格后的段落")
    d.save(path)
    bs = _ingest(path)
    assert [b.index for b in bs] == list(range(len(bs)))
    kinds = [b.kind for b in bs]
    assert kinds == [
        BlockKind.PARAGRAPH,
        BlockKind.PARAGRAPH,
        BlockKind.TABLE,
        BlockKind.PARAGRAPH,
    ]


def test_index_跨图片块连续():
    path = _build("/tmp/t_idximg.docx", ["文字段落"])
    d = Document(path)
    p = d.add_paragraph()
    p.add_run().add_picture(BytesIO(_png()), width=Cm(2))
    d.add_paragraph("结尾段落")
    d.save(path)
    bs = _ingest(path)
    assert [b.index for b in bs] == list(range(len(bs)))
    assert bs[1].kind is BlockKind.IMAGE


def test_index_跨复选框内容控件块连续():
    """勾选框复选框段落也是块，index 必须计入。"""
    path = "/tmp/t_idx_cb.docx"
    d = Document()
    d.add_paragraph("前置段落")
    p = d.add_paragraph()
    p._p.append(_checkbox_xml("1"))
    p.add_run("任务清单")
    d.add_paragraph("后置段落")
    d.save(path)
    bs = _ingest(path)
    assert [b.index for b in bs] == [0, 1, 2]
    assert "任务清单" in bs[1].text
    assert bs[1].parent_paragraph is not None


def test_index_在真实模板上连续_块数与预期一致(template_blocks):
    """真实模板 26 个块，index 0..25 连续无跳号。"""
    assert len(template_blocks) == 26
    assert [b.index for b in template_blocks] == list(range(26))


def test_块可按_index_排序_与解析顺序一致():
    bs = _ingest(_params_docx())
    shuffled = list(reversed(bs))
    shuffled.sort(key=lambda b: b.index)
    assert [b.block_id for b in shuffled] == [b.block_id for b in bs]


# ---------- 3. runs 保留：文本、粗体/斜体、中文字体字号 ----------

def test_runs_保留原文逐字文本():
    spec = [
        [
            {"text": "第一章 绪论", "bold": False},
        ],
    ]
    path = _build("/tmp/t_runs_text.docx", spec)
    b = _ingest(path)[0]
    assert b.text == "第一章 绪论"
    assert "".join(r.text for r in b.runs) == b.text


def test_runs_保留粗体():
    path = _build(
        "/tmp/t_runs_bold.docx",
        [[{"text": "这是粗体", "bold": True}, {"text": "后面不是", "bold": None}]],
    )
    b = _ingest(path)[0]
    assert [r.bold for r in b.runs] == [True, None]


def test_runs_保留斜体():
    """斜体 run 保留 italic=True；未设置的 run 为 None（三态不被压成 False）。"""
    path = _build(
        "/tmp/t_runs_italic.docx",
        [[{"text": "斜的", "italic": True}, {"text": "未设置的", "italic": None}]],
    )
    b = _ingest(path)[0]
    assert b.runs[0].italic is True
    assert b.runs[1].italic is None


def test_runs_保留中文字体():
    path = _build(
        "/tmp/t_runs_font.docx", [[{"text": "宋体小四", "font": "宋体", "size": 12}]]
    )
    b = _ingest(path)[0]
    assert b.runs[0].font_cn == "宋体"


def test_runs_保留字号():
    path = _build(
        "/tmp/t_runs_size.docx", [[{"text": "三号字", "font": "宋体", "size": 16}]]
    )
    b = _ingest(path)[0]
    assert b.runs[0].size_pt == 16.0


def test_runs_multirun_kept():
    """多 run 段落的文本不能被拆散或丢弃。"""
    path = _build(
        "/tmp/t_runs_multi.docx",
        [
            [
                {"text": "姓名：", "bold": True},
                {"text": "张三"},
                {"text": "学号：", "bold": True},
                {"text": "20230001"},
            ]
        ],
    )
    b = _ingest(path)[0]
    assert len(b.runs) == 4
    assert b.text == "姓名：张三学号：20230001"


# ---------- 4. 图片 / 表格识别并标 protected ----------

def test_图片被识别为_image_块():
    path = _build("/tmp/t_img.docx", [])
    d = Document(path)
    d.add_paragraph("图片前段落")
    d.add_paragraph().add_run().add_picture(BytesIO(_png()), width=Cm(3))
    d.save(path)
    bs = _ingest(path)
    kinds = [b.kind for b in bs]
    assert BlockKind.IMAGE in kinds
    img = next(b for b in bs if b.kind is BlockKind.IMAGE)
    assert img.asset.kind is BlockKind.IMAGE


def test_图片块被标为_protected():
    path = _build("/tmp/t_img_prot.docx", [])
    d = Document(path)
    d.add_paragraph().add_run().add_picture(BytesIO(_png()), width=Cm(3))
    d.save(path)
    b = next(b for b in _ingest(path) if b.kind is BlockKind.IMAGE)
    assert b.protected is True


def test_普通段落块不是_protected():
    bs = _ingest(_params_docx())
    assert all(b.protected is False for b in bs)


def test_表格块被识别为_table():
    path = _build("/tmp/t_tbl.docx", [])
    d = Document(path)
    t = d.add_table(2, 3)
    t.rows[0].cells[0].text = "表头一"
    d.save(path)
    b = next(b for b in _ingest(path) if b.kind is BlockKind.TABLE)
    assert b.asset.kind is BlockKind.TABLE


def test_表格块被标为_protected():
    path = _build("/tmp/t_tbl_prot.docx", [])
    d = Document(path)
    d.add_table(1, 1)
    d.save(path)
    b = next(b for b in _ingest(path) if b.kind is BlockKind.TABLE)
    assert b.protected is True


def test_块对象_kind_区分段落表格图片():
    path = _build("/tmp/t_kinds.docx", ["正文段落"])
    d = Document(path)
    d.add_table(1, 1)
    d.add_paragraph().add_run().add_picture(BytesIO(_png()), width=Cm(2))
    d.save(path)
    assert {b.kind for b in _ingest(path)} == {
        BlockKind.PARAGRAPH,
        BlockKind.TABLE,
        BlockKind.IMAGE,
    }


# ---------- 表格行列数（真实模板） ----------

def test_模板表格行列数(template_blocks):
    """模板里那个实验信息表：5 行 1 列（见 fixture 实测）。"""
    tbl = next(b for b in template_blocks if b.kind is BlockKind.TABLE)
    assert tbl.asset.rows == 5
    assert tbl.asset.cols == 1


def test_表格行数与列数各自独立():
    path = _build("/tmp/t_dims.docx", [])
    d = Document(path)
    d.add_table(3, 4)
    d.save(path)
    tbl = next(b for b in _ingest(path) if b.kind is BlockKind.TABLE)
    assert (tbl.asset.rows, tbl.asset.cols) == (3, 4)


def test_表格单行单列边界():
    path = _build("/tmp/t_dims1.docx", [])
    d = Document(path)
    d.add_table(1, 1)
    d.save(path)
    tbl = next(b for b in _ingest(path) if b.kind is BlockKind.TABLE)
    assert (tbl.asset.rows, tbl.asset.cols) == (1, 1)


def test_表格文字仍保留():
    path = _build("/tmp/t_dims_text.docx", [])
    d = Document(path)
    t = d.add_table(2, 2)
    for i in range(2):
        for j in range(2):
            t.rows[i].cells[j].text = f"R{i}C{j}"
    d.save(path)
    tbl = next(b for b in _ingest(path) if b.kind is BlockKind.TABLE)
    assert "R0C0" in tbl.text and "R1C1" in tbl.text


# ---------- 5. 图片 rel_id 可读 ----------

def test_图片块有_rel_id():
    path = _build("/tmp/t_relid.docx", [])
    d = Document(path)
    d.add_paragraph().add_run().add_picture(BytesIO(_png()), width=Cm(3))
    d.save(path)
    img = next(b for b in _ingest(path) if b.kind is BlockKind.IMAGE)
    assert img.asset.rel_id is not None
    assert img.asset.rel_id.startswith("rId")


def test_rel_id_可在文档部件里解析():
    path = _build("/tmp/t_relid2.docx", [])
    d = Document(path)
    d.add_paragraph().add_run().add_picture(BytesIO(_png()), width=Cm(3))
    d.save(path)
    img = next(b for b in _ingest(path) if b.kind is BlockKind.IMAGE)
    rid = img.asset.rel_id
    part = d.part.related_parts[rid]
    assert part is not None
    assert len(part.blob) > 0


def test_模板图片块_rel_id_可读(template_blocks):
    """模板头图 rId7 可解析出图片部件。"""
    img = next(b for b in template_blocks if b.kind is BlockKind.IMAGE)
    assert img.asset.rel_id == "rId7"
    d = Document(TEMPLATE)
    assert d.part.related_parts[img.asset.rel_id] is not None


def test_图片块保留宽高_cm():
    path = _build("/tmp/t_relid3.docx", [])
    d = Document(path)
    d.add_paragraph().add_run().add_picture(BytesIO(_png()), width=Cm(4), height=Cm(2))
    d.save(path)
    img = next(b for b in _ingest(path) if b.kind is BlockKind.IMAGE)
    assert img.asset.width_cm == pytest.approx(4.0)
    assert img.asset.height_cm == pytest.approx(2.0)


def test_模板图片块尺寸可读(template_blocks):
    img = next(b for b in template_blocks if b.kind is BlockKind.IMAGE)
    assert img.asset.width_cm == pytest.approx(9.664, abs=1e-3)
    assert img.asset.height_cm == pytest.approx(1.789, abs=1e-3)


# ---------- 6. section_index 取值正确 ----------

def test_单节文档全部块_section_index_为_0():
    bs = _ingest(_params_docx())
    assert all(b.section_index == 0 for b in bs)


def test_模板封面区块_section_index_为_0(template_blocks):
    cover = [b for b in template_blocks if b.index <= 18]
    assert cover and all(b.section_index == 0 for b in cover)


def test_模板正文区块_section_index_为_1(template_blocks):
    body = [b for b in template_blocks if b.index >= 19]
    assert body and all(b.section_index == 1 for b in body)


def test_模板两种_section_index_都出现(template_blocks):
    assert {b.section_index for b in template_blocks} == {0, 1}


def test_分节后段落_section_index_递增():
    path = _build("/tmp/t_sec.docx", ["封面文字"])
    d = Document(path)
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("正文第一段")
    d.save(path)
    bs = _ingest(path)
    sections = {b.text.strip(): b.section_index for b in bs if b.text.strip()}
    assert sections["封面文字"] == 0
    assert sections["正文第一段"] == 1


def test_三节文档_section_index_依次_0_1_2():
    path = _build("/tmp/t_sec3.docx", ["第一节"])
    d = Document(path)
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("第二节")
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("第三节")
    d.save(path)
    bs = [b for b in _ingest(path) if b.text.strip()]
    assert [b.section_index for b in bs] == [0, 1, 2]


def test_section_index_不依赖分节符段落自身是否为块():
    """分节符段落的 section_index 仍属于前一节（它自己还没换节）。"""
    path = _build("/tmp/t_secbreak.docx", ["第一节段落"])
    d = Document(path)
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("第二节段落")
    d.save(path)
    bs = _ingest(path)
    texts = {b.text.strip(): b.section_index for b in bs}
    assert texts["第一节段落"] == 0
    assert texts["第二节段落"] == 1


# ---------- 其它导入层不变量 ----------

def test_块顺序与文档物理顺序一致():
    """解析顺序必须是文档顺序（后续依赖 b[i] 相邻关系）。"""
    path = _build("/tmp/t_order.docx", ["第一段", "第二段", "第三段"])
    bs = _ingest(path)
    assert [b.text for b in bs] == ["第一段", "第二段", "第三段"]


def test_段落块保留_xml_element_与_parent_paragraph():
    bs = _ingest(_params_docx())
    for b in bs:
        assert b.xml_element is not None
        assert b.parent_paragraph is not None
        assert b.parent_paragraph._element is b.xml_element


def test_表格块_xml_element_是表格节点():
    path = _build("/tmp/t_tblxml.docx", [])
    d = Document(path)
    d.add_table(2, 2)
    d.save(path)
    tbl = next(b for b in _ingest(path) if b.kind is BlockKind.TABLE)
    assert tbl.xml_element.tag == qn("w:tbl")
    assert tbl.parent_paragraph is None


def test_空文档解析返回空列表():
    path = _build("/tmp/t_empty.docx", [])
    assert _ingest(path) == []


def test_纯段落文档所有块_kind_为段落():
    bs = _ingest(_params_docx())
    assert {b.kind for b in bs} == {BlockKind.PARAGRAPH}


def test_块文本与_runs_拼接一致():
    """text 必须等于 runs 的拼接（编辑器按 run 改字后能对上）。"""
    bs = _ingest(_params_docx())
    for b in bs:
        assert b.text == "".join(r.text for r in b.runs)


# ---------- 内部工具 ----------

def _png() -> bytes:
    return base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
    )


def _checkbox_xml(tag: str):
    """构造一个 w:sdt 复选框内容控件元素（返回 lxml element，而非 bytes）。"""
    from lxml import etree

    return etree.fromstring(
        f'<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:sdtPr><w14:checkbox '
        f'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        f'<w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>'
        f'<w:sdtContent><w:r><w:t>{tag}</w:t></w:r></w:sdtContent></w:sdt>',
        parser=None,
    )


def _ingest(path):
    return ingest(path)