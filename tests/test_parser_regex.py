"""正则规则表分类测试（tf/parser/regex_parser.py）。

覆盖：
- 规则表逐条：摘要 / 关键词 / 目录 / 参考文献 / 致谢附录 / 章标题 /
  三级标题 / 二级标题 / 图表标题 / 参考文献条目 / 封面标题 / 封面信息行
- 规则顺序敏感性：三级先于二级；具体标题规则先于封面兜底规则
- 上下文守卫：in_references 置位与重置、in_abstract、prev_heading_level
- 启发式兜底：粗体 + 短 + 不以句末标点收尾
- 受保护块直通（IMAGE/TABLE/FORMULA，1.0/PRESET）与 needs_llm 联动

约定：
- 所有预期值均按源码实际行为逐条核对，不凭规则字面推测。
- `test_规则表每条规则都有对应用例` 会随 RULES 变动而失败，用于强制补测。
"""
from __future__ import annotations

import base64
from io import BytesIO

import pytest
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Cm, Pt

from tf.core.enums import BlockKind, BlockSource, BlockType
from tf.core.models import Block
from tf.ingestion.engine import ingest
from tf.parser.regex_parser import RULES, Context, _classify_block, classify

_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


# --------------------------------------------------------------------------
# 辅助
# --------------------------------------------------------------------------
def _build(tmp_path, specs, name="docx"):
    """按 specs 生成 docx；spec 支持 text / bold / center / size_pt / new_section。"""
    path = tmp_path / f"{name}.docx"
    doc = Document()
    for spec in specs:
        if spec.get("new_section"):
            doc.add_section(WD_SECTION.NEW_PAGE)
        para = doc.add_paragraph()
        run = para.add_run(spec["text"])
        if spec.get("bold"):
            run.bold = True
        if spec.get("size_pt"):
            run.font.size = Pt(spec["size_pt"])
        if spec.get("center"):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(str(path))
    return str(path)


def _classify(tmp_path, specs):
    """ingest + classify 全流程（与 cli parse 子命令一致）。"""
    blocks = ingest(_build(tmp_path, specs))
    return classify(blocks)


def _only(tmp_path, text, **opts):
    """单段文档，返回唯一 Block。"""
    blocks = _classify(tmp_path, [{"text": text, **opts}])
    assert len(blocks) == 1
    return blocks[0]


def _tag(block):
    return block.semantic_tag


# --------------------------------------------------------------------------
# 规则表逐条样例
# (specs, 待检块下标, 期望 block_type, 期望 confidence)
# --------------------------------------------------------------------------
_PROBES = [
    # 0 摘要标题：允许字间空白
    ([{"text": "摘要"}], 0, BlockType.ABSTRACT_TITLE, 0.99),
    ([{"text": "摘 要"}], 0, BlockType.ABSTRACT_TITLE, 0.99),
    ([{"text": "摘  要"}], 0, BlockType.ABSTRACT_TITLE, 0.99),
    # 1 关键词：全角/半角冒号，或「关键字」
    ([{"text": "关键词：深度学习；图像"}], 0, BlockType.KEYWORDS, 0.98),
    ([{"text": "关键字：甲"}], 0, BlockType.KEYWORDS, 0.98),
    # 2 目录
    ([{"text": "目录"}], 0, BlockType.TOC, 0.99),
    ([{"text": "目 录"}], 0, BlockType.TOC, 0.99),
    # 3 参考文献标题
    ([{"text": "参考文献"}], 0, BlockType.REFERENCE_TITLE, 0.99),
    ([{"text": "参 考 文 献"}], 0, BlockType.REFERENCE_TITLE, 0.99),
    # 4/5 致谢、附录（两条规则同 tag 同 confidence）
    ([{"text": "致谢"}], 0, BlockType.APPENDIX_TITLE, 0.90),
    ([{"text": "致 谢"}], 0, BlockType.APPENDIX_TITLE, 0.90),
    ([{"text": "附录A"}], 0, BlockType.APPENDIX_TITLE, 0.90),
    ([{"text": "附录"}], 0, BlockType.APPENDIX_TITLE, 0.90),
    # 6 章标题：第X章，X 限中文数字（含「两」）
    ([{"text": "第一章 绪论"}], 0, BlockType.HEADING_1, 0.98),
    ([{"text": "第 三 章 方法"}], 0, BlockType.HEADING_1, 0.98),
    ([{"text": "第十一章 总结"}], 0, BlockType.HEADING_1, 0.98),
    # 7 三级标题：1.2.3
    ([{"text": "1.2.3 小节"}], 0, BlockType.HEADING_3, 0.96),
    ([{"text": "1.2.3"}], 0, BlockType.HEADING_3, 0.96),
    ([{"text": "1.2.3.4 深"}], 0, BlockType.HEADING_3, 0.96),
    # 8 二级标题：1.2（后随空白、顿号或字符串结束）
    ([{"text": "1.2 小节"}], 0, BlockType.HEADING_2, 0.95),
    ([{"text": "1.2"}], 0, BlockType.HEADING_2, 0.95),
    ([{"text": "1.2、甲"}], 0, BlockType.HEADING_2, 0.95),
    ([{"text": "12.34 号"}], 0, BlockType.HEADING_2, 0.95),
    # 9 图/表标题
    ([{"text": "图1 系统框图"}], 0, BlockType.CAPTION, 0.97),
    ([{"text": "图2"}], 0, BlockType.CAPTION, 0.97),
    ([{"text": "图 2 框图"}], 0, BlockType.CAPTION, 0.97),
    ([{"text": "表1 数据"}], 0, BlockType.CAPTION, 0.97),
    ([{"text": "表 2 数据"}], 0, BlockType.CAPTION, 0.97),
    # 10 参考文献条目：需 in_references 守卫（前置「参考文献」标题）
    (
        [{"text": "参考文献"}, {"text": "[1] 张三. 论文[J]. 2020."}],
        1,
        BlockType.REFERENCE,
        0.98,
    ),
    ([{"text": "参考文献"}, {"text": "[12] 李四"}], 1, BlockType.REFERENCE, 0.98),
    # 11 封面标题：section 0 + 短 + 居中 + 关键词
    ([{"text": "本科毕业论文", "center": True}], 0, BlockType.COVER_TITLE, 0.90),
    ([{"text": "实验报告", "center": True}], 0, BlockType.COVER_TITLE, 0.90),
    ([{"text": "毕 业 设 计 （ 论 文 ）", "center": True}], 0, BlockType.COVER_TITLE, 0.90),
    ([{"text": "图像处理课程设计报告", "center": True}], 0, BlockType.COVER_TITLE, 0.90),
    # 12 封面信息行
    ([{"text": "学号：20230001"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "姓名: 张三"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "课程：图像处理"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "学院：光电"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "专业：计算机"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "班级：一班"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "指导教师：李四"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "成绩：优"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "评语：好"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "实验名称：测光"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "实验地点：A301"}], 0, BlockType.COVER_META, 0.92),
    ([{"text": "实验日期：2026-01-01"}], 0, BlockType.COVER_META, 0.92),
]


def _probe_ids():
    return [f"{specs[index]['text']}→{tag.name}" for specs, index, tag, _ in _PROBES]


@pytest.mark.parametrize(
    "specs,index,expected_tag,expected_confidence",
    _PROBES,
    ids=_probe_ids(),
)
def test_规则表样例分类正确(tmp_path, specs, index, expected_tag, expected_confidence):
    blocks = _classify(tmp_path, specs)
    block = blocks[index]
    assert block.semantic_tag == expected_tag
    assert block.confidence == pytest.approx(expected_confidence)
    # 走规则表命中的块，来源一律 REGEX
    assert block.source == BlockSource.REGEX


def test_规则表每条规则都有对应用例():
    """RULES 新增/调整时必须同步补用例，避免规则静默无覆盖。"""
    covered = {(tag, conf) for _, _, tag, conf in _PROBES}
    for rule in RULES:
        assert (rule.block_type, rule.confidence) in covered, (
            f"规则 {rule.pattern!r} → {rule.block_type}/{rule.confidence} 缺少用例"
        )


# --------------------------------------------------------------------------
# 规则顺序敏感性
# --------------------------------------------------------------------------
def test_规则顺序_三级标题先于二级标题(tmp_path):
    """1.2.3 必须判为 HEADING_3，而非被二级规则截胡。"""
    idx_3 = next(i for i, r in enumerate(RULES) if r.block_type == BlockType.HEADING_3)
    idx_2 = next(i for i, r in enumerate(RULES) if r.block_type == BlockType.HEADING_2)
    assert idx_3 < idx_2, "三级标题规则必须排在二级标题规则之前"

    block = _only(tmp_path, "1.2.3 小节")
    assert _tag(block) == BlockType.HEADING_3


def test_规则顺序_具体标题规则先于封面兜底(tmp_path):
    """封面规则是 catch-all `.*`，必须排在具名标题规则之后。"""
    idx_cover = next(i for i, r in enumerate(RULES) if r.block_type == BlockType.COVER_TITLE)
    idx_abs = next(
        i for i, r in enumerate(RULES) if r.block_type == BlockType.ABSTRACT_TITLE
    )
    assert idx_abs < idx_cover

    # 居中 + 关键词，但正文是「参考文献」：应判标题而非封面
    blocks = _classify(tmp_path, [{"text": "参考文献", "center": True}])
    assert _tag(blocks[0]) == BlockType.REFERENCE_TITLE
    assert blocks[0].confidence == pytest.approx(0.99)

    blocks = _classify(tmp_path, [{"text": "摘要", "center": True}])
    assert _tag(blocks[0]) == BlockType.ABSTRACT_TITLE

    blocks = _classify(tmp_path, [{"text": "目 录", "center": True}])
    assert _tag(blocks[0]) == BlockType.TOC

    blocks = _classify(tmp_path, [{"text": "附录A", "center": True}])
    assert _tag(blocks[0]) == BlockType.APPENDIX_TITLE


# --------------------------------------------------------------------------
# 未命中任何规则 → 正文兜底
# --------------------------------------------------------------------------
@pytest.mark.parametrize(
    "text",
    [
        "摘要：本文研究了图像分类方法。",   # 摘要有尾随内容，非整段标题
        "关键词 深度学习",                 # 缺冒号
        "目录一",                          # 非整段
        "第1章 绪论",                      # 阿拉伯数字不在中文数字集
        "第X章 绪论",                      # 拉丁字母不在中文数字集
        "第章 绪论",                       # 缺中文数字
        "图形学",                          # 「图」后非数字
        "表结构",                          # 「表」后非数字
        "课程名称：图像处理",              # 「课程」后非冒号
        "课程名称",                        # 无冒号
        "课程名称 图像处理",               # 无冒号
        "这是一段普通的正文内容。",         # 普通正文
        "短句",                            # 非粗体短句
        "本科毕业论文",                    # 封面关键词但缺居中
        "光电学院简介",                    # 居中短句但无封面关键词
    ],
)
def test_未命中规则_落为正文(tmp_path, text):
    block = _only(tmp_path, text, center=text == "光电学院简介")
    assert _tag(block) == BlockType.BODY
    assert block.confidence == pytest.approx(0.60)
    assert block.source == BlockSource.REGEX


def test_封面标题_长度上限25字(tmp_path):
    """封面标题守卫要求 len(text.strip()) <= 25。"""
    ok = "毕业" + "生" * 21 + "报告"
    too_long = "毕业" + "生" * 22 + "报告"
    assert len(ok) == 25
    assert len(too_long) == 26

    assert _tag(_only(tmp_path, ok, center=True)) == BlockType.COVER_TITLE
    assert _tag(_only(tmp_path, too_long, center=True)) == BlockType.BODY


@pytest.mark.xfail(
    reason="已知缺陷：Context.section_index 从未从 block 同步，守卫 c.section_index == 0 恒为真；"
    "第 1 节及以后的封面标题仍会被误判为 COVER_TITLE。已上报 Lead。",
    strict=False,
)
def test_封面标题_仅第0节生效(tmp_path):
    """按规则意图：仅封面所在 section 0 才允许判 COVER_TITLE。"""
    blocks = _classify(
        tmp_path,
        [
            {"text": "正文段落"},
            {"text": "本科毕业论文", "center": True, "new_section": True},
        ],
    )
    assert blocks[1].section_index == 1
    assert _tag(blocks[1]) == BlockType.BODY


# --------------------------------------------------------------------------
# 上下文：in_references
# --------------------------------------------------------------------------
def test_in_references_由参考文献标题置位(tmp_path):
    blocks = _classify(
        tmp_path, [{"text": "参考文献"}, {"text": "[1] 张三. 论文[J]. 2020."}]
    )
    assert _tag(blocks[0]) == BlockType.REFERENCE_TITLE
    assert _tag(blocks[1]) == BlockType.REFERENCE


def test_in_references_未置位时编号条目判正文(tmp_path):
    """guard 不满足 → 继续往下试，最终落正文。"""
    for text in ("[1] 张三", "[12] 李四", "[a] 非数字"):
        block = _only(tmp_path, text)
        assert _tag(block) == BlockType.BODY
        assert block.source == BlockSource.REGEX


def test_in_references_被章标题重置(tmp_path):
    """章标题 → 一级标题会清掉 in_references，其后的编号行回到正文。"""
    blocks = _classify(
        tmp_path,
        [
            {"text": "参考文献"},
            {"text": "[1] 张三"},
            {"text": "第一章 绪论"},
            {"text": "[2] 李四"},
        ],
    )
    assert [_tag(b) for b in blocks] == [
        BlockType.REFERENCE_TITLE,
        BlockType.REFERENCE,
        BlockType.HEADING_1,
        BlockType.BODY,
    ]


def test_in_references_不被二级标题重置(tmp_path):
    """二级标题只更新 prev_heading_level，不清 in_references。"""
    blocks = _classify(
        tmp_path,
        [
            {"text": "参考文献"},
            {"text": "[1] 张三"},
            {"text": "1.2 小节"},
            {"text": "[2] 李四"},
        ],
    )
    assert [_tag(b) for b in blocks] == [
        BlockType.REFERENCE_TITLE,
        BlockType.REFERENCE,
        BlockType.HEADING_2,
        BlockType.REFERENCE,
    ]


def test_附录标题不设in_references(tmp_path):
    blocks = _classify(tmp_path, [{"text": "附录A"}, {"text": "[1] 张三"}])
    assert _tag(blocks[0]) == BlockType.APPENDIX_TITLE
    assert _tag(blocks[1]) == BlockType.BODY


# --------------------------------------------------------------------------
# 上下文：_classify_block 白盒
# --------------------------------------------------------------------------
def test_上下文_摘要置in_abstract():
    ctx = Context()
    block = Block(block_id="p_0", index=0, kind=BlockKind.PARAGRAPH, text="摘要")
    _classify_block(block, ctx)
    assert block.semantic_tag == BlockType.ABSTRACT_TITLE
    assert ctx.in_abstract is True
    assert ctx.in_references is False


def test_上下文_章标题置in_abstract为假并记level():
    ctx = Context(in_abstract=True, in_references=True)
    block = Block(block_id="p_0", index=0, kind=BlockKind.PARAGRAPH, text="第一章 绪论")
    _classify_block(block, ctx)
    assert block.semantic_tag == BlockType.HEADING_1
    assert ctx.prev_heading_level == 1
    assert ctx.in_abstract is False
    assert ctx.in_references is False


@pytest.mark.parametrize(
    "text,level",
    [("1.2 小节", 2), ("1.2.3 小节", 3)],
)
def test_上下文_prev_heading_level(text, level):
    ctx = Context()
    block = Block(block_id="p_0", index=0, kind=BlockKind.PARAGRAPH, text=text)
    _classify_block(block, ctx)
    assert ctx.prev_heading_level == level


def test_上下文_参考文献标题置in_references():
    ctx = Context()
    block = Block(block_id="p_0", index=0, kind=BlockKind.PARAGRAPH, text="参考文献")
    _classify_block(block, ctx)
    assert ctx.in_references is True


# --------------------------------------------------------------------------
# 启发式兜底
# --------------------------------------------------------------------------
def test_粗体短句_判一级标题_低置信(tmp_path):
    block = _only(tmp_path, "这是粗体短句", bold=True)
    assert _tag(block) == BlockType.HEADING_1
    assert block.confidence == pytest.approx(0.55)
    assert block.source == BlockSource.REGEX


def test_粗体启发式_长度上限30字(tmp_path):
    """len(text.strip()) <= 30 才算「短」。"""
    short = "粗" * 30
    long_ = "粗" * 31
    assert _tag(_only(tmp_path, short, bold=True)) == BlockType.HEADING_1
    assert _tag(_only(tmp_path, long_, bold=True)) == BlockType.BODY


@pytest.mark.parametrize("tail", ["。", "；", "，", "."])
def test_粗体启发式_句末标点排除(tmp_path, tail):
    text = "粗体结尾" + tail
    assert _tag(_only(tmp_path, text, bold=True)) == BlockType.BODY


def test_非粗体短句_不触发启发式(tmp_path):
    assert _tag(_only(tmp_path, "短句")) == BlockType.BODY


@pytest.mark.xfail(
    reason="已知缺陷：engine._extract_runs 以 `w:b` 元素是否存在判定粗体，忽略 w:val=\"0\"；"
    "python-docx 写入的显式 bold=False 被读成 True，从而误触发粗体标题启发式。已上报 Lead。",
    strict=False,
)
def test_显式非粗体_不应触发粗体启发式(tmp_path):
    """显式 bold=False 应等价于非粗体，落正文。"""
    path = tmp_path / "nobold.docx"
    doc = Document()
    run = doc.add_paragraph().add_run("这是显式非粗体短句")
    run.bold = False
    doc.save(str(path))

    block = classify(ingest(str(path)))[0]
    assert block.runs[0].bold is False
    assert _tag(block) == BlockType.BODY


# --------------------------------------------------------------------------
# 空段落与受保护块
# --------------------------------------------------------------------------
@pytest.mark.parametrize("text", ["", "   ", "\t"])
def test_空段落_打UNKNOWN且不进LLM(tmp_path, text):
    block = _only(tmp_path, text)
    assert _tag(block) == BlockType.UNKNOWN
    assert block.confidence == pytest.approx(0.0)
    assert block.source is None
    assert block.needs_llm() is False


def test_受保护图片块_直通IMAGE(tmp_path):
    path = tmp_path / "img.docx"
    doc = Document()
    doc.add_paragraph().add_run().add_picture(BytesIO(_PNG), width=Cm(2))
    doc.save(str(path))

    block = classify(ingest(str(path)))[0]
    assert block.kind == BlockKind.IMAGE
    assert _tag(block) == BlockType.IMAGE
    assert block.confidence == pytest.approx(1.0)
    assert block.source == BlockSource.PRESET
    assert block.protected is True


def test_受保护公式块_直通FORMULA(tmp_path):
    path = tmp_path / "formula.docx"
    doc = Document()
    para = doc.add_paragraph()
    para._p.append(
        parse_xml(
            f'<m:oMathPara {nsdecls("m")}>'
            "<m:oMath><m:r><m:t>x^2</m:t></m:r></m:oMath>"
            "</m:oMathPara>"
        )
    )
    doc.save(str(path))

    block = classify(ingest(str(path)))[0]
    assert block.kind == BlockKind.FORMULA
    assert _tag(block) == BlockType.FORMULA
    assert block.confidence == pytest.approx(1.0)
    assert block.source == BlockSource.PRESET
    assert block.protected is True


def test_受保护表格块_直通TABLE(tmp_path):
    path = tmp_path / "table.docx"
    doc = Document()
    doc.add_paragraph("表格前的说明")
    doc.add_table(rows=2, cols=2)
    doc.save(str(path))

    blocks = classify(ingest(str(path)))
    table = [b for b in blocks if b.kind == BlockKind.TABLE]
    assert len(table) == 1
    assert _tag(table[0]) == BlockType.TABLE
    assert table[0].confidence == pytest.approx(1.0)
    assert table[0].source == BlockSource.PRESET
    assert table[0].protected is True


def test_非段落块_不被打标():
    """SECTION_BREAK 等非段落块保持 semantic_tag=None，不落任何规则。"""
    block = Block(block_id="p_0", index=0, kind=BlockKind.SECTION_BREAK, text="x")
    classify([block])
    assert block.semantic_tag is None
    assert block.confidence == pytest.approx(0.0)
    assert block.source is None


# --------------------------------------------------------------------------
# 返回值语义与 needs_llm 联动
# --------------------------------------------------------------------------
def test_classify_原地返回同一列表():
    blocks = [Block(block_id="p_0", index=0, kind=BlockKind.PARAGRAPH, text="摘要")]
    result = classify(blocks)
    assert result is blocks
    assert blocks[0].semantic_tag == BlockType.ABSTRACT_TITLE


def test_needs_llm_仅低置信正则段落为真(tmp_path):
    blocks = _classify(
        tmp_path,
        [
            {"text": "摘要"},              # 0.99 REGEX 段落 → False
            {"text": "这是一段正文。"},      # 0.60 REGEX 段落 → True
            {"text": "这是粗体短句", "bold": True},  # 0.55 → True
            {"text": ""},                 # UNKNOWN → False
        ],
    )
    assert [b.needs_llm() for b in blocks] == [False, True, True, False]