"""核心测试：tf.core 数据模型 + 正则解析器 + 规则表。

运行：
    PYTHONPATH=/home/dsh/workspace/.tf-deps python -m pytest tests/test_core.py -v

覆盖：
- Block / RunInfo / AssetInfo 数据模型
- 图片块、表格块、公式块识别
- 标题层级与自动编号（正则侧）
- 封面识别
- rules.json 加载与默认样式
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from tf.core.enums import BlockKind, BlockSource, BlockType
from tf.core.models import AssetInfo, Block, RunInfo
from tf.parser.regex_parser import classify
from tf.template.rules import load_rules, rules_from_dict


# ---------------------------------------------------------------- 数据模型

def test_block_defaults():
    b = Block(block_id="b0001_abc", index=0, kind=BlockKind.PARAGRAPH, runs=[])
    assert b.text == ""
    assert b.runs == []
    assert b.asset is None
    assert b.semantic_tag is None
    assert b.confidence == 0.0
    assert b.protected is False


def test_block_requires_id_index_kind():
    with pytest.raises(TypeError):
        Block(runs=[])  # block_id/index/kind 均必填


def test_runs_flatten():
    b = Block(
        block_id="b0002_x", index=1, kind=BlockKind.PARAGRAPH,
        text="hello world",
        runs=[RunInfo(text="hello ", bold=True), RunInfo(text="world")],
    )
    assert b.text == "hello world"
    assert len(b.runs) == 2
    assert b.runs[0].bold is True
    assert b.runs[1].text == "world"


def test_asset_info():
    a = AssetInfo(kind=BlockKind.IMAGE, width_cm=10.0, height_cm=5.0)
    assert a.width_cm == 10.0 and a.height_cm == 5.0
    b = Block(block_id="b0003_img", index=2, kind=BlockKind.IMAGE, asset=a)
    assert b.kind == BlockKind.IMAGE
    assert b.asset.width_cm == 10.0


def test_asset_optional():
    # kind 必填，其余可选
    a = AssetInfo(kind=BlockKind.TABLE, rows=2, cols=3)
    assert a.rows == 2 and a.cols == 3
    assert Block(block_id="b0004_x", index=3, kind=BlockKind.PARAGRAPH).asset is None


# ---------------------------------------------------------------- 正则分类

_counter = [0]


def _para(text, **kw):
    _counter[0] += 1
    kw.setdefault("index", _counter[0])
    kw.setdefault("block_id", f"b{kw['index']:04d}_t")
    kw.setdefault("kind", BlockKind.PARAGRAPH)
    return Block(runs=[RunInfo(text=text)], text=text, **kw)


def test_heading_numbering():
    blocks = [
        _para("第一章 绪论"),
        _para("1.1 研究背景"),
        _para("1.1.1 详细"),
        _para("正文一段。"),
    ]
    classify(blocks)
    assert blocks[0].semantic_tag == BlockType.HEADING_1
    assert blocks[1].semantic_tag == BlockType.HEADING_2
    assert blocks[2].semantic_tag == BlockType.HEADING_3
    assert blocks[3].semantic_tag == BlockType.BODY


def test_abstract_and_keywords():
    blocks = [
        _para("摘 要"),
        _para("这是摘要内容。"),
        _para("关键词：论文；格式；工具"),
    ]
    classify(blocks)
    assert blocks[0].semantic_tag == BlockType.ABSTRACT_TITLE
    assert blocks[1].semantic_tag == BlockType.BODY  # 摘要正文在标题后
    assert blocks[2].semantic_tag == BlockType.KEYWORDS


def test_cover_title():
    # 封面标题要求居中（_is_centered 检查 parent_paragraph.alignment）
    class _Align:
        name = "CENTER"

    class _CenteredPara:
        alignment = _Align()

    b0 = Block(block_id="b0000_ct", index=0, kind=BlockKind.PARAGRAPH,
               text="本科毕业论文", runs=[RunInfo(text="本科毕业论文")],
               parent_paragraph=_CenteredPara())
    b1 = _para("姓名：张三")
    classify([b0, b1])
    assert b0.semantic_tag == BlockType.COVER_TITLE
    assert b1.semantic_tag == BlockType.COVER_META


def test_caption_and_reference():
    blocks = [
        _para("图 1 系统架构图"),
        _para("参考文献"),
        _para("[1] 张三. 论文题目[J]. 期刊, 2020."),
    ]
    classify(blocks)
    assert blocks[0].semantic_tag == BlockType.CAPTION
    assert blocks[1].semantic_tag == BlockType.REFERENCE_TITLE
    assert blocks[2].semantic_tag == BlockType.REFERENCE


def test_non_paragraph_kind():
    b_img = Block(block_id="b0000_i", index=0, kind=BlockKind.IMAGE,
                  asset=AssetInfo(kind=BlockKind.IMAGE, width_cm=3.0, height_cm=2.0))
    b_tbl = Block(block_id="b0001_t", index=1, kind=BlockKind.TABLE)
    b_fml = Block(block_id="b0002_f", index=2, kind=BlockKind.FORMULA,
                  runs=[RunInfo(text="a^2+b^2=c^2")])
    classify([b_img, b_tbl, b_fml])
    assert b_img.semantic_tag == BlockType.IMAGE
    assert b_tbl.semantic_tag == BlockType.TABLE
    assert b_fml.semantic_tag == BlockType.FORMULA


# ---------------------------------------------------------------- 规则表

def test_load_default_rules():
    rules = load_rules(Path(__file__).parent.parent / "tf" / "parser" / "rules_min.json")
    assert rules.default.size_pt == 12.0
    assert BlockType.BODY in rules.styles
    assert rules.styles[BlockType.BODY].first_line_indent_chars == 2.0
    assert rules.heading_cfg.get("levels") == 3


def test_rules_from_dict_custom():
    r = rules_from_dict({
        "default": {"font_cn": "宋体", "size_pt": 12},
        "styles": {"body": {"size_pt": 14}},
    })
    assert r.styles[BlockType.BODY].size_pt == 14
    assert r.default.font_cn == "宋体"


def test_rules_align():
    r = rules_from_dict({"default": {"align": "center"}, "styles": {}})
    assert r.default.align.value == "center"


def test_rules_resolve_fallback():
    r = rules_from_dict({
        "default": {"size_pt": 12},
        "styles": {"body": {"size_pt": 14}},
    })
    assert r.resolve(BlockType.BODY).size_pt == 14
    # 未定义的类型回落到 default
    assert r.resolve(BlockType.HEADING_1).size_pt == 12


# ------------------------------------------------- 字段级级联（docs/06 §2.2）

def test_resolve_partial_override_keeps_unset_fields():
    """specific 只覆盖显式给出的字段，其余从 default 继承。"""
    r = rules_from_dict({
        "default": {"font_cn": "宋体", "font_en": "Times New Roman", "size_pt": 12},
        "styles": {"body": {"size_pt": 10}},
    })
    body = r.resolve(BlockType.BODY)
    assert body.size_pt == 10
    assert body.font_cn == "宋体"          # 未指定 → 继承
    assert body.font_en == "Times New Roman"

    # 未定义的类型返回 default 本身（不被任何样式污染）
    head = r.resolve(BlockType.HEADING_1)
    assert head.font_cn == "宋体"
    assert head.size_pt == 12


def test_resolve_partial_override_align_and_indent():
    """典型规则表：标题只给字号/加粗，对齐与缩进继承 default。"""
    r = rules_from_dict({
        "default": {"font_cn": "宋体", "size_pt": 12, "align": "justify",
                    "line_spacing": 1.5, "first_line_indent_chars": 2.0},
        "styles": {"heading_2": {"font_cn": "黑体", "size_pt": 14, "bold": True}},
    })
    h2 = r.resolve(BlockType.HEADING_2)
    assert (h2.font_cn, h2.size_pt, h2.bold) == ("黑体", 14.0, True)
    assert h2.line_spacing == 1.5
    assert h2.first_line_indent_chars == 2.0
    assert h2.align.value == "justify"


def test_resolve_default_not_mutated():
    """resolve 不得就地修改 RuleSet.default。"""
    r = rules_from_dict({
        "default": {"font_cn": "宋体"},
        "styles": {"body": {"size_pt": 10}},
    })
    r.resolve(BlockType.BODY)
    assert r.default.size_pt is None
    assert r.default.font_cn == "宋体"


def test_rules_defaults_key_alias():
    """docs/06 §1.3 的 `defaults` 写法与内置规则文件的 `default` 等价。"""
    a = rules_from_dict({"default": {"font_cn": "宋体", "size_pt": 12}})
    b = rules_from_dict({"defaults": {"font_cn": "宋体", "size_pt": 12}})
    assert b.default.font_cn == "宋体"
    assert b.default.size_pt == 12
    assert b.default.to_dict() == a.default.to_dict()


def test_rules_unknown_toplevel_key_ignored():
    """拼错的顶层键不阻断加载（只告警）。"""
    r = rules_from_dict({"defualt": {"font_cn": "宋体"}, "styles": {}})
    assert r.default.font_cn is None
    assert r.resolve(BlockType.BODY).font_cn is None