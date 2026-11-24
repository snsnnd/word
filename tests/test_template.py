"""`tf/template` 层测试（docs/07 §2 测试清单 · test_template）。

覆盖：
- `load_template(master_path, rules_path)` → TemplateSpec 契约（master_path / rules /
  style_snapshot / warnings）
- `_check_master` 母版告警（命名样式过少）
- `extract_master_styles` 页面尺寸、方向、命名样式、分节
- `load_rules` 对随包分发的 `tf/parser/rules_min.json` 的解析结果（default /
  resolve(BODY) / resolve(REFERENCE) / heading.levels）
- rules.json 边界：未知顶层键、未知块类型名（均只告警不阻断）、`defaults` 别名
- `validate_rules` 在缺少 `jsonschema` 时静默降级（不装包，用 sys.modules 模拟缺失）

约定：
- 不修改 `tf/` 源码、不安装依赖；测试只在 tmp_path 下造临时母版。
- 页面尺寸用范围/近似断言（不同 Word 版本的 EMU 换算会有微小差异）。
"""
from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

import docx
import pytest
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:  # 允许直接 `python3 -m pytest tests/` 运行
    sys.path.insert(0, str(ROOT))

from tf.core.enums import BlockType  # noqa: E402
from tf.template.engine import _check_master, load_template  # noqa: E402
from tf.template.master import extract_master_styles  # noqa: E402
from tf.template.rules import (  # noqa: E402
    default_rules_path,
    load_rules,
    rules_from_dict,
    validate_rules,
)

#: 仓库自带母版（真实模板，命名样式少、以直接格式化为主）
FIXTURE_MASTER = ROOT / "tests" / "fixtures" / "source_template.docx"
#: 随包分发的规则表
RULES_MIN = ROOT / "tf" / "parser" / "rules_min.json"

#: _style_to_dict 始终会输出的 11 个字段（缺失项为 None）
STYLE_FIELDS = {
    "font_cn", "font_en", "size_pt", "bold", "italic", "align",
    "line_spacing", "first_line_indent_chars", "hanging_indent_chars",
    "space_before_pt", "space_after_pt",
}

#: 触发「母版几乎只使用默认样式」告警的阈值（engine.py:72 `<= 2`）
FEW_STYLE_THRESHOLD = 2


# --------------------------------------------------------------------------- #
# 辅助：构造最小母版
# --------------------------------------------------------------------------- #
def _prune_styles(document, keep_names: set[str]) -> None:
    """从 styles.xml 中删掉名字不在 keep_names 的样式，制造「命名样式极少」的母版。

    python-docx 的内置样式名在 XML 里是小写（如 `heading 1`），
    而 `style.name` 读出来是 `Heading 1`，故这里按 XML 名过滤。
    """
    styles_el = document.styles.element
    for style in list(styles_el):
        if style.tag != qn("w:style"):
            continue
        name_el = style.find(qn("w:name"))
        name = name_el.get(qn("w:val")) if name_el is not None else None
        if name not in keep_names:
            styles_el.remove(style)


def _make_minimal_master(path: Path) -> Path:
    """造一个只保留 Normal + Heading 1 两个段落样式的母版。"""
    document = docx.Document()
    document.add_heading("Master Heading", level=1)
    document.add_paragraph("MASTER-PLACEHOLDER")
    _prune_styles(document, {"Normal", "heading 1"})
    document.save(str(path))
    return path


# --------------------------------------------------------------------------- #
# load_template → TemplateSpec
# --------------------------------------------------------------------------- #
def test_load_template_returns_spec_with_master_path(tmp_path):
    """load_template 的 master_path 必须是入参的字符串形式（供渲染层回读母版）。"""
    master = _make_minimal_master(tmp_path / "master.docx")

    spec = load_template(master, RULES_MIN)

    assert spec.master_path == str(master)
    assert spec.name == "custom"
    assert spec.version == "0.1.0"
    # rules 已解析成 RuleSet，不是 None
    assert spec.rules is not None
    assert spec.rules.default.font_cn == "宋体"


def test_load_template_snapshot_contains_paragraph_styles(tmp_path):
    """style_snapshot 必须带 paragraph_styles（渲染层靠它做未配置项兜底）。"""
    master = _make_minimal_master(tmp_path / "master.docx")

    spec = load_template(master, RULES_MIN)

    assert isinstance(spec.style_snapshot, dict)
    para_styles = spec.style_snapshot["paragraph_styles"]
    assert isinstance(para_styles, dict) and para_styles
    assert "Normal" in para_styles
    # 每个样式条目至少含 _style_to_dict 的基础字段
    assert STYLE_FIELDS.issubset(para_styles["Normal"].keys())


def test_load_template_logs_warning_for_minimal_master(tmp_path, caplog):
    """母版命名样式过少 → load_template 打印「母版几乎只使用默认样式」告警。"""
    master = _make_minimal_master(tmp_path / "master.docx")

    with caplog.at_level(logging.WARNING, logger="tf.template.engine"):
        load_template(master, RULES_MIN)

    messages = [r.getMessage() for r in caplog.records]
    assert any("母版几乎只使用默认样式" in m for m in messages), messages


def test_load_template_quiet_for_named_style_rich_master(tmp_path, caplog):
    """命名样式充足（仓库自带母版 15 个段落样式）→ 不产生母版告警。"""
    snapshot = extract_master_styles(FIXTURE_MASTER)
    assert len(snapshot["paragraph_styles"]) > FEW_STYLE_THRESHOLD

    with caplog.at_level(logging.WARNING, logger="tf.template.engine"):
        load_template(FIXTURE_MASTER, RULES_MIN)

    assert "母版几乎只使用默认样式" not in caplog.text


def test_load_template_warnings_not_exposed_on_spec(tmp_path):
    """契约缺口（当前实现行为，非期望）：_check_master 的告警只进日志，
    TemplateSpec 没有 warnings 字段（engine.py:37-43 未传、models.py:264-280 无该字段）。
    若将来补上该字段，本测试会失败，用于提醒同步更新契约测试。"""
    master = _make_minimal_master(tmp_path / "master.docx")

    spec = load_template(master, RULES_MIN)

    assert not hasattr(spec, "warnings")
    assert "warnings" not in spec.to_dict()


# --------------------------------------------------------------------------- #
# _check_master
# --------------------------------------------------------------------------- #
def test_check_master_warns_when_few_named_styles():
    """命名样式数 <= 2 时必须产出告警（阈值来自 engine.py:72）。"""
    rules = load_rules(RULES_MIN)
    snapshot = {
        "paragraph_styles": {name: {} for name in ("Normal", "Heading 1")},
        "character_styles": {},
    }

    warnings = _check_master(rules, snapshot)

    assert isinstance(warnings, list) and warnings
    assert any("母版几乎只使用默认样式" in w for w in warnings)
    assert any("命名样式=2" in w for w in warnings)


def test_check_master_silent_with_one_style():
    """只有 Normal 一个样式 → 仍告警，且计数正确。"""
    rules = load_rules(RULES_MIN)

    warnings = _check_master(rules, {"paragraph_styles": {"Normal": {}}})

    assert any("命名样式=1" in w for w in warnings)


def test_check_master_silent_when_styles_rich():
    """命名样式充足 → 无告警（返回空 list，而非 None）。"""
    rules = load_rules(RULES_MIN)
    snapshot = extract_master_styles(FIXTURE_MASTER)
    assert len(snapshot["paragraph_styles"]) > FEW_STYLE_THRESHOLD

    warnings = _check_master(rules, snapshot)

    assert warnings == []


def test_check_master_handles_missing_snapshot_keys():
    """snapshot 缺 paragraph_styles 键时不抛异常（按 0 个样式处理）。"""
    rules = load_rules(RULES_MIN)

    warnings = _check_master(rules, {})

    assert any("命名样式=0" in w for w in warnings)


# --------------------------------------------------------------------------- #
# extract_master_styles
# --------------------------------------------------------------------------- #
def test_extract_master_styles_page_size_and_margins():
    """仓库母版是 A4 纵向：宽/高/页边距应与 A4 相符（范围断言）。"""
    snapshot = extract_master_styles(FIXTURE_MASTER)
    page = snapshot["page"]

    assert set(page) == {
        "width_cm", "height_cm",
        "left_margin_cm", "right_margin_cm",
        "top_margin_cm", "bottom_margin_cm",
    }
    # A4 = 21.0 × 29.7 cm；仓库母版 21.001 × 29.7
    assert 20.5 <= page["width_cm"] <= 21.5
    assert 29.2 <= page["height_cm"] <= 30.2
    # 纵向：宽 < 高
    assert page["width_cm"] < page["height_cm"]
    # 页边距（仓库母版：上下 2.54，左 3.17，右 3.17）
    assert 2.0 <= page["top_margin_cm"] <= 3.0
    assert 2.0 <= page["bottom_margin_cm"] <= 3.0
    assert 2.5 <= page["left_margin_cm"] <= 3.5
    assert 2.5 <= page["right_margin_cm"] <= 3.5
    assert all(isinstance(v, float) for v in page.values())


def test_extract_master_styles_named_style_list():
    """返回的 paragraph_styles 是「样式名 → 样式字典」，含 Normal/Caption 等命名样式。"""
    snapshot = extract_master_styles(FIXTURE_MASTER)
    para_styles = snapshot["paragraph_styles"]

    assert isinstance(para_styles, dict)
    assert "Normal" in para_styles
    assert "Caption" in para_styles
    for name, style_dict in para_styles.items():
        assert isinstance(name, str)
        assert isinstance(style_dict, dict)
        assert STYLE_FIELDS.issubset(style_dict.keys())


def test_extract_master_styles_sections():
    """分节信息：与 doc.sections 数量一致，每节带 page 子字典。"""
    snapshot = extract_master_styles(FIXTURE_MASTER)
    sections = snapshot["sections"]

    assert isinstance(sections, list)
    assert len(sections) == len(docx.Document(str(FIXTURE_MASTER)).sections)
    assert len(sections) >= 1
    for section in sections:
        assert set(section) == {"start_type", "page"}
        assert set(section["page"]) == {
            "width_cm", "height_cm",
            "left_margin_cm", "right_margin_cm",
            "top_margin_cm", "bottom_margin_cm",
        }
        assert section["page"] == snapshot["page"]
    # 已知偏差：master.py:90 把 WD_SECTION_START 的 name（"NEW_PAGE"）喂给
    # _enum_align（只认 left/center/right/justify），因此 start_type 恒为 None。
    # 这里不断言具体值，只断言键存在且类型为 str|None。
    for section in sections:
        assert section["start_type"] is None or isinstance(section["start_type"], str)


def test_extract_master_styles_on_minimal_master(tmp_path):
    """裁剪后的母版只应报出 Normal 与 Heading 1 两个段落样式。"""
    master = _make_minimal_master(tmp_path / "minimal.docx")

    snapshot = extract_master_styles(master)

    assert set(snapshot["paragraph_styles"]) == {"Normal", "Heading 1"}
    assert snapshot["page"]["width_cm"] > 0
    assert snapshot["sections"]  # 至少一个分节
    assert set(snapshot) == {
        "paragraph_styles", "character_styles", "page", "sections", "default_font",
    }


def test_extract_master_styles_default_font_keys():
    """default_font 走同一套 _style_to_dict，至少含基础字段。"""
    snapshot = extract_master_styles(FIXTURE_MASTER)

    assert isinstance(snapshot["default_font"], dict)
    assert STYLE_FIELDS.issubset(snapshot["default_font"].keys())


# --------------------------------------------------------------------------- #
# load_rules：随包分发的 rules_min.json
# --------------------------------------------------------------------------- #
def test_default_rules_path_points_to_repo_rules_min():
    """default_rules_path() 必须指向包内的 parser/rules_min.json。"""
    path = default_rules_path()

    assert path == ROOT / "tf" / "parser" / "rules_min.json"
    assert path.is_file()


def test_load_rules_default_style_from_rules_min():
    """default 段：中文字体宋体、小四（12.0pt）。"""
    rules = load_rules(RULES_MIN)
    default = rules.default

    assert default.font_cn == "宋体"
    assert default.font_en == "Times New Roman"
    assert default.size_pt == pytest.approx(12.0)
    assert default.align is not None and default.align.value == "justify"
    assert default.line_spacing == pytest.approx(1.5)


def test_load_rules_resolve_body_first_line_indent():
    """正文首行缩进 2 字符（字段级级联后应保留 default 的其他字段）。"""
    rules = load_rules(RULES_MIN)

    body = rules.resolve(BlockType.BODY)

    assert body.first_line_indent_chars == pytest.approx(2.0)
    assert body.size_pt == pytest.approx(12.0)
    assert body.font_cn == "宋体"


def test_load_rules_resolve_reference_hanging_indent():
    """参考文献悬挂缩进 2 字符（且仍继承 default 的 font_cn）。"""
    rules = load_rules(RULES_MIN)

    reference = rules.resolve(BlockType.REFERENCE)

    assert reference.hanging_indent_chars == pytest.approx(2.0)
    assert reference.size_pt == pytest.approx(10.5)
    assert reference.font_cn == "宋体"


def test_load_rules_heading_levels():
    """heading.levels == 3（渲染层据此决定标题层级上限）。"""
    rules = load_rules(RULES_MIN)

    assert rules.heading_cfg["levels"] == 3
    assert rules.heading_cfg["numbering"] == "auto"
    assert rules.heading_cfg["caption_prefix"]["figure"] == "图"
    assert rules.heading_cfg["caption_prefix"]["table"] == "表"


def test_load_rules_styles_are_keyed_by_block_type():
    """styles 的键必须是 BlockType 枚举（不是裸字符串）。"""
    rules = load_rules(RULES_MIN)

    assert len(rules.styles) == 15
    assert all(isinstance(k, BlockType) for k in rules.styles)
    for expected in (
        BlockType.COVER_TITLE, BlockType.ABSTRACT_TITLE, BlockType.BODY,
        BlockType.HEADING_1, BlockType.HEADING_2, BlockType.HEADING_3,
        BlockType.CAPTION, BlockType.REFERENCE, BlockType.TOC,
    ):
        assert expected in rules.styles


def test_load_rules_page_cfg_from_rules_min():
    """page 段：A4 + 四边页边距（cm）。"""
    rules = load_rules(RULES_MIN)

    assert rules.page["size"] == "A4"
    assert rules.page["margins_cm"]["top"] == pytest.approx(2.54)
    assert rules.page["margins_cm"]["bottom"] == pytest.approx(2.54)
    assert rules.page["margins_cm"]["left"] == pytest.approx(3.0)
    assert rules.page["margins_cm"]["right"] == pytest.approx(2.5)


def test_load_rules_accepts_str_path():
    """load_rules 同时接受 str 与 Path（CLI 传的是 str）。"""
    from_path = load_rules(RULES_MIN)
    from_str = load_rules(str(RULES_MIN))

    assert from_path.default.font_cn == from_str.default.font_cn
    assert from_str.default.size_pt == pytest.approx(12.0)


# --------------------------------------------------------------------------- #
# rules_from_dict：边界与告警
# --------------------------------------------------------------------------- #
def test_rules_defaults_is_alias_of_default():
    """`defaults` 是 `default` 的等价写法（docs/06 §1.3 契约写法）。"""
    data = {
        "defaults": {"font_cn": "黑体", "size_pt": 15.0},
        "styles": {"body": {"font_cn": "宋体"}},
    }

    rules = rules_from_dict(data)

    assert rules.default.font_cn == "黑体"
    assert rules.default.size_pt == pytest.approx(15.0)
    # specific 未写的字段继续继承 default
    body = rules.resolve(BlockType.BODY)
    assert body.font_cn == "宋体"
    assert body.size_pt == pytest.approx(15.0)


def test_rules_explicit_default_wins_over_defaults():
    """同时给出 default 与 defaults 时以 default 为准。"""
    rules = rules_from_dict(
        {"default": {"font_cn": "宋体"}, "defaults": {"font_cn": "黑体"}}
    )

    assert rules.default.font_cn == "宋体"


def test_rules_unknown_toplevel_key_warns(caplog, tmp_path):
    """未识别的顶层键：告警 + 忽略，不阻断加载。"""
    path = tmp_path / "rules.json"
    path.write_text(
        json.dumps(
            {"default": {"font_cn": "宋体"}, "styles": {}, "made_up_section": {"a": 1}},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    with caplog.at_level(logging.WARNING, logger="tf.template.rules"):
        rules = load_rules(path)

    assert any("未识别的顶层键" in r.getMessage() for r in caplog.records)
    assert any("made_up_section" in r.getMessage() for r in caplog.records)
    # 其余字段照常解析
    assert rules.default.font_cn == "宋体"


def test_rules_unknown_style_type_warns(caplog):
    """styles 里出现未知块类型名：告警 + 忽略该条，其余样式保留。"""
    data = {
        "default": {"font_cn": "宋体"},
        "styles": {
            "body": {"size_pt": 12.0},
            "not_a_real_block": {"size_pt": 99.0},
        },
    }

    with caplog.at_level(logging.WARNING, logger="tf.template.rules"):
        rules = rules_from_dict(data)

    assert any("styles 含未知块类型" in r.getMessage() for r in caplog.records)
    assert any("not_a_real_block" in r.getMessage() for r in caplog.records)
    assert list(rules.styles) == [BlockType.BODY]
    assert rules.styles[BlockType.BODY].size_pt == pytest.approx(12.0)


def test_rules_from_dict_accepts_empty_mapping():
    """空 dict 不抛异常，产出全 None 的 default 与空 styles。"""
    rules = rules_from_dict({})

    assert rules.default.font_cn is None
    assert rules.styles == {}
    assert rules.heading_cfg == {}
    assert rules.page == {}


# --------------------------------------------------------------------------- #
# validate_rules：schema / jsonschema 缺失时的软失败
# --------------------------------------------------------------------------- #
def test_validate_rules_accepts_repo_rules_min():
    """仓库自带的 rules_min.json 必须能通过随包 schema 校验。"""
    data = json.loads(RULES_MIN.read_text(encoding="utf-8"))

    validate_rules(data)  # 不抛异常即通过


def test_validate_rules_silent_when_jsonschema_missing(monkeypatch):
    """缺少 jsonschema 时静默返回（不抛异常、不告警），不安装依赖。

    用 `sys.modules["jsonschema"] = None` 模拟 import 失败
    （CPython 会抛 ImportError: import of jsonschema halted; None in sys.modules）。
    """
    monkeypatch.setitem(sys.modules, "jsonschema", None)
    data = {"default": {"font_cn": "宋体"}, "styles": {}}

    assert validate_rules(data) is None  # 不抛异常


def test_validate_rules_rejects_bad_structure_when_jsonschema_present():
    """jsonschema 可用时结构非法必须报错（证明 schema 真的在校验，而非空跑）。

    环境没有 jsonschema 时跳过 —— 该分支由上面的降级测试覆盖。
    """
    pytest.importorskip("jsonschema")

    with pytest.raises(Exception) as excinfo:
        validate_rules({"default": {"size_pt": "not-a-number"}, "styles": {}})

    assert excinfo.type.__name__ == "ValidationError"