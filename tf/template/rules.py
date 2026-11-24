"""规则表加载与校验（见 docs/01 §2、docs/06 §2）。

rules.json 的结构见 docs/01。这里负责：
- 读 JSON
- 用 jsonschema 校验（schema 文件在 schemas/rules.schema.json）
- 转成 RuleSet 数据模型
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Optional

from tf.core.enums import Align, BlockType
from tf.core.models import RuleSet, Style

log = logging.getLogger(__name__)

_DEFAULT_RULES_NAME = "rules_min.json"

#: 规则表顶层允许出现的键（docs/06 §1.3；default/defaults 为同一字段两种写法）
_KNOWN_KEYS = {
    "default", "defaults", "styles", "heading", "image",
    "table", "formula", "section", "page", "classify", "finalize",
}


def _align(value: Optional[str]) -> Optional[Align]:
    if value is None:
        return None
    return Align(value)


def load_rules(path: str | Path) -> RuleSet:
    """从 JSON 文件加载规则表。"""
    p = Path(path)
    with p.open("r", encoding="utf-8") as f:
        data = json.load(f)
    validate_rules(data)
    return rules_from_dict(data)


def rules_from_dict(data: dict[str, Any]) -> RuleSet:
    """从 dict 构造 RuleSet（供测试/编程式配置使用）。

    容错：顶层 `defaults`（docs/06 §1.3 契约写法）与 `default`（内置规则文件写法）
    等价；未识别的顶层键与未知块类型只告警不阻断。
    """
    default_raw = data.get("default")
    if default_raw is None:
        default_raw = data.get("defaults", {}) or {}
    default = _style_from_dict(default_raw or {})
    styles: dict[BlockType, Style] = {}
    for key, val in (data.get("styles") or {}).items():
        try:
            bt = BlockType(key)
        except ValueError:
            # 未知类型名忽略（前向兼容）
            log.warning("[rules] styles 含未知块类型 %r，已忽略", key)
            continue
        styles[bt] = _style_from_dict(val)
    for key in data:
        if key not in _KNOWN_KEYS:
            log.warning("[rules] 未识别的顶层键 %r，已忽略", key)
    return RuleSet(
        default=default,
        styles=styles,
        heading_cfg=data.get("heading", {}) or {},
        image_cfg=data.get("image", {}) or {},
        table_cfg=data.get("table", {}) or {},
        formula_cfg=data.get("formula", {}) or {},
        section_cfg=data.get("section", {}) or {},
        page=data.get("page", {}) or {},
        classify=data.get("classify", {}) or {},
    )


def _style_from_dict(d: dict[str, Any]) -> Style:
    return Style(
        font_cn=d.get("font_cn"),
        font_en=d.get("font_en"),
        size_pt=d.get("size_pt"),
        bold=d.get("bold"),
        italic=d.get("italic"),
        align=_align(d.get("align")),
        line_spacing=d.get("line_spacing"),
        first_line_indent_chars=d.get("first_line_indent_chars"),
        hanging_indent_chars=d.get("hanging_indent_chars"),
        space_before_pt=d.get("space_before_pt"),
        space_after_pt=d.get("space_after_pt"),
    )


def validate_rules(data: dict[str, Any]) -> None:
    """用 jsonschema 校验规则表结构。schema 缺失时跳过（软失败）。"""
    from tf import __file__ as pkg_file
    schema_path = Path(pkg_file).parent.parent / "schemas" / "rules.schema.json"
    if not schema_path.exists():
        return
    try:
        import jsonschema
    except ImportError:
        return
    with schema_path.open("r", encoding="utf-8") as f:
        schema = json.load(f)
    jsonschema.validate(instance=data, schema=schema)


def default_rules_path() -> Path:
    """返回随包分发的默认规则表路径。"""
    from tf import __file__ as pkg_file
    return Path(pkg_file).parent / "parser" / _DEFAULT_RULES_NAME