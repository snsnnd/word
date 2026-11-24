"""模板引擎入口（见 docs/01 §4、docs/06 §2.1）。

load_template(master_path, rules_path) -> TemplateSpec

流程：
1. 读 rules.json → RuleSet
2. 抽取母版样式快照 → style_snapshot
3. 合成 TemplateSpec
4. 记录告警（母版缺样式/规则引用未知类型等）
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Optional

from tf.core.models import TemplateSpec
from tf.template.master import extract_master_styles
from tf.template.rules import load_rules, rules_from_dict

log = logging.getLogger(__name__)


def load_template(
    master_path: str | Path,
    rules_path: str | Path,
    name: str = "custom",
    version: str = "0.1.0",
) -> TemplateSpec:
    """加载模板规范。"""
    master_path = str(master_path)
    rules = load_rules(rules_path)

    snapshot = extract_master_styles(master_path)
    warnings = _check_master(rules, snapshot)

    spec = TemplateSpec(
        name=name,
        version=version,
        master_path=master_path,
        rules=rules,
        style_snapshot=snapshot,
    )
    for w in warnings:
        log.warning("[template] %s", w)
    return spec


def load_template_from_rules(
    master_path: str | Path,
    rules_dict: dict[str, Any],
    name: str = "custom",
    version: str = "0.1.0",
) -> TemplateSpec:
    """从 dict 直接构造（供编程式使用/测试）。"""
    rules = rules_from_dict(rules_dict)
    snapshot = extract_master_styles(master_path)
    return TemplateSpec(
        name=name,
        version=version,
        master_path=str(master_path),
        rules=rules,
        style_snapshot=snapshot,
    )


def _check_master(rules, snapshot: dict[str, Any]) -> list[str]:
    """检查母版与规则表的匹配度，产出告警。"""
    warnings: list[str] = []
    para_styles = snapshot.get("paragraph_styles", {})
    # 命名样式太少 → 渲染层需依赖直接格式化
    if len(para_styles) <= 2:
        warnings.append(
            "母版几乎只使用默认样式（命名样式=%d），渲染将主要依赖直接格式化。" % len(para_styles)
        )
    # 规则引用的样式名若母版不存在，提示（不阻断，渲染层会新建样式）
    return warnings