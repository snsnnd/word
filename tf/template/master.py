"""母版样式抽取（见 docs/01 §1、docs/01 §4）。

从母版 docx 中提取：
- 段落样式（style 名 → 字体/字号/对齐/行距/缩进）
- 字符样式（run 级）
- 页面设置（页边距/纸张）
- 分节信息

抽到的样式作为 style_snapshot 存进 TemplateSpec，供渲染层做「未配置项的默认回落」。

注意：真实模板可能几乎全是直接格式化、命名样式只有 Normal（本仓库的
tests/fixtures/source_template.docx 就是这种情况）。此时 style_snapshot
主要作为兜底，渲染层会以 rules.json 的显式规则为主。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Optional

import docx
from docx.shared import Pt

from tf.core.enums import Align


def _enum_align(alignment: Any) -> Optional[str]:
    """python-docx 的 WD_ALIGN_PARAGRAPH → 我们的对齐字符串。"""
    if alignment is None:
        return None
    name = getattr(alignment, "name", None)
    if name is None:
        name = str(alignment)
    name = name.lower()
    mapping = {
        "left": "left",
        "center": "center",
        "right": "right",
        "justify": "justify",
        "distribute": "justify",
    }
    return mapping.get(name)


def extract_master_styles(docx_path: str | Path) -> dict[str, Any]:
    """抽取母版样式快照。

    返回 dict：
      {
        "paragraph_styles": {style_name: {font_cn, font_en, size_pt, bold,
                            align, line_spacing, first_line_indent_chars, ...}},
        "character_styles": {...},
        "page": {width_cm, height_cm, margins_cm},
        "sections": [{start_type, ...}],
        "default_font": {...},
      }
    """
    path = Path(docx_path)
    d = docx.Document(str(path))
    snapshot: dict[str, Any] = {
        "paragraph_styles": {},
        "character_styles": {},
        "page": {},
        "sections": [],
        "default_font": {},
    }

    # 段落样式
    for style in d.styles:
        try:
            if style.type is not None and style.type == 1:  # WD_STYLE_TYPE.PARAGRAPH
                snapshot["paragraph_styles"][style.name] = _style_to_dict(style)
            elif style.type is not None and style.type == 2:  # CHARACTER
                snapshot["character_styles"][style.name] = _style_to_dict(style)
        except Exception:
            continue

    # 页面 + 分节
    for s in d.sections:
        page = {
            "width_cm": _cm(s.page_width),
            "height_cm": _cm(s.page_height),
            "left_margin_cm": _cm(s.left_margin),
            "right_margin_cm": _cm(s.right_margin),
            "top_margin_cm": _cm(s.top_margin),
            "bottom_margin_cm": _cm(s.bottom_margin),
        }
        if not snapshot["page"]:
            snapshot["page"] = page
        snapshot["sections"].append({
            "start_type": _enum_align(getattr(s.start_type, "name", str(s.start_type))),
            "page": page,
        })

    # 默认字体（Normal 样式）
    try:
        normal = d.styles["Normal"]
        snapshot["default_font"] = _style_to_dict(normal)
    except Exception:
        pass

    return snapshot


def _style_to_dict(style: Any) -> dict[str, Any]:
    """把一个 python-docx 样式转成 dict（best-effort，缺失字段为 None）。"""
    out: dict[str, Any] = {
        "font_cn": None,
        "font_en": None,
        "size_pt": None,
        "bold": None,
        "italic": None,
        "align": None,
        "line_spacing": None,
        "first_line_indent_chars": None,
        "hanging_indent_chars": None,
        "space_before_pt": None,
        "space_after_pt": None,
    }
    try:
        font = style.font
        if font is not None:
            out["font_en"] = font.name
            if font.size is not None:
                out["size_pt"] = font.size.pt
            out["bold"] = font.bold
            out["italic"] = font.italic
            # 中文字体需要读 rPr/rFonts w:eastAsia
            cn = _east_asia_font(style)
            if cn:
                out["font_cn"] = cn
    except Exception:
        pass
    try:
        pf = style.paragraph_format
        if pf is not None:
            out["align"] = _enum_align(pf.alignment)
            if pf.line_spacing is not None:
                out["line_spacing"] = pf.line_spacing
            if pf.space_before is not None:
                out["space_before_pt"] = pf.space_before.pt
            if pf.space_after is not None:
                out["space_after_pt"] = pf.space_after.pt
            # 首行缩进（字符数需要字号换算，这里保留原始 pt）
            if pf.first_line_indent is not None:
                out["first_line_indent_pt"] = pf.first_line_indent.pt
    except Exception:
        pass
    return out


def _east_asia_font(style: Any) -> Optional[str]:
    """从样式的 rPr/rFonts 读 w:eastAsia（中文字体）。"""
    try:
        rpr = style.element.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
        )
        if rpr is None:
            return None
        rfonts = rpr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts"
        )
        if rfonts is None:
            return None
        return rfonts.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia"
        )
    except Exception:
        return None


def _cm(emu: Optional[int]) -> Optional[float]:
    """EMU → cm。"""
    if emu is None:
        return None
    try:
        return round(int(emu) / 360000.0, 3)
    except Exception:
        return None