"""渲染层 OOXML 底层操作（字体 / 段落格式 / 缩进 / 大纲级别）。

只做「把 Style 字段落到 python-docx 对象」的最小集合，不含业务判断。
所有函数原地修改传入的 python-docx 对象。
"""
from __future__ import annotations

from typing import Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn as _qn
from docx.shared import Pt

from tf.core.enums import Align

# 字符数缩进：OOXML w:firstLineChars / w:hangingChars 以「1/100 字符」为单位
_CHARS_UNIT = 100.0
# 字符数 → twips 兜底（1 pt = 20 twips）
_TWIPPS_PER_PT = 20.0

_ALIGN_MAP = {
    Align.LEFT: WD_ALIGN_PARAGRAPH.LEFT,
    Align.CENTER: WD_ALIGN_PARAGRAPH.CENTER,
    Align.RIGHT: WD_ALIGN_PARAGRAPH.RIGHT,
    Align.JUSTIFY: WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def set_run_font(
    run,
    *,
    font_cn: Optional[str] = None,
    font_en: Optional[str] = None,
    size_pt: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
) -> None:
    """设置单个 run 的字体（中文 eastAsia / 西文 ascii+hAnsi）、字号、加粗、倾斜。

    不触碰 vertAlign（上标/下标）等其它属性。
    """
    if font_en:
        run.font.name = font_en  # 同时写 w:ascii 与 w:hAnsi
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if font_cn:
        rpr = run._r.get_or_add_rPr()
        rfonts = rpr.find(_qn("w:rFonts"))
        if rfonts is None:
            from docx.oxml import OxmlElement
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(_qn("w:eastAsia"), font_cn)


def set_para_format(
    paragraph,
    *,
    align: Optional[Align] = None,
    line_spacing: Optional[float] = None,
    first_line_indent_chars: Optional[float] = None,
    hanging_indent_chars: Optional[float] = None,
    space_before_pt: Optional[float] = None,
    space_after_pt: Optional[float] = None,
    font_size_pt: float = 12.0,
) -> None:
    """设置段落格式：对齐 / 行距（倍数）/ 首行缩进 / 悬挂缩进 / 段前后距。

    缩进用「字符数」语义（中文排版惯例），落到 w:firstLineChars / w:hangingChars，
    并给出 twips 兜底，兼容不支持 Chars 的阅读器。
    """
    pf = paragraph.paragraph_format
    if align is not None:
        pf.alignment = _ALIGN_MAP[align]
    if line_spacing is not None:
        pf.line_spacing = float(line_spacing)
    if space_before_pt is not None:
        pf.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        pf.space_after = Pt(space_after_pt)

    if first_line_indent_chars is None and hanging_indent_chars is None:
        return

    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    if first_line_indent_chars is not None:
        ind.set(_qn("w:firstLineChars"), str(int(round(first_line_indent_chars * _CHARS_UNIT))))
        ind.set(_qn("w:firstLine"), str(int(round(first_line_indent_chars * font_size_pt * _TWIPPS_PER_PT))))
    if hanging_indent_chars is not None:
        ind.set(_qn("w:hangingChars"), str(int(round(hanging_indent_chars * _CHARS_UNIT))))
        ind.set(_qn("w:hanging"), str(int(round(hanging_indent_chars * font_size_pt * _TWIPPS_PER_PT))))


def set_outline_level(paragraph, level: int) -> None:
    """设置段落大纲级别（w:outlineLvl，0 起）。供导航窗格 / 目录 \\o 识别标题。"""
    pPr = paragraph._p.get_or_add_pPr()
    ol = pPr.find(_qn("w:outlineLvl"))
    if ol is None:
        from docx.oxml import OxmlElement
        ol = OxmlElement("w:outlineLvl")
        pPr.append(ol)
    ol.set(_qn("w:val"), str(level - 1))