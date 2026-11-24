"""核心枚举定义（见 docs/06 §1.1）。"""
from __future__ import annotations

from enum import Enum


class BlockKind(str, Enum):
    """块的大类，决定「怎么渲染」。"""
    PARAGRAPH = "paragraph"
    TABLE = "table"
    IMAGE = "image"
    FORMULA = "formula"
    FOOTNOTE_REF = "footnote_ref"
    SECTION_BREAK = "section_break"
    TOC = "toc"


class BlockType(str, Enum):
    """语义标签（细分用途），决定「套什么样式」。"""
    COVER_TITLE = "cover_title"
    COVER_META = "cover_meta"
    ABSTRACT_TITLE = "abstract_title"
    ABSTRACT_BODY = "abstract_body"
    KEYWORDS = "keywords"
    TOC = "toc"
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    HEADING_4 = "heading_4"
    BODY = "body"
    CAPTION = "caption"
    REFERENCE = "reference"
    REFERENCE_TITLE = "reference_title"
    APPENDIX_TITLE = "appendix_title"
    IMAGE = "image"
    TABLE = "table"
    FORMULA = "formula"
    FOOTNOTE = "footnote"
    UNKNOWN = "unknown"


class BlockSource(str, Enum):
    """标签来源，用于校对台展示「系统认为」vs「你修正」。"""
    REGEX = "regex"
    LLM = "llm"
    HUMAN = "human"
    PRESET = "preset"


class Align(str, Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


class VertAlign(str, Enum):
    BASELINE = "baseline"
    SUPERSCRIPT = "superscript"
    SUBSCRIPT = "subscript"