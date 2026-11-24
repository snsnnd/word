"""核心数据模型（见 docs/06 §1）。

所有跨层传递的数据都在这里定义。使用 dataclass + to_dict/from_dict 以便序列化
（校对台存盘、契约测试、跨进程传递）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from tf.core.enums import (
    Align,
    BlockKind,
    BlockSource,
    BlockType,
    VertAlign,
)


@dataclass
class RunInfo:
    """run 级信息，供字体分离/上下标判断使用。"""
    text: str
    font_cn: Optional[str] = None
    font_en: Optional[str] = None
    size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    vert_align: Optional[VertAlign] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "font_cn": self.font_cn,
            "font_en": self.font_en,
            "size_pt": self.size_pt,
            "bold": self.bold,
            "italic": self.italic,
            "vert_align": self.vert_align.value if self.vert_align else None,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "RunInfo":
        va = d.get("vert_align")
        return cls(
            text=d.get("text", ""),
            font_cn=d.get("font_cn"),
            font_en=d.get("font_en"),
            size_pt=d.get("size_pt"),
            bold=d.get("bold"),
            italic=d.get("italic"),
            vert_align=VertAlign(va) if va else None,
        )


@dataclass
class AssetInfo:
    """资产信息（图片尺寸/表格结构/公式/引用）。"""
    kind: BlockKind
    width_cm: Optional[float] = None
    height_cm: Optional[float] = None
    rows: Optional[int] = None
    cols: Optional[int] = None
    has_header: Optional[bool] = None
    omath: bool = False
    ole: bool = False
    rel_id: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind.value,
            "width_cm": self.width_cm,
            "height_cm": self.height_cm,
            "rows": self.rows,
            "cols": self.cols,
            "has_header": self.has_header,
            "omath": self.omath,
            "ole": self.ole,
            "rel_id": self.rel_id,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "AssetInfo":
        return cls(
            kind=BlockKind(d["kind"]),
            width_cm=d.get("width_cm"),
            height_cm=d.get("height_cm"),
            rows=d.get("rows"),
            cols=d.get("cols"),
            has_header=d.get("has_header"),
            omath=d.get("omath", False),
            ole=d.get("ole", False),
            rel_id=d.get("rel_id"),
        )


@dataclass
class Block:
    """文档块（见 docs/06 §1.2）。

    xml_element / parent_paragraph 为运行时对象，不序列化（见 to_dict）。
    """
    block_id: str
    index: int
    kind: BlockKind
    protected: bool = False
    asset: Optional[AssetInfo] = None
    text: str = ""
    runs: list[RunInfo] = field(default_factory=list)
    xml_element: Any = None          # 运行时：lxml element，不序列化
    parent_paragraph: Any = None     # 运行时：python-docx Paragraph
    section_index: Optional[int] = None
    semantic_tag: Optional[BlockType] = None
    confidence: float = 0.0
    source: Optional[BlockSource] = None

    def needs_llm(self) -> bool:
        """正则规则表触发、但置信度低的块，进入 LLM 兜底。"""
        return (
            self.source == BlockSource.REGEX
            and self.confidence < 0.70
            and self.kind == BlockKind.PARAGRAPH
            and self.semantic_tag != BlockType.UNKNOWN
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "index": self.index,
            "kind": self.kind.value,
            "protected": self.protected,
            "asset": self.asset.to_dict() if self.asset else None,
            "text": self.text,
            "runs": [r.to_dict() for r in self.runs],
            "section_index": self.section_index,
            "semantic_tag": self.semantic_tag.value if self.semantic_tag else None,
            "confidence": self.confidence,
            "source": self.source.value if self.source else None,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Block":
        return cls(
            block_id=d["block_id"],
            index=d["index"],
            kind=BlockKind(d["kind"]),
            protected=d.get("protected", False),
            asset=AssetInfo.from_dict(d["asset"]) if d.get("asset") else None,
            text=d.get("text", ""),
            runs=[RunInfo.from_dict(r) for r in d.get("runs", [])],
            xml_element=None,
            parent_paragraph=None,
            section_index=d.get("section_index"),
            semantic_tag=BlockType(d["semantic_tag"]) if d.get("semantic_tag") else None,
            confidence=d.get("confidence", 0.0),
            source=BlockSource(d["source"]) if d.get("source") else None,
        )


@dataclass
class Style:
    """样式规则（见 docs/06 §1.3）。None 表示继承。"""
    font_cn: Optional[str] = None
    font_en: Optional[str] = None
    size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    align: Optional[Align] = None
    line_spacing: Optional[float] = None
    first_line_indent_chars: Optional[float] = None
    hanging_indent_chars: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "font_cn": self.font_cn,
            "font_en": self.font_en,
            "size_pt": self.size_pt,
            "bold": self.bold,
            "italic": self.italic,
            "align": self.align.value if self.align else None,
            "line_spacing": self.line_spacing,
            "first_line_indent_chars": self.first_line_indent_chars,
            "hanging_indent_chars": self.hanging_indent_chars,
            "space_before_pt": self.space_before_pt,
            "space_after_pt": self.space_after_pt,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Style":
        a = d.get("align")
        return cls(
            font_cn=d.get("font_cn"),
            font_en=d.get("font_en"),
            size_pt=d.get("size_pt"),
            bold=d.get("bold"),
            italic=d.get("italic"),
            align=Align(a) if a else None,
            line_spacing=d.get("line_spacing"),
            first_line_indent_chars=d.get("first_line_indent_chars"),
            hanging_indent_chars=d.get("hanging_indent_chars"),
            space_before_pt=d.get("space_before_pt"),
            space_after_pt=d.get("space_after_pt"),
        )


@dataclass
class RuleSet:
    """规则表（见 docs/06 §1.3 / docs/01）。"""
    default: Style = field(default_factory=Style)
    styles: dict[BlockType, Style] = field(default_factory=dict)
    heading_cfg: dict[str, Any] = field(default_factory=dict)
    image_cfg: dict[str, Any] = field(default_factory=dict)
    table_cfg: dict[str, Any] = field(default_factory=dict)
    formula_cfg: dict[str, Any] = field(default_factory=dict)
    section_cfg: dict[str, Any] = field(default_factory=dict)
    page: dict[str, Any] = field(default_factory=dict)
    classify: dict[str, Any] = field(default_factory=dict)

    def resolve(self, block_type: BlockType) -> Style:
        """字段级级联：specific 的非空字段覆盖 default，未指定字段继承 default。

        见 docs/06 §2.2。例如 body 只设 size_pt=10 时，font_cn/line_spacing
        仍继承 default，而不是被置空。
        """
        base = self.default.to_dict()
        specific = self.styles.get(block_type)
        if specific is None:
            return self.default
        for key, value in specific.to_dict().items():
            if value is not None:
                base[key] = value
        return Style.from_dict(base)

    def to_dict(self) -> dict[str, Any]:
        return {
            "default": self.default.to_dict(),
            "styles": {k.value: v.to_dict() for k, v in self.styles.items()},
            "heading": self.heading_cfg,
            "image": self.image_cfg,
            "table": self.table_cfg,
            "formula": self.formula_cfg,
            "section": self.section_cfg,
            "page": self.page,
            "classify": self.classify,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "RuleSet":
        return cls(
            default=Style.from_dict(d.get("default", {})),
            styles={BlockType(k): Style.from_dict(v) for k, v in d.get("styles", {}).items()},
            heading_cfg=d.get("heading", {}),
            image_cfg=d.get("image", {}),
            table_cfg=d.get("table", {}),
            formula_cfg=d.get("formula", {}),
            section_cfg=d.get("section", {}),
            page=d.get("page", {}),
            classify=d.get("classify", {}),
        )


@dataclass
class TemplateSpec:
    """模板规范：母版路径 + 规则表 + 样式快照（见 docs/06 §1.3）。"""
    name: str
    version: str
    master_path: str
    rules: RuleSet
    style_snapshot: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "version": self.version,
            "master_path": self.master_path,
            "rules": self.rules.to_dict(),
            "style_snapshot": self.style_snapshot,
        }


@dataclass
class RenderResult:
    """渲染产物（见 docs/06 §1.5）。"""
    docx_path: str
    pdf_path: Optional[str] = None
    qa_report: dict[str, Any] = field(default_factory=dict)
    engine: str = "none"


@dataclass
class FinalizeResult:
    """收尾结果（见 docs/06 §1.5）。"""
    engine: str
    docx_path: Optional[str] = None
    pdf_path: Optional[str] = None
    toc_updated: bool = False
    fields_updated: bool = False
    degraded: bool = False
    warnings: list[str] = field(default_factory=list)


@dataclass
class Fix:
    """单条人工修正（见 docs/06 §1.6）。"""
    block_id: str
    type: BlockType
    source: BlockSource = BlockSource.HUMAN
    note: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "type": self.type.value,
            "source": self.source.value,
            "note": self.note,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Fix":
        return cls(
            block_id=d["block_id"],
            type=BlockType(d["type"]),
            source=BlockSource(d.get("source", "human")),
            note=d.get("note"),
        )