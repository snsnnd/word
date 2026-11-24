"""正则结构解析（见 docs/03 §2、docs/06 §2.3）。

把 Block 列表打上 semantic_tag。第一层：有序规则表，逐条正则匹配。
每个命中规则产出 (BlockType, confidence, need_llm)。

规则表刻意做小、可读、可测；模糊情况留给 LLM 兜底（llm_parser.py）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable, Optional

from tf.core.enums import BlockKind, BlockSource, BlockType
from tf.core.models import Block

# 中文数字（含「两」）
_CN_NUM = "一二三四五六七八九十百零〇两"

# 封面标题关键词
_COVER_TITLE_KW = re.compile(r"(实验|毕业|学位|课程|设计|实训|实习|课程设计|毕业论文|学位论文).*(报告|论文|说明书|设计)")


@dataclass
class Rule:
    pattern: str
    block_type: BlockType
    confidence: float
    guard: Optional[Callable[["Context", Block], bool]] = None
    regex: Optional[re.Pattern] = field(default=None, repr=False)

    def __post_init__(self):
        self.regex = re.compile(self.pattern)


@dataclass
class Context:
    """解析上下文，随遍历更新。"""
    section_index: int = 0
    prev_heading_level: int = 0
    in_abstract: bool = False
    in_references: bool = False


def _is_short(text: str, n: int) -> bool:
    return len(text.strip()) <= n


def _is_centered(block: Block) -> bool:
    """best-effort 判断段落是否居中（用于封面标题判断）。"""
    p = block.parent_paragraph
    if p is not None and p.alignment is not None:
        name = getattr(p.alignment, "name", "").lower()
        if "center" in name:
            return True
    # 退化：检查 run 里是否有居中标记（罕见）
    return False


def _is_bold(block: Block) -> bool:
    return any(r.bold for r in block.runs if r.bold is not None)


# 规则表：顺序敏感，第一条命中生效。
# guard(ctx, block) -> bool，None 表示无额外条件。
RULES: list[Rule] = [
    Rule(r"^摘\s*要$", BlockType.ABSTRACT_TITLE, 0.99),
    Rule(r"^(关\s*键\s*词|关键字)\s*[:：]", BlockType.KEYWORDS, 0.98),
    Rule(r"^目\s*录$", BlockType.TOC, 0.99),
    Rule(r"^参\s*考\s*文\s*献$", BlockType.REFERENCE_TITLE, 0.99),
    Rule(r"^致\s*谢$", BlockType.APPENDIX_TITLE, 0.90),
    Rule(r"^附\s*录", BlockType.APPENDIX_TITLE, 0.90),
    # 章标题：第X章
    Rule(r"^第\s*[" + _CN_NUM + r"]+\s*章", BlockType.HEADING_1, 0.98),
    # 三级：1.2.3
    Rule(r"^\d+\.\d+\.\d+", BlockType.HEADING_3, 0.96),
    # 二级：1.2（后跟空格/非数字非点）
    Rule(r"^\d+\.\d+(?:[\s、]|(?![\d.]))", BlockType.HEADING_2, 0.95),
    # 图/表标题
    Rule(r"^(图|表)\s?\d+", BlockType.CAPTION, 0.97),
    # 参考文献条目
    Rule(r"^\[\d+\]", BlockType.REFERENCE, 0.98, guard=lambda c, b: c.in_references),
    # 封面区（section 0 内，短、居中、标题关键词；允许字间空格）
    Rule(
        r".*",
        BlockType.COVER_TITLE,
        0.90,
        guard=lambda c, b: (
            c.section_index == 0
            and _is_short(b.text, 25)
            and _is_centered(b)
            and _COVER_TITLE_KW.search(re.sub(r"\s+", "", b.text)) is not None
        ),
    ),
    # 封面信息行（课程/学院/专业/班级/学号/姓名/指导教师...）
    Rule(
        r"^(课程|学院|专业|班级|学号|姓名|指导教师|实验地点|实验日期|实验名称|成绩|评语)\s*[:：]",
        BlockType.COVER_META,
        0.92,
    ),
]

# 编译一次
for _r in RULES:
    _r.regex = re.compile(_r.pattern)


def classify(blocks: list[Block]) -> list[Block]:
    """就地给每个 Block 打 semantic_tag / confidence / source。

    受保护块（图片/表格/公式）直接给对应 tag，不走正则。
    """
    ctx = Context()
    for block in blocks:
        _classify_block(block, ctx)
    return blocks


def _classify_block(block: Block, ctx: Context) -> None:
    # 非段落块：按 kind 直接定 tag
    if block.kind == BlockKind.IMAGE:
        block.semantic_tag = BlockType.IMAGE
        block.confidence = 1.0
        block.source = BlockSource.PRESET
        return
    if block.kind == BlockKind.TABLE:
        block.semantic_tag = BlockType.TABLE
        block.confidence = 1.0
        block.source = BlockSource.PRESET
        return
    if block.kind == BlockKind.FORMULA:
        block.semantic_tag = BlockType.FORMULA
        block.confidence = 1.0
        block.source = BlockSource.PRESET
        return
    if block.kind != BlockKind.PARAGRAPH:
        return

    text = block.text.strip()
    # 空段落：先给一个占位 tag，后续由空行清理处理
    if not text:
        block.semantic_tag = BlockType.UNKNOWN
        block.confidence = 0.0
        block.source = None
        return

    # 有序规则匹配
    for rule in RULES:
        if rule.regex.search(text) is None:
            continue
        if rule.guard is not None and not rule.guard(ctx, block):
            continue
        block.semantic_tag = rule.block_type
        block.confidence = rule.confidence
        block.source = BlockSource.REGEX
        _update_context(ctx, block)
        return

    # 启发式：粗体 + 短 + 不以句号结尾 → 疑似标题（低置信，交 LLM 兜底）
    if _is_bold(block) and _is_short(text, 30) and not text.endswith(("。", "；", "，", ".")):
        block.semantic_tag = BlockType.HEADING_1
        block.confidence = 0.55
        block.source = BlockSource.REGEX
        _update_context(ctx, block)
        return

    # 默认正文
    block.semantic_tag = BlockType.BODY
    block.confidence = 0.60
    block.source = BlockSource.REGEX
    _update_context(ctx, block)


def _update_context(ctx: Context, block: Block) -> None:
    tag = block.semantic_tag
    if tag == BlockType.HEADING_1:
        ctx.prev_heading_level = 1
        ctx.in_abstract = False
        ctx.in_references = False
    elif tag in (BlockType.HEADING_2, BlockType.HEADING_3, BlockType.HEADING_4):
        ctx.prev_heading_level = {BlockType.HEADING_2: 2, BlockType.HEADING_3: 3, BlockType.HEADING_4: 4}[tag]
    elif tag == BlockType.ABSTRACT_TITLE:
        ctx.in_abstract = True
    elif tag == BlockType.REFERENCE_TITLE:
        ctx.in_references = True