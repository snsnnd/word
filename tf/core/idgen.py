"""block_id 生成（见 docs/06 §1.4）。

block_id = "b" + zero_pad(index, 4) + "_" + short_hash(块文本前64字 + 块XML结构摘要)

设计目标：
- 稳定：同一文档重复解析得到同一 ID，使 review_fixes.json 可重放。
- 可读：前缀是序号，肉眼能定位。
- 抗碰撞：加结构摘要哈希，避免文本相同但上下文不同的块撞 ID。
"""
from __future__ import annotations

import hashlib
from typing import Any


def _structure_digest(xml_element: Any) -> str:
    """取块 XML 的结构摘要（不含具体文本内容，只含标签骨架），用于抗碰撞。"""
    if xml_element is None:
        return ""
    try:
        from lxml import etree
        # 只序列化标签名，去掉文本/属性值，避免内容变化导致 ID 变化
        tags = [el.tag for el in xml_element.iter() if isinstance(el.tag, str)]
        skeleton = "|".join(tags)
        return hashlib.sha1(skeleton.encode("utf-8")).hexdigest()[:8]
    except Exception:
        return ""


def make_block_id(index: int, text: str, xml_element: Any = None) -> str:
    """生成稳定 block_id。"""
    prefix = f"b{index:04d}_"
    h = hashlib.sha1()
    h.update((text or "")[:64].encode("utf-8"))
    struct = _structure_digest(xml_element)
    if struct:
        h.update(struct.encode("utf-8"))
    return prefix + h.hexdigest()[:8]


def parse_index(block_id: str) -> int:
    """从 block_id 反解序号，用于排序。"""
    try:
        return int(block_id[1:5])
    except (ValueError, IndexError):
        return 1 << 30