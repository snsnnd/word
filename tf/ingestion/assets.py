"""资产探测（见 docs/02 §5）。

把 OOXML 元素探测成 kind + AssetInfo。所有函数只读，不修改元素。

命名空间常量集中在这里，避免各处硬编码。
"""
from __future__ import annotations

from typing import Optional

from lxml import etree

from tf.core.enums import BlockKind
from tf.core.models import AssetInfo

# OOXML 命名空间
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
}

W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]
WP = "{%s}" % NS["wp"]
A = "{%s}" % NS["a"]
R = "{%s}" % NS["r"]
O = "{%s}" % NS["o"]

EMU_PER_CM = 360000.0


def qn(tag: str) -> str:
    """'w:p' → '{namespace}p'。"""
    prefix, local = tag.split(":", 1)
    return "{%s}%s" % (NS[prefix], local)


def detect_paragraph_kind(p_element) -> tuple[BlockKind, Optional[AssetInfo]]:
    """探测一个 w:p 元素的 kind 与资产信息。

    返回 (kind, asset_info)。判定优先级（见 docs/02 §5）：
    - 有非空文字 → PARAGRAPH（图/公式 run 在渲染时只改 rPr，原样保留）
    - 无文字 + 图片 → IMAGE
    - 无文字 + OMML/OLE/pict → FORMULA
    - 否则 → PARAGRAPH（空段）

    asset_info 仅在需要时返回（图片宽高、OMML/OLE 标记）。
    """
    if p_element is None:
        return BlockKind.PARAGRAPH, None

    text = "".join(t.text or "" for t in p_element.iter(qn("w:t"))).strip()
    img = _detect_image(p_element)

    # 有文字：一律段落（图片作为 asset 记录，渲染时不破坏）
    if text:
        return BlockKind.PARAGRAPH, img

    # 纯图片
    if img is not None:
        return BlockKind.IMAGE, img

    # 纯公式
    if _has_omath(p_element):
        return BlockKind.FORMULA, AssetInfo(kind=BlockKind.FORMULA, omath=True)
    if _has_ole(p_element) or _has_pict(p_element):
        return BlockKind.FORMULA, AssetInfo(kind=BlockKind.FORMULA, ole=_has_ole(p_element))

    return BlockKind.PARAGRAPH, None


def _detect_image(p_element) -> Optional[AssetInfo]:
    """提取图片宽高（cm）与 rel_id。"""
    # DrawingML 图片
    for tag in ("wp:inline", "wp:anchor"):
        for drawing in p_element.iterfind(".//" + qn(tag)):
            extent = drawing.find(qn("wp:extent"))
            blip = drawing.find(".//" + qn("a:blip"))
            w_cm = h_cm = None
            rel_id = None
            if extent is not None:
                cx = extent.get("cx")
                cy = extent.get("cy")
                if cx is not None:
                    w_cm = round(int(cx) / EMU_PER_CM, 3)
                if cy is not None:
                    h_cm = round(int(cy) / EMU_PER_CM, 3)
            if blip is not None:
                rel_id = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
            return AssetInfo(
                kind=BlockKind.IMAGE,
                width_cm=w_cm,
                height_cm=h_cm,
                rel_id=rel_id,
            )
    # VML 图片（旧式 w:pict/v:imagedata）
    for imagedata in p_element.iterfind(".//" + qn("v:imagedata")):
        rel_id = imagedata.get(qn("r:id"))
        return AssetInfo(kind=BlockKind.IMAGE, rel_id=rel_id)
    return None


def _has_omath(p_element) -> bool:
    for tag in ("m:oMathPara", "m:oMath"):
        if p_element.find(".//" + qn(tag)) is not None:
            return True
    return False


def _has_ole(p_element) -> bool:
    if p_element.find(".//" + qn("w:object")) is not None:
        return True
    # o:OLEObject
    for el in p_element.iter():
        if isinstance(el.tag, str) and el.tag.endswith("}OLEObject"):
            return True
    return False


def _has_pict(p_element) -> bool:
    return p_element.find(".//" + qn("w:pict")) is not None


def table_dimensions(tbl_element) -> tuple[int, int, bool]:
    """返回表格 (rows, cols, has_header)。

    has_header 的启发式：存在 w:tblHeader 标记，或第一行加粗。
    """
    rows = tbl_element.findall(qn("w:tr"))
    n_rows = len(rows)
    n_cols = 0
    has_header = False
    if rows:
        first_cells = rows[0].findall(qn("w:tc"))
        n_cols = len(first_cells)
        # tblHeader 标记
        for tr in rows:
            trpr = tr.find(qn("w:trPr"))
            if trpr is not None and trpr.find(qn("w:tblHeader")) is not None:
                has_header = True
                break
        # 第一行 run 加粗
        if not has_header and first_cells:
            for tc in first_cells:
                for r in tc.iter(qn("w:r")):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is not None and rpr.find(qn("w:b")) is not None:
                        has_header = True
                        break
                if has_header:
                    break
    return n_rows, n_cols, has_header


def has_footnote_ref(p_element) -> bool:
    return p_element.find(".//" + qn("w:footnoteReference")) is not None


def has_section_break(p_element) -> bool:
    """段落的 pPr 里是否有 sectPr（分节符）。"""
    ppr = p_element.find(qn("w:pPr"))
    if ppr is None:
        return False
    return ppr.find(qn("w:sectPr")) is not None