# ⑥ 全系统接口契约（Interfaces & API）

> **全员必读。** 本文档定义整个 TF 的数据模型、各层协议（Protocol）与顶层编排 API。所有模块实现这些接口即可插拔替换、独立测试。

版本：契约版本 `0.1.0`（草案）。

---

## 1. 数据模型（`tf.core.models`）

### 1.1 枚举

```python
from enum import Enum

class BlockKind(str, Enum):
    paragraph = "paragraph"
    heading = "heading"
    image = "image"
    table = "table"
    formula = "formula"
    formula_number = "formula_number"
    footnote_ref = "footnote_ref"
    section_break = "section_break"
    toc = "toc"

class BlockType(str, Enum):
    cover_title = "cover_title"
    cover_meta = "cover_meta"
    abstract_title = "abstract_title"
    abstract_body = "abstract_body"
    keywords = "keywords"
    toc = "toc"
    heading_1 = "heading_1"
    heading_2 = "heading_2"
    heading_3 = "heading_3"
    heading_4 = "heading_4"
    body = "body"
    caption = "caption"
    reference = "reference"
    reference_title = "reference_title"
    appendix_title = "appendix_title"
    image = "image"
    table = "table"
    formula = "formula"
    footnote = "footnote"
    unknown = "unknown"

class BlockSource(str, Enum):
    regex = "regex"
    llm = "llm"
    human = "human"
    preset = "preset"
    unknown = "unknown"
```

### 1.2 块与资产

```python
from dataclasses import dataclass, field

@dataclass
class RunInfo:
    text: str
    font_cn: str | None = None
    font_en: str | None = None
    size_pt: float | None = None
    bold: bool = False
    italic: bool = False
    vert_align: str | None = None   # "superscript" | "subscript" | None

@dataclass
class AssetInfo:
    kind: BlockKind
    width_cm: float | None = None
    height_cm: float | None = None
    rows: int | None = None          # table
    cols: int | None = None          # table
    has_header: bool | None = None   # table
    omath: bool | None = None        # formula: is OMML
    ole: bool | None = None          # formula: is MathType OLE
    rel_id: str | None = None        # image relationship id

@dataclass
class Block:
    block_id: str                     # 稳定 ID（见 §1.4）
    index: int
    kind: BlockKind
    protected: bool
    asset: AssetInfo | None = None
    text: str = ""
    runs: list[RunInfo] = field(default_factory=list)
    section_index: int = 0
    semantic_tag: BlockType | None = None
    confidence: float = 0.0
    source: BlockSource = BlockSource.unknown
    # 渲染期使用，不序列化
    xml_element: Any | None = None
    parent_paragraph: Any | None = None

    def needs_llm(self) -> bool: ...
```

### 1.3 样式与规则

```python
@dataclass
class Style:
    font_cn: str | None = None
    font_en: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    align: str | None = None          # left|center|right|justify
    line_spacing: float | None = None
    first_line_indent_chars: float | None = None
    hanging_indent_chars: float | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None

@dataclass
class RuleSet:
    defaults: Style
    styles: dict[BlockType, Style]
    assets: dict
    postprocess: dict
    page: dict
    finalization: dict

    def resolve(self, block_type: BlockType) -> Style:
        """返回 defaults 叠加 styles[block_type] 后的最终样式。"""
```

### 1.4 `block_id` 稳定性（重放的关键）

`block_id` 由 Ingestion 生成，必须**跨运行稳定**，否则校对台的修正无法重放：

```
block_id = "b" + zero_pad(index, 4) + "_" + short_hash(块文本前64字 + 块XML结构摘要)
```

- `index`：文档流顺序（保证同文档稳定）。
- `short_hash`：内容指纹（保证即使顺序微调也能对齐）。
- 校对台以 `block_id` 为键回写修正；渲染前先按 `block_id` 应用修正。

### 1.5 模板规格

```python
@dataclass
class TemplateSpec:
    master_path: Path
    master_styles: dict[str, Style]
    rules: RuleSet
    placeholders: dict[str, str]
    warnings: list[str] = field(default_factory=list)
```

### 1.6 结果模型

```python
@dataclass
class RenderResult:
    docx_path: Path
    pdf_path: Path | None
    stats: dict          # {"images": 14, "tables": 5, "formulas": 23, ...}
    warnings: list[str]
    degraded: bool       # 是否发生 LLM/Word 降级

@dataclass
class FinalizeResult:
    updated_toc: bool
    updated_fields: bool
    exported_pdf: bool
    engine_used: str     # win32com | vba | none

@dataclass
class Fix:
    block_id: str
    from_type: BlockType
    to_type: BlockType
    by: str              # human
    at: str              # ISO8601
```

---

## 2. 各层协议（Protocol）

```python
from typing import Protocol

class TemplateEngine(Protocol):
    def load(self, master_path: Path, rules_path: Path,
             metadata: dict | None = None) -> TemplateSpec: ...

class IngestionEngine(Protocol):
    def ingest(self, docx_path: Path) -> list[Block]: ...

class StructureParser(Protocol):
    def parse(self, blocks: list[Block], spec: TemplateSpec) -> list[Block]: ...

class ReviewConsole(Protocol):
    def build_outline(self, blocks: list[Block]) -> "OutlineNode": ...
    def apply_fixes(self, blocks: list[Block], fixes: list[Fix]) -> list[Block]: ...
    def load_fixes(self, doc_id: str) -> list[Fix]: ...
    def save_fixes(self, doc_id: str, fixes: list[Fix]) -> None: ...

class RenderEngine(Protocol):
    def render(self, blocks: list[Block], spec: TemplateSpec,
               out_path: Path) -> RenderResult: ...
    def finalize(self, docx_path: Path, spec: TemplateSpec) -> FinalizeResult: ...
```

---

## 3. 顶层编排 API（`tf.pipeline`）

```python
@dataclass
class PipelineConfig:
    docx_path: Path
    master_path: Path
    rules_path: Path
    metadata: dict | None = None
    out_dir: Path = Path("out")
    doc_id: str | None = None
    use_llm: bool = True
    use_review: bool = True      # 是否启用人机校对（CLI 可跳过）
    llm: dict | None = None      # LLM 连接参数
    finalization: str = "auto"

class ThesisFormatter:
    def __init__(self, cfg: PipelineConfig): ...

    def run(self) -> RenderResult:
        """完整流水线：ingest → parse → (review) → render → finalize。"""
        ...

    def run_stages(self) -> dict:
        """分步执行，返回每阶段中间产物（调试用）。"""
        ...

# 便捷入口
def format_document(cfg: PipelineConfig) -> RenderResult:
    return ThesisFormatter(cfg).run()
```

编排逻辑：

```
run():
  spec   = template_engine.load(master, rules, metadata)
  blocks = ingestion.ingest(docx)
  blocks = parser.parse(blocks, spec)
  if cfg.use_review:
      fixes = review.load_fixes(doc_id)
      if fixes: blocks = review.apply_fixes(blocks, fixes)
  result = render.render(blocks, spec, out_dir/docx)
  if spec.rules.finalization: fin = render.finalize(result.docx_path, spec)
  return result
```

---

## 4. 命令行接口（`tf.cli`）

```bash
# 一键全流程（CLI 无校对台时跳过人工，直接用解析结果）
tf run --doc thesis.docx --master master.docx --rules rules.json --out out/

# 分步调试
tf ingest --doc thesis.docx --dump blocks.json
tf parse  --blocks blocks.json --rules rules.json --dump tagged.json
tf review --blocks tagged.json            # 启校对台
tf render --blocks tagged.json --master master.docx --rules rules.json --out out/

# 质检
tf report --out out/report.json
```

---

## 5. HTTP 接口（校对台 Web 形态）

| 方法 | 路径 | 说明 |
|------|------|------|
| GET  | `/api/review/{doc_id}` | 获取大纲 JSON |
| POST | `/api/review/{doc_id}/fixes` | 提交修正 |
| POST | `/api/review/{doc_id}/render` | 触发渲染 |
| GET  | `/api/report/{doc_id}` | 获取质检报告 |

大纲 JSON 结构：

```jsonc
{
  "doc_id": "毕业论文_v3",
  "root": {
    "type": "heading_1", "block_id": "b0001_1a2b", "text": "第一章 绪论",
    "children": [
      { "type": "heading_2", "block_id": "b0002_3c4d", "text": "1.1 研究背景",
        "confidence": 0.98, "source": "regex", "children": [] },
      { "type": "body", "block_id": "b0032_9f8e", "text": "本文提出了一种……",
        "confidence": 0.62, "source": "llm", "suspect": true, "children": [] }
    ]
  }
}
```

---

## 6. 版本与兼容

- 契约版本号 `0.1.0`，破坏性变更升 minor。
- 所有模型 dataclass 提供 `to_dict()/from_dict()`，JSON 序列化用于 blocks.json / 大纲 / 报告。
- `xml_element` 不序列化（仅内存态）。