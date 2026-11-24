# ② 文档接入与「资产保护」层（Ingestion）

> 这是**防翻车的核心**。直接读取 `.docx` 的底层 XML，把文章拆成一个个**块（Block）**，而不是粗暴转成纯文本。

---

## 1. 为什么不能用「纯文本化」方案

若把 `.docx` 转成纯文本再排版，会永久丢失：

- 公式（OMML 会变成乱码或丢失）
- 图片位置、表格结构、脚注引用
- 上下标、域、书签、交叉引用

因此 TF 采用**块级（Block-level）解析**：保留每个块的原始 OOXML 元素引用，解析阶段只「读取并分类」，渲染阶段才「选择性修改」。

---

## 2. 核心抽象：Block

每个 Block 是文档流中的一个语义单元，同时携带：

- **原始 XML 节点引用**（用于渲染时定位/修改，甚至原样保留）
- **资产标志**（是否为保护资产）
- **语义标签**（解析阶段填充）

```python
@dataclass
class Block:
    block_id: str                  # 稳定 ID，如 "b0001"
    index: int                     # 文档流顺序
    kind: BlockKind                # paragraph | table | image | formula | footnote_ref | section_break | toc
    protected: bool                # 是否受保护（不改内部结构）
    asset: AssetInfo | None        # 资产信息（图片尺寸、表格维度、公式类型等）
    text: str                      # 纯文本（段落文本；保护块可为空或占位）
    runs: list[RunInfo]            # run 级信息（字体/字号/上下标/加粗）
    xml_element: Any               # 原始 OOXML 元素（lxml element）
    parent_paragraph: Any | None   # 所属段落（表格/图片挂在段落下）
    section_index: int             # 所属分节
    semantic_tag: SemanticTag | None  # 解析/校对后填充
    confidence: float = 0.0        # 分类置信度（正则/LLM）
    source: str = "regex"          # regex | llm | human | preset
```

> `xml_element` 是「资产保护」的物理基础：渲染时若块 `protected=True` 且不需改外框，直接保留该元素；需改外框时只动其 `inline/anchor` 与尺寸属性，**绝不触碰内部结构**。

---

## 3. Block 分类（Kind）

| Kind | 含义 | 是否保护 | 渲染时行为 |
|------|------|---------|-----------|
| `paragraph` | 普通文本段落 | 否 | 套样式、字体分离 |
| `heading` | 标题段落（解析后确定） | 否 | 套 Heading N |
| `image` | 图片（inline/anchor） | **是** | 只调尺寸/居中/锚点 |
| `table` | 表格 | **是**（结构） | 转三线表 / 调字号/重复表头 |
| `formula` | OMML 公式 / MathType OLE | **是** | 保留内部，仅调整对外框/编号/行距 |
| `formula_number` | 公式编号 | 否 | 右对齐 |
| `footnote_ref` | 脚注引用 | **是** | 原样保留 |
| `subscript/superscript` | 上下标 run | **是** | 原样保留（不转普通文本） |
| `section_break` | 分节符 | **是** | 保留，映射到 `page.sections` |
| `toc` | 目录域 | **是** | 保留，收尾刷新 |
| `caption` | 图/表题注 | 否 | 套 Caption |

---

## 4. 接入流程

```
ingest(docx_path) -> list[Block]
  │
  ├─ 1. 解压 .docx，读取 word/document.xml
  ├─ 2. 遍历 body 顶层子元素（w:p / w:tbl / w:sectPr）
  ├─ 3. 对每个元素构造 Block：
  │     ├─ 段落：提取 runs（文本/字体/上下标/加粗）
  │     ├─ 表格：记录维度、单元格文本，标记 protected
  │     ├─ 图片：从 w:drawing/w:pict 提取，标记 protected + 尺寸
  │     ├─ 公式：识别 m:oMathPara / m:oMath / w:object(OLE) → formula
  │     ├─ 分节：w:sectPr → section_break，递增 section_index
  │     └─ 域：w:fldSimple / instrText 含 "TOC" → toc
  ├─ 4. 保留每个块对原始 xml 元素的引用
  └─ 5. 返回 Block 列表（文档流顺序）
```

---

## 5. 资产识别细节

### 5.1 图片（image）
- 位置：`w:drawing/wp:inline` 或 `w:drawing/wp:anchor`（也可能是旧式 `w:pict/v:shape`）。
- 提取：`wp:extent` 得原始尺寸（EMU），转 cm；`a:blip/@r:embed` 得关系 ID。
- 保护：渲染只改 `wp:extent` 与 `wp:effectExtent`，**不动 blip 引用**。

### 5.2 表格（table）
- 提取：行×列维度、合并信息（`w:gridSpan`/`w:vMerge`）、首行是否为表头。
- 三线表转换：由渲染层执行（见 §05），Ingestion 只负责标注「这是个表格 + 维度 + 是否有表头」。

### 5.3 公式（formula）
- **OMML**：`m:oMathPara`（独立公式段）或 `m:oMath`（行内）。保留整段 OMML 树。
- **MathType**：以 OLE 对象（`w:object` + `o:OLEObject`）或 `w:pict` 呈现。这类是「疑似 MathType」，渲染时**只告警不改**，除非用户确认。
- 判定：段落中 `m:oMath` 存在 → formula；`w:object`/`m:oMathPara` → formula；纯 `w:pict` 且非图片 → 疑似 formula，标 `flag_suspect_matex`。

### 5.4 脚注 / 上下标
- 脚注引用 `w:footnoteReference`：原样保留。
- 上下标：run 的 `w:vertAlign w:val="superscript|subscript"`：保留 vertAlign，**不合并、不转普通文本**。

---

## 6. 接口契约（详见 §06）

```python
class IngestionEngine(Protocol):
    def ingest(self, docx_path: Path) -> list[Block]: ...
    def extract_assets(self, block: Block) -> AssetInfo | None: ...
```

---

## 7. 边界与容错

- 损坏的 XML / 不支持的元素：记入 `warnings`，该块降级为 `paragraph`（保留纯文本），不中断流水线。
- 超大文档：分块流式处理，避免一次性加载全部 DOM（可选优化）。
- 只读保护：Ingestion 阶段**绝不修改**源文档，所有修改在渲染阶段基于原 XML 副本进行。

---

## 8. 竞品参考

- 块级解析与公式保留参考 [docx-formatter-skill](https://github.com/CP-here/docx-formatter-skill) 与 [dsh-formula-docx](https://github.com/Red-Rayux/dsh-formula-docx)（OMML 原生公式处理）。