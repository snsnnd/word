# ③ 结构解析引擎（Structure Parser）

> 把所有文本块分类：封面、摘要、一级/二级/三级标题、正文、图注、参考文献、关键词……

策略：**两层混合**——第一层正则硬规则（快、免费），第二层 LLM 兜底（准）。

---

## 1. 语义标签体系（SemanticTag）

`SemanticTag` 是解析层与渲染层之间的「通用语言」。渲染层只认标签套样式，不关心文本内容。

```python
class BlockType(str, Enum):
    # 结构
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
    caption = "caption"          # 图/表题
    reference = "reference"
    reference_title = "reference_title"
    appendix_title = "appendix_title"
    # 非文本块
    image = "image"
    table = "table"
    formula = "formula"
    footnote = "footnote"
    unknown = "unknown"
```

> `heading_1/2/3` 对应母版内置 `Heading 1/2/3`；`body` 对应 `Normal`。

---

## 2. 第一层：正则硬规则

正则规则是**有序**的（先匹配先生效），每条规则输出 `(BlockType, confidence, need_llm)`。

```jsonc
[
  { "name": "chapter_cn", "pattern": "^第[一二三四五六七八九十百零〇]+章", "type": "heading_1", "confidence": 0.98 },
  { "name": "chapter_num", "pattern": "^Chapter\\s+\\d+",              "type": "heading_1", "confidence": 0.95, "flags": "I" },
  { "name": "section_cn",  "pattern": "^第[一二三四五六七八九十百零〇]+节", "type": "heading_2", "confidence": 0.95 },
  { "name": "section_num", "pattern": "^\\d+\\.\\d+\\s",               "type": "heading_2", "confidence": 0.9 },
  { "name": "subsection_num", "pattern": "^\\d+\\.\\d+\\.\\d+\\s",     "type": "heading_3", "confidence": 0.9 },
  { "name": "figure_caption", "pattern": "^图\\s?\\d+",               "type": "caption", "confidence": 0.95 },
  { "name": "table_caption",  "pattern": "^表\\s?\\d+",               "type": "caption", "confidence": 0.95 },
  { "name": "reference_item", "pattern": "^\\[\\d+\\]",               "type": "reference", "confidence": 0.9 },
  { "name": "keywords",       "pattern": "^(关键词|关键字|Key\\s?words)", "type": "keywords", "confidence": 0.95, "flags": "I" },
  { "name": "abstract_cn",    "pattern": "^摘\\s?要",                 "type": "abstract_title", "confidence": 0.95 },
  { "name": "abstract_en",    "pattern": "^Abstract",                "type": "abstract_title", "confidence": 0.95, "flags": "I" },
  { "name": "reference_title","pattern": "^参考文献",                  "type": "reference_title", "confidence": 0.95 },
  { "name": "appendix",       "pattern": "^附\\s*录",                 "type": "appendix_title", "confidence": 0.9 }
]
```

**正则层的启发式规则**（无需配置，硬编码在 `RegexParser`）：
- 段落整体加粗 + 字数 ≤ 30 + 不以句号结尾 → 疑似标题，`need_llm=True`。
- 段落字数 ≤ 15 且独占一行 → 疑似标题，交 LLM。
- 位于文档前部（封面区）的短段 → 疑似 cover 元素。

---

## 3. 第二层：LLM 兜底

### 3.1 何时调用 LLM
仅对满足以下任一条件的块调用，避免浪费：
- 正则未命中，且 `kind == paragraph`；
- 正则命中但 `confidence < 阈值`（如 0.9）；
- 启发式标记 `need_llm=True`；
- 文本较短（≤ 60 字）但疑似标题。

### 3.2 调用方式（严格 JSON 输出）

把一批待判定块打包（带 `block_id`、文本、格式特征）发给轻量模型：

**Prompt 骨架**：
```
你是文档结构标注助手。判断每个块的语义类型，只能从给定枚举中选择。
输出 JSON 数组，元素形如 {"block_id": "b0012", "type": "heading_2"}。
若无法确定，type 用 "unknown"。不要输出任何解释文字。
[可选：给出上下文相邻块，帮助判断层级]
待判定块：[{block_id, text, bold, font_size_pt, align}, ...]
```

**响应约束**：
- 强制 JSON（若平台支持 json mode / response_format）。
- 校验返回：`block_id` 必须存在、`type` 必须在 `BlockType` 内，否则回退 `unknown` 并告警。

### 3.3 批量与降级
- 批大小可配（默认 30 块/批），失败自动重试一次。
- LLM 不可用/超时 → 该批全部回退到正则结果或 `body`，并在质检日志标注「LLM 降级」。

---

## 4. 上下文增强（可选，提升准确率）

对标题层级判断，LLM 可接收**相邻块**作为上下文，例如：
- 若前一块是 `heading_1`，当前疑似标题可能是 `heading_2`。
- 若当前块后跟 `body`，更可能是标题而非正文。

上下文窗口可配置（默认前后各 2 块）。

---

## 5. 后处理与一致性校验

解析完成后跑一遍一致性修正：
- **标题层级单调性**：不允许 `heading_3` 直接跟在 `heading_1` 后而无 `heading_2`（除非配置允许），违者告警。
- **编号连续性**：`heading_2` 的 `1.1`、`1.2` 编号应递增，错乱则告警。
- **图题/表题配对**：`caption` 前应有 `image`/`table`，否则告警。
- **参考文献块**：`reference` 只允许出现在 `reference_title` 之后。

修正结果写入 `Block.semantic_tag`、`confidence`、`source`。

---

## 6. 接口契约（详见 §06）

```python
class StructureParser(Protocol):
    def parse(self, blocks: list[Block], spec: TemplateSpec) -> list[Block]:
        """就地填充 semantic_tag / confidence / source，返回同一批块。"""
```

内部实现分两步：
- `RegexParser.parse(blocks) -> list[Block]`（第一层）
- `LLMParser.parse(blocks) -> list[Block]`（第二层，只处理 `needs_llm` 的块）

---

## 7. 设计权衡

| 方案 | 优点 | 缺点 | 采纳 |
|------|------|------|------|
| 纯正则 | 快、免费、可解释 | 覆盖率有限 | ✅ 第一层 |
| 纯 LLM | 准、泛化 | 慢、贵、不稳定 | ❌ |
| **正则+LLM** | 快 + 准，成本可控 | 需维护规则 | ✅ 本设计 |
| 规则+LLM+人工 | 最准 | 多一步 | ✅ 配合 §04 校对台 |

---

## 8. 竞品参考

- 规则抽取/分类思路参考 [format-thesis-docx](https://github.com/wksudud/format-thesis-docx)。
- AI 语义标注参考 [OpenThesis](https://github.com/Gluxggg/OpenThesis)（AI 驱动的模板解析）。