# ① 模板引擎（Template Engine）

> 原则：**不要让程序从零画一个 Word。** 一切视觉规则来自「母版 + 规则表」。

模板引擎的职责：加载母版 `.docx`、解析母版里预置的样式，并校验/合并规则表 JSON，产出统一的 `TemplateSpec` 供下游使用。

---

## 1. 母版（Master .docx）的约定

母版是人工在 Word 里精修好的 `.docx`，TF 只「读取并复用」它，**不回写修改母版本体**（避免污染源）。母版需预先包含：

### 1.1 结构元素
- **封面**：标题、作者、导师、日期等占位（用连续书签或约定占位文本，如 `{{TITLE}}`、`{{AUTHOR}}`）。
- **目录占位**：一个带 `TOC` 域的段落（Word 里「引用→目录」插入即可）。
- **分节符**：封面/摘要/正文/参考文献之间的分节（决定页码格式：罗马数字 vs 阿拉伯数字）。
- **页眉页脚**：奇偶页不同、首页不同等。
- **内置样式**：`Heading 1/2/3`、`Normal（正文）`、`Caption（图题）`、`TOC 1/2/3` 等。

### 1.2 内置样式清单（TF 依赖这些样式名）
TF **只认 Word 内置样式名**（styleId），不认自定义别名，避免跨文档漂移：

| 用途 | 内置样式名（styleId） | 说明 |
|------|----------------------|------|
| 一级标题 | `Heading 1` | 章标题 |
| 二级标题 | `Heading 2` | 节标题 |
| 三级标题 | `Heading 3` | 小节标题 |
| 正文 | `Normal` | 段落正文 |
| 图题 | `Caption` | 图/表题注 |
| 目录项 | `TOC 1` / `TOC 2` / `TOC 3` | 目录条目 |
| 参考文献 | `Normal` + 规则表指定悬挂缩进 | 或用自定义 `Bibliography` 但需在母版预置 |

> 若母版缺某个样式，模板引擎在加载时**警告**并给出「建议在 Word 中补齐」的提示，而不是静默失败。

### 1.3 占位符约定
封面等可变文本使用 `{{...}}` 占位，渲染时由「元数据 JSON」替换：

```json
{ "title": "……", "author": "……", "advisor": "……", "date": "2026-11" }
```

---

## 2. 规则表（Rules JSON）Schema

规则表描述**程序必须执行的硬规则**（母版搞不定或不想污染母版的部分）。它**与具体论文无关**，可跨文档复用（例如「本科毕业论文规范」「某期刊投稿规范」各一份）。

### 2.1 顶层结构

```jsonc
{
  "schema_version": "1.0",
  "name": "本科毕业论文规范 v1",
  "defaults": { /* 全局默认，见下 */ },
  "styles": { /* 按块类型覆盖，见下 */ },
  "assets": { /* 图片/表格/公式处理，见下 */ },
  "postprocess": { /* 后处理开关，见下 */ },
  "page": { /* 页面与分节，见下 */ },
  "finalization": { /* Word 原生收尾，见下 */ }
}
```

### 2.2 `defaults`（全局默认）

```jsonc
"defaults": {
  "body_font_cn": "宋体",
  "body_font_en": "Times New Roman",
  "body_font_size_pt": 12,
  "line_spacing": 1.5,
  "first_line_indent_chars": 2,      // 首行缩进字符数（中文按字符）
  "space_before_pt": 0,
  "space_after_pt": 0
}
```

### 2.3 `styles`（按块类型覆盖）

键是**语义标签**（与 §06 的 `BlockType` 对齐），值是覆盖 `defaults` 的样式片段：

```jsonc
"styles": {
  "heading_1": { "font_cn": "黑体", "font_en": "Arial", "size_pt": 16, "bold": true, "align": "center", "space_before_pt": 24, "space_after_pt": 12 },
  "heading_2": { "font_cn": "黑体", "size_pt": 14, "bold": true, "align": "left" },
  "heading_3": { "font_cn": "黑体", "size_pt": 12, "bold": true },
  "body":      { "font_cn": "宋体", "size_pt": 12, "first_line_indent_chars": 2, "line_spacing": 1.5, "align": "justify" },
  "caption":   { "font_cn": "宋体", "size_pt": 10.5, "align": "center", "bold": false },
  "reference": { "font_cn": "宋体", "size_pt": 10.5, "hanging_indent_chars": 2, "align": "left" },
  "abstract":  { "font_cn": "宋体", "size_pt": 12, "line_spacing": 1.5 },
  "keywords":  { "font_cn": "宋体", "size_pt": 12, "bold_prefix": true }
}
```

### 2.4 `assets`（资产处理）

```jsonc
"assets": {
  "image": {
    "max_width_cm": 12,            // 图片最大宽度
    "align": "center",            // 图片居中
    "keep_aspect_ratio": true,    // 保持宽高比
    "anchor": "inline"            // inline 优先，避免浮动错乱
  },
  "table": {
    "style": "three_line",        // three_line | keep_as_is | grid
    "font_size_pt": 10.5,
    "header_bold": true,
    "repeat_header_row": true     // 跨页重复表头
  },
  "formula": {
    "mode": "keep_as_is",         // keep_as_is | normalize_spacing
    "numbered": true,             // 公式是否带编号
    "number_align": "right",      // 编号右对齐
    "flag_suspect_matex": true    // 疑似 MathType 公式仅告警不改
  }
}
```

### 2.5 `postprocess`（后处理开关）

```jsonc
"postprocess": {
  "remove_empty_paragraphs": true,   // 清多余空行（保留分节/分页处）
  "merge_split_runs": true,          // 合并被拆散的 run（修复字体错乱）
  "normalize_whitespace": true,      // 中英文间空格规范（可选）
  "punct_convert": false             // 标点转换（保守默认关闭）
}
```

### 2.6 `page`（页面与分节）

```jsonc
"page": {
  "size": "A4",
  "margins_cm": { "top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5 },
  "sections": [
    { "name": "cover",      "page_number_format": "none" },
    { "name": "frontmatter","page_number_format": "upper_roman" },  // 摘要/目录
    { "name": "body",       "page_number_format": "decimal", "start": 1 },
    { "name": "references", "page_number_format": "decimal", "continue": true }
  ]
}
```

### 2.7 `finalization`（Word 原生收尾）

```jsonc
"finalization": {
  "update_toc": true,            // 刷新目录域
  "update_fields": true,         // 刷新交叉引用等域
  "export_pdf": true,            // 导出 PDF 预览
  "engine": "auto"               // auto | win32com | vba | none
}
```

---

## 3. TemplateSpec 数据模型

模板引擎输出的核心产物，供渲染层使用：

```python
@dataclass
class TemplateSpec:
    master_path: Path            # 母版文件路径
    master_styles: dict[str, Style]   # 从母版读到的样式字典
    rules: RuleSet               # 合并后的规则表
    placeholders: dict[str, str] # 占位符映射（封面元数据）
    warnings: list[str]          # 加载期告警（如缺样式）
```

`RuleSet` 是规则表 JSON 经 JSON Schema 校验后的强类型对象，提供 `resolve(block_type) -> Style` 的合并方法（`defaults` ⊕ `styles[type]`）。

---

## 4. 加载流程

```
load_template(master_path, rules_path, metadata=None)
  │
  ├─ 1. 打开母版 .docx，解析 styles.xml → master_styles
  ├─ 2. 校验 styles 字典完整性，缺样式 → warnings
  ├─ 3. 读取 rules.json，JSON Schema 校验
  ├─ 4. 合并 rules（defaults ⊕ styles）
  ├─ 5. 加载占位符元数据（可选，缺省留空）
  └─ 6. 返回 TemplateSpec
```

---

## 5. 接口契约（详见 §06）

```python
class TemplateEngine(Protocol):
    def load(self, master_path: Path, rules_path: Path,
             metadata: dict | None = None) -> TemplateSpec: ...
```

---

## 6. 竞品参考

- 母版解析套内容思路参考 [OpenThesis](https://github.com/Gluxggg/OpenThesis)（模板引擎）。
- 规则抽取参考 [format-thesis-docx](https://github.com/wksudud/format-thesis-docx)（从要求中抽取规则）。
- 三线表/OMML 细节参考 [docx-formatter-skill](https://github.com/CP-here/docx-formatter-skill)。