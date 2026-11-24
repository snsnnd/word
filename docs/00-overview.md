# 论文排版工具（Thesis Formatter）总览

> 工作代号：`thesis-formatter`（下称 **TF**）
> 目标：把「母版 + 规则表 + 智能解析 + 人工校对 + 原生渲染」串成一条可复用流水线，专治论文/公文/期刊投稿这类**高度公式化的排版**需求。

---

## 1. 我们要解决的核心矛盾

Word 排版之所以痛苦，是因为**内容与表现深度耦合**：

- 一篇论文里混着中文、英文、公式（OMML）、MathType 对象、图片、表格、脚注、上下标、域（TOC/交叉引用）。
- 格式规范又是**成体系、强规则**的：字体、字号、行距、首行缩进、三线表、图表题、页码分节。

直接「从零画 Word」既低效又易翻车。因此 TF 的设计哲学是：

> **不要重新发明 Word。要接管并「结构化」它。**

把 Word 文档拆成「块（Block）」，给块打标签，按标签套母版样式，最后在 Word 原生层做收尾。

---

## 2. 一句话架构

```
母版 .docx  +  规则表 .json
        │
        ▼
┌─────────────────────────────────────────────┐
│  ① Template Engine  载入母版 + 解析规则        │
│  ② Ingestion        读 .docx → Block 列表     │
│                     (保护图片/表格/公式/脚注)  │
│  ③ Structure Parser 正则硬规则 + LLM 兜底     │
│                     → 每块一个语义标签         │
│  ④ Review Console   树状大纲 + 人工纠偏        │
│  ⑤ Render Engine    灌入母版 + 套样式 + 后处理│
│                     + Word 原生收尾(Windows)   │
└─────────────────────────────────────────────┘
        │
        ▼
   成品 .docx  (+ 可选 .pdf)  +  质检日志
```

每一层都是**可独立测试、可插拔替换**的模块，通过统一的数据契约（`Block`、`StyleRule`、`PipelineResult`）通信。

---

## 3. 文档地图

| 文档 | 主题 | 读者 |
|------|------|------|
| [`01-template-engine.md`](./01-template-engine.md) | 模板引擎：母版约定 + JSON 规则表 Schema | 架构/实现 |
| [`02-ingestion.md`](./02-ingestion.md) | 文档接入与资产保护：Block 拆分、资产识别 | 实现 |
| [`03-structure-parser.md`](./03-structure-parser.md) | 结构解析引擎：正则 + LLM 混合分类 | 实现/算法 |
| [`04-review-console.md`](./04-review-console.md) | 轻量校对台：树状大纲 + 人工纠偏 | 实现/前端 |
| [`05-render-engine.md`](./05-render-engine.md) | 渲染与后处理：灌母版 + 后处理 + Word 收尾 | 实现 |
| [`06-interfaces-api.md`](./06-interfaces-api.md) | 全系统接口契约（数据模型 + 各层 API） | **全员必读** |
| [`07-project-structure.md`](./07-project-structure.md) | 目录结构、依赖、测试与里程碑 | 全员 |
| [`08-open-questions.md`](./08-open-questions.md) | 待决问题与设计权衡 | 评审 |
| [`09-handover.md`](./09-handover.md) | **交接说明**：现状、环境、代码地图、已知限制、下一步 | **接手者必读** |

**阅读顺序建议**：先读 `06`（数据契约）与 `07`（结构），再按 `01→05` 看各模块细节。

---

## 4. 关键设计原则（贯穿所有文档）

1. **母版优先（Master-first）**：绝不硬编码样式参数。所有视觉规则来自母版 + 规则表。
2. **资产保护（Asset-protection）**：图片/表格/公式/脚注/上下标在解析阶段就被「冻结」为原子资产，渲染阶段只改外框，绝不改内部结构。
3. **两层解析（Regex → LLM）**：正则先筛掉 90% 明显的块，剩下的模糊块才交给 LLM，兼顾速度与准确率、控制成本。
4. **人工兜底（Human-in-the-loop）**：正式渲染前必须有一个 30 秒可完成的校对台，把「自动化效率」和「人工兜底」结合得最丝滑。
5. **原生收尾（Native finalization）**：目录域、交叉引用、页码只能由 Word 自己刷新；能纯 Python 做的绝不依赖 Word，依赖 Word 的必须是可选增强路径。
6. **可观测（Observable）**：每次运行产出结构化质检日志（处理了多少图/表、哪些疑似公式待人工核对）。
7. **幂等 & 可重放**：同一份「文章 + 母版 + 规则 + 校对修正」重复运行，结果一致。

---

## 5. 与现有开源方案的关系（先看竞品，再动手）

调研结论（详见各模块引用）：已有若干相关项目，但多为「Markdown → docx」或「单一模板替换」，**缺少本工具强调的「母版+规则+AI解析+人工校对+资产保护」完整闭环**：

- **[OpenThesis](https://github.com/Gluxggg/OpenThesis)**：AI 驱动的模板引擎，解析 .docx 模板并导出成品（论文/期刊/公文）。思路最接近，但其核心是「解析模板套内容」，本工具更强调**输入文档的分块语义解析 + 资产保护 + 人工校对台**。
- **[format-thesis-docx](https://github.com/wksudud/format-thesis-docx)**：DOCX 论文格式自动化，从要求中抽取规则、审计稿件、生成修正副本。与本文档的「规则表 + 质检日志」思路一致，可借鉴其规则抽取方式。
- **[docx-formatter-skill](https://github.com/CP-here/docx-formatter-skill)**：中文文档格式化 skill，支持 OMML 公式、三线表、目录与引用上标。三线表、OMML 处理可借鉴。
- **[Red-Rayux/dsh-formula-docx](https://github.com/Red-Rayux/dsh-formula-docx)**：Markdown → 原生 Word 公式（OMML）+ PDF 校验。
- **[ruledoc](https://pypi.org/project/ruledoc/)**、**[latex-docx](https://pypi.org/project/latex-docx/)**：规则驱动/LaTeX 公式转 docx 的参考实现。

> 结论：TF 不重造这些能力，而是**把它们组合成一条带「人工校对台」的流水线**，并补上「保护块」这一防翻车核心。

---

## 6. 技术栈选型（草案）

| 关注点 | 选型 | 理由 |
|--------|------|------|
| 语言 | **Python 3.10+** | docx 生态最成熟 |
| 文档读写 | `python-docx` + 直接操作 OOXML（`lxml`） | 前者快速搭骨架，后者处理公式/域/资产 |
| 母版样式 | OOXML `styles.xml` + 规则表 JSON | 母版定基调，规则表补硬规则 |
| 公式识别 | OMML（`<m:oMath>`）、MathType OLE 对象 | 原生解析，不转纯文本 |
| 正则解析 | 内置 `re` | 快、免费、可解释 |
| LLM 兜底 | 轻量模型（ERNIE-5.5-Flash 等） | 便宜快速，输出严格 JSON |
| 渲染收尾 | `pywin32`（Windows 可选）+ 内嵌 VBA | 刷新 TOC/交叉引用、导出 PDF |
| 校对台 | 本地 Web（轻量前端）或 CLI 树 | 30 秒可完成人工纠偏 |
| 配置 | JSON Schema 校验 | 规则表可校验、可复用 |

> 注：Linux/无 Word 环境下，渲染收尾逐级降级 —— 有 LibreOffice 则 headless 导出 PDF 预览（目录域不刷新，`degraded=True`），都没有则只生成 docx + 提示用户手动更新域。