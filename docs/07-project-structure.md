# ⑦ 项目结构、依赖、测试与里程碑

---

## 1. 目录结构

```
thesis-formatter/
├── docs/                          # 本设计文档
│   ├── 00-overview.md
│   ├── 01-template-engine.md
│   ├── 02-ingestion.md
│   ├── 03-structure-parser.md
│   ├── 04-review-console.md
│   ├── 05-render-engine.md
│   ├── 06-interfaces-api.md       # 接口契约（全员必读）
│   ├── 07-project-structure.md
│   └── 08-open-questions.md
│
├── pyproject.toml                 # 依赖与打包
├── README.md
│
├── tf/                            # 主包
│   ├── __init__.py
│   ├── core/
│   │   ├── models.py              # Block/Style/RuleSet/TemplateSpec/... (见 06)
│   │   ├── enums.py               # BlockKind/BlockType/BlockSource
│   │   └── idgen.py               # block_id 生成（稳定）
│   │
│   ├── template/                  # ① 模板引擎
│   │   ├── engine.py              # TemplateEngine.load
│   │   ├── master.py              # 母版 styles.xml 解析
│   │   └── rules.py               # RuleSet 加载 + JSON Schema 校验 + resolve
│   │
│   ├── ingestion/                 # ② 文档接入与资产保护
│   │   ├── engine.py              # IngestionEngine.ingest
│   │   ├── ooxml.py               # document.xml 遍历
│   │   └── assets.py              # 图片/表格/公式/脚注识别
│   │
│   ├── parser/                    # ③ 结构解析
│   │   ├── regex_parser.py        # 第一层正则
│   │   ├── llm_parser.py          # 第二层 LLM 兜底
│   │   ├── consistency.py         # 一致性校验（层级/编号/配对）
│   │   └── rules_default.json     # 内置正则规则表
│   │
│   ├── review/                    # ④ 校对台
│   │   ├── console.py             # build_outline / apply_fixes / 持久化
│   │   ├── web.py                 # 轻量 Web（FastAPI/Flask）
│   │   ├── static/                # 前端（树状大纲 UI）
│   │   └── fixes_store.py         # review_fixes.json 读写
│   │
│   ├── render/                    # ⑤ 渲染与后处理
│   │   ├── engine.py              # RenderEngine.render / finalize
│   │   ├── styler.py              # 套样式 + 字体分离
│   │   ├── assets_render.py       # 图片缩放/三线表/公式处理
│   │   ├── postprocess.py         # 清空行/合并 run/空格
│   │   ├── sections.py            # 分节与页码
│   │   └── finalize_win32.py      # Windows win32com 收尾（可选）
│   │
│   ├── pipeline.py                # 顶层编排 ThesisFormatter（见 06）
│   ├── cli.py                     # 命令行入口
│   └── report.py                  # 质检报告生成
│
├── schemas/                       # JSON Schema
│   ├── rules.schema.json
│   ├── fixes.schema.json
│   └── outline.schema.json
│
├── tests/
│   ├── fixtures/                  # 样例 docx / 母版 / 规则
│   │   ├── master_min.docx
│   │   ├── sample_paper.docx
│   │   └── rules_min.json
│   ├── test_template.py
│   ├── test_ingestion.py
│   ├── test_parser_regex.py
│   ├── test_parser_llm.py         # 可 mock
│   ├── test_review.py
│   ├── test_render.py
│   ├── test_postprocess.py
│   └── test_pipeline_e2e.py
│
└── examples/
    └── demo.sh                    # 端到端演示脚本
```

---

## 2. 依赖（`pyproject.toml` 草案）

```toml
[project]
name = "thesis-formatter"
version = "0.1.0"
requires-python = ">=3.10"
dependencies = [
  "python-docx>=1.1.0",     # 文档骨架
  "lxml>=5.0",              # OOXML 直接操作
  "jsonschema>=4.0",        # 规则表校验
  "pydantic>=2.0",          # 可选：强类型与序列化
]

[project.optional-dependencies]
llm     = ["openai>=1.0", "httpx>=0.27"]   # 轻量模型客户端（OpenAI 兼容协议）
web     = ["fastapi>=0.110", "uvicorn>=0.29"]
win32   = ["pywin32>=306"]                  # 仅 Windows
pdf     = ["docx2pdf>=0.1.8"]              # 可选 PDF（Windows/Word）
dev     = ["pytest>=8.0", "ruff>=0.5", "mypy>=1.10"]

[project.scripts]
tf = "tf.cli:main"
```

> 说明：`pywin32` 与 `docx2pdf` 仅 Windows 且可选；Linux 下走降级路径（见 §05）。

---

## 3. 关键外部依赖说明

| 依赖 | 用途 | 平台 |
|------|------|------|
| `python-docx` | 文档骨架读写 | 全平台 |
| `lxml` | 直接操作 OOXML（公式/域/资产） | 全平台 |
| 轻量 LLM（OpenAI 兼容 API） | 结构兜底分类 | 全平台（需网络） |
| `pywin32` | Word 原生收尾（TOC/域/PDF） | Windows |
| LibreOffice `soffice` | 降级 PDF 预览 | 全平台（可选） |

---

## 4. 测试策略

- **单元测试**：每个模块独立，LLM 与 Word 均 mock。
- **快照测试**：`render` 产出 docx → 解包比对 `styles.xml`/`document.xml` 关键片段。
- **端到端**：`tests/fixtures/sample_paper.docx` + `master_min.docx` + `rules_min.json` 跑通全流水线，断言 `RenderResult.stats` 与告警。
- **回归**：校对台重放——同一 `doc_id` 应用 `review_fixes.json` 后结果幂等。
- **契约测试**：校验各模块输出符合 §06 数据模型（`to_dict()` 可序列化）。

---

## 5. 里程碑（建议 4 个阶段）

### M1 — 骨架与契约（2–3 天）
- [x] `core/models.py`、`enums.py`、`idgen.py`
- [x] `template/engine.py` 能加载母版样式 + 规则表
- [x] `ingestion/engine.py` 能把简单 docx 拆成 Block
- [x] CLI `tf ingest` + 契约测试

### M2 — 解析与渲染闭环（3–4 天）
- [x] `regex_parser` + 内置规则
- [x] `llm_parser`（含降级）
- [x] `render/engine` 灌母版 + 套样式 + 字体分离 + 三线表 + 图片缩放
- [x] `report` 质检报告
- [x] `test_pipeline_e2e` 跑通（无 Word 环境）

### M3 — 人机校对与收尾（3–4 天）
- [x] `review/console` + 本地 Web 校对台
- [x] `fixes_store` 重放
- [x] `finalize_win32`（Windows 验证 TOC/域/PDF）
- [x] Linux 降级路径 + LibreOffice 预览

### M4 — 打磨与扩展（持续）
- [ ] 更多规则模板（期刊/公文）
- [ ] 规则表图形化编辑器
- [ ] MathType 公式转换（可选）
- [ ] 批量处理目录级联
- [ ] 性能优化（大文档流式）

---

## 6. 风险与缓解

| 风险 | 影响 | 缓解 |
|------|------|------|
| 母版样式名不统一 | 渲染错乱 | 只用内置 styleId + 加载期告警 |
| 公式（MathType）无法解析 | 排版失败 | 保护块只告警不改 + 质检清单提示 |
| Linux 无 Word 无法刷域 | PDF 目录不对 | 降级 + 明确提示手动 F9 |
| LLM 误判标题 | 结构错 | 正则优先 + 人工校对台兜底 |
| block_id 不稳定 | 重放失效 | 基于 index + 内容哈希（§06.1.4） |
| 合并单元格三线表 | 表错乱 | 保留 gridSpan/vMerge + 快照测试 |

---

## 7. 竞品参考汇总

- [OpenThesis](https://github.com/Gluxggg/OpenThesis) — AI 驱动模板引擎（论文/期刊/公文）
- [format-thesis-docx](https://github.com/wksudud/format-thesis-docx) — 论文格式自动化 + 规则抽取 + 审计
- [docx-formatter-skill](https://github.com/CP-here/docx-formatter-skill) — 中文格式化，OMML/三线表/目录
- [dsh-formula-docx](https://github.com/Red-Rayux/dsh-formula-docx) — Markdown→OMML + PDF 校验
- [ruledoc](https://pypi.org/project/ruledoc/) / [latex-docx](https://pypi.org/project/latex-docx/) — 规则驱动 / LaTeX 公式转 docx