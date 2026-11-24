# ⑨ 交接说明（Handover）

> 交接范围：`thesis-formatter`（TF）当前代码库全部内容。
> 仓库：`git@github.com:snsnnd/word.git`
> 状态：M1–M3 主链路（接入 → 分类 → 渲染 → 收尾）已跑通，测试全绿。

---

## 1. 一句话现状

`ingest → classify → render → finalize` 四个阶段均可运行，CLI 端到端能产出
`.docx`，并在装有 LibreOffice 的机器上**真的导出 `.pdf`**（Word/WPS 环境走 COM 刷新目录域）。

```
tests/：230 collected → 228 passed, 2 xfailed（约 15s，含一次真实 LibreOffice 无头导出）
tf/：2032 行实现；tests/：2514 行测试
```

---

## 2. 环境与快速开始

```bash
# 1) 依赖（本机已预装在 .tf-deps，不是全局 site-packages）
export PYTHONPATH=/home/dsh/workspace/.tf-deps
#    python-docx 1.2.0 / lxml 6.1.3 / jsonschema 4.26.0；Python 3.10.12

# 2) 测试
cd /home/dsh/workspace/thesis-formatter
python3 -m pytest tests/ -q

# 3) 端到端跑一遍（无母版时用空白文档 + 内置规则）
python3 -m tf.cli ingest tests/fixtures/source_template.docx        # 块 JSON → stdout
python3 -m tf.cli parse  tests/fixtures/source_template.docx --out /tmp/blocks.json
python3 -m tf.cli render tests/fixtures/source_template.docx --template tests/fixtures/source_template.docx --out /tmp/out.docx
python3 -m tf.cli run    tests/fixtures/source_template.docx --out /tmp/out.docx   # 同一流水线
```

实测输出：

```
[run] 已生成 /tmp/smoke/out.docx（26/26 块）
[pdf] 已导出 /tmp/smoke/out.pdf（引擎：libreoffice）
  [finalize:libreoffice] 目录/交叉引用未刷新（LibreOffice 命令行不更新域）。请在 Word 中按 Ctrl+A → F9 后重新导出。
```

> ⚠️ **忘记 `PYTHONPATH` 会直接 `ModuleNotFoundError: No module named 'docx'`**，这是新人第一个坑。

---

## 3. 代码地图

| 文件 | 职责 | 文档 |
|------|------|------|
| `tf/core/enums.py` | `BlockKind` / `BlockType` / `Align` / `VertAlign` / `BlockSource` 枚举 | `docs/06 §1` |
| `tf/core/models.py` | `Block` / `RunInfo` / `AssetInfo` / `Style` / `TemplateSpec` / `RenderResult` 等数据契约 | `docs/06 §1` |
| `tf/core/idgen.py` | 确定性块 ID（`b0001`…，跨重复解析稳定、可反解序号） | `docs/02 §3` |
| `tf/ingestion/engine.py` | `ingest(docx) -> list[Block]`：段落/表格/图片/公式/分节符 → 块；只读，不改 DOM | `docs/02` |
| `tf/ingestion/assets.py` | 资产探测：`detect_paragraph_kind`、脚注引用、分节符、`qn` / `W` 命名空间助手 | `docs/02 §4` |
| `tf/parser/regex_parser.py` | `classify(blocks, rules)`：两层解析里的正则层（封面/摘要/标题/图题表题/参考文献/公式…） | `docs/03` |
| `tf/parser/rules_min.json` | 内置最小规则表（`default` + `styles`），通过 `schemas/rules.schema.json` 校验 | `docs/01 §2` |
| `tf/template/rules.py` | 规则表加载 / 校验 / 字段级继承级联（`resolve_style`） | `docs/01 §2.2` |
| `tf/template/master.py` | 母版 docx 载入、样式快照、母版缺失时的空白文档兜底 | `docs/01` |
| `tf/template/engine.py` | `TemplateSpec` 组装（母版 + 规则 + 快照） | `docs/01 §2.7` |
| `tf/render/engine.py` | `render(blocks, spec, out)` 灌母版 + 套样式 + 保护块搬运；`finalize(docx, spec)` 收尾 | `docs/05` |
| `tf/render/ooxml.py` | 直接操作 `w:xml` 的底层助手（字体分离、大纲级别、段落属性） | `docs/05 §3` |
| `tf/cli.py` | `ingest` / `parse` / `run` / `render` 四个子命令；告警走 stderr，成功行走 stdout | `docs/07 §3` |
| `tf/review/__init__.py` | **空占位**：人机校对台尚未实现（见 §6） | `docs/04` |

测试与文档的对应关系在 `docs/07 §2 测试清单`，每个测试文件头部也写了覆盖清单。

---

## 4. 收尾（finalize）现状 —— 本轮重点

`finalize(docx_path, spec) -> FinalizeResult`，逐级降级，**任何情况都不抛异常**：

| 优先级 | 引擎 | 探测方式 | 目录/域 | PDF |
|-------|------|---------|---------|-----|
| 1 | `word` | `sys.platform` 为 win32 且 `import win32com.client` 成功 | ✅ 刷新 | ✅ 导出 |
| 2 | `libreoffice` | PATH 上 `soffice` / `libreoffice` | ❌ 不刷新 | ✅ 导出 |
| 3 | `none` | 都没有 | ❌ | ❌ |

- Word 分支：`DispatchEx` 起独立进程 → `Fields.Update()` → 遍历 `TablesOfContents(i).Update()`
  → `Repaginate()` → `ExportAsFixedFormat(ExportFormat=17 /*PDF*/, OptimizeFor=0)`；
  `finally` 里 `Close(Saving=False)` + `Quit()`，异常不残留 `WINWORD` 进程。
- LibreOffice 分支：`--headless --norestore --invisible --convert-to pdf --outdir <docx 目录>`，
  并加 `-env:UserInstallation=file://<系统临时目录>` **隔离用户配置**——不隔离时与常驻
  `soffice` 抢 profile 会间歇性静默失败（这正是之前全量测试偶发红的原因）。
  临时 profile 用完即删，不留垃圾在用户产物目录。
- 降级语义：`degraded=True` + `warnings`（面向用户的下一步动作），退出码仍是 `0`；
  PDF 路径固定为 `docx_path.with_suffix(".pdf")`。
- `toc_updated` / `fields_updated` / `pdf_path` / `docx_path` 供 CLI 与质检报告引用。

> **Windows 分支尚未在真实 Windows+Word 上验证**（本机是 Linux），逻辑按 `docs/05 §6.2`
> 编写、用 stub 覆盖了调用顺序，但 `DispatchEx` 实机行为需在 Windows 上回归一次。

---

## 5. 本轮（M3）改动摘要

1. **`tf/render/engine.py`**：把 `finalize` 从「恒返回 `engine="none"` 的桩」改成真实三级降级实现
   （`_detect_finalizer` / `_finalize_with_word` / `_finalize_with_libreoffice` / `_find_soffice`）。
2. **LibreOffice profile 隔离**：修复全量测试下 `test_finalize_uses_libreoffice_to_export_pdf` 的间歇性失败。
3. **`tf/cli.py`**：收尾后若拿到 PDF 就打印 `[pdf] 已导出 …（引擎：…）`；告警继续走 stderr。
4. **`tests/test_render.py`**：新增/改写收尾测试（`none` 用 monkeypatch 钉死降级，`libreoffice` 真跑导出并断言 `toc_updated is False` + 告警文案）。
5. **`tests/test_cli.py`**：加 autouse fixture 把引擎探测钉成 `none`，让 CLI 测试不依赖宿主是否装 Office、也不再每次触发一次无头转换（22 个测试从 ~10s 降到 ~1s）。
6. **文档**：`docs/05 §6` 的引擎矩阵/降级说明、`docs/00` 的降级注记，均按实现更新；新增本文件。

---

## 6. 已知限制 / 未完成（按优先级）

| # | 事项 | 说明 | 建议入口 |
|---|------|------|---------|
| 0 | **2 个已定位的解析缺陷**（测试标记 `xfail`） | ① `regex_parser` 的 `Context.section_index` 从未从 block 同步，守卫 `c.section_index == 0` 恒为真 → 第 1 节及以后的「封面标题」仍被误判为 `COVER_TITLE`；② `ingestion/engine._extract_runs` 用「`w:b` 元素是否存在」判断粗体，忽略 `w:val="0"`，导致显式 `bold=False`（python-docx 写入）被读成 True，进而误触发粗体标题启发式 | `tests/test_parser_regex.py:252`、`:401`；实现分别在 `tf/parser/regex_parser.py`（`Context` 构造处逐块同步 `section_index`）、`tf/ingestion/engine.py::_extract_runs`（改读 `w:val`，`0/false/off` 视为非粗体） |
| 1 | **校对台未实现** | `tf/review/` 只有一个空 `__init__.py`；`docs/04` 的树状大纲 + 人工纠偏尚无代码 | `docs/04` |
| 2 | **无顶层编排 API** | `docs/06 §3` 的 `tf/pipeline.py`（`PipelineConfig` / `ThesisFormatter.run`）不存在，CLI 里手写了 `_run_pipeline` | `tf/cli.py::_run_pipeline` |
| 3 | **LLM 分类层缺失** | 只有正则层；`use_llm` 配置项无处生效，模糊块一律走兜底规则 | `tf/parser/regex_parser.py` |
| 4 | **Windows 收尾未实机验证** | 见 §4 末尾 | 需 Windows + Word 环境 |
| 5 | **不支持 `.doc` 输入** | 只吃 `.docx`；`tests/fixtures/source_template.doc` 是历史样例，无测试使用 | `docs/08 Q1` |
| 6 | **`examples/` 为空** | 目录占位，可放「输入样例 → 输出样例 → report.json」 | — |
| 7 | **未打包** | 无 `pyproject.toml` / `setup.cfg`，只能 `python3 -m tf.cli` 或加 `sys.path` | — |
| 8 | **质检报告只返回 dict** | `RenderResult.qa_report` 是结构化 dict，`report.json` / `.md` 落盘尚未实现（`docs/05 §7` 有样例格式） | `docs/05 §7` |

---

## 7. 测试地图

| 文件 | 用例数 | 覆盖 |
|------|-------|------|
| `tests/test_core.py` | 19 | 数据模型、枚举、块 ID 生成与反解 |
| `tests/test_ingestion.py` | 45 | 段落/表格/图片/公式/分节符拆块、`_protected`、`rel_id`、`section_index`、run 级字号 |
| `tests/test_parser_regex.py` | 93 | 正则分类：封面、摘要、各级标题、图表题、参考文献、公式、边界与回归样例 |
| `tests/test_template.py` | 30 | 规则加载/校验（jsonschema 缺失时降级）、字段级继承级联、母版样式快照 |
| `tests/test_render.py` | 21 | 灌母版、样式套用（缩进/大纲级别/悬挂缩进）、保护块、**finalize 三引擎分支** |
| `tests/test_cli.py` | 22 | 四个子命令端到端、stdout/stderr 分离约定、退出码 2、`--out` 落盘 |

约定（改测试前先读）：
- 测试只写 `tmp_path`，绝不写仓库目录；仓库自带样例是 `tests/fixtures/source_template.docx`（接入+分类后恒为 **26 块**）。
- 不引入自定义 pytest marker；环境相关用 `pytest.mark.skipif`，环境差异用 `monkeypatch` 钉死。
- `2 xfailed` 是**已知产品缺陷**的登记（`@pytest.mark.xfail` 带 reason，见 `tests/test_parser_regex.py:252` 与 `:401`）：
  ① 封面标题的 `section_index` 守卫失效；② 显式非粗体被误读为粗体。
  修好这两个 bug 后它们会变成 `XPASS`，届时把 `xfail` 标记删掉即可（建议加 `strict=True` 防止悄悄过期）。
- 改动 `finalize` 时注意：**不要把 `libreoffice` 分支改成 `toc_updated=True`**——LO 命令行确实不刷新域，这条正是 `docs/05 §6.3` 的契约。

---

## 8. 下一阶段建议

1. **校对台 MVP**（`tf/review`）：先做 `build_outline`（复用 `BlockType` 树）+ `apply_fixes`（按块 ID 改类型/对齐），不做前端，先出 JSON + 文本大纲。
2. **`tf/pipeline.py`**：把 `cli._run_pipeline` 抽成 `ThesisFormatter.run()`，CLI 只负责参数解析；顺手把 `report.json` / `report.md` 落盘补上。
3. **LLM 层**：按 `docs/03` 的低置信度阈值筛块，接口留 `use_llm` 开关，缺 key 时自动回退正则。
4. **Windows 回归**：在 Windows 上跑 `pytest -k finalize`，确认 `ExportAsFixedFormat` 与 `Quit()` 无残留进程，再补一条 `docs/05 §6.2` 的实测记录。
5. **打包**：`pyproject.toml` + `tf` 入口点（`tf = tf.cli:main`），可选依赖分组 `[office]`（pywin32）。

---

## 9. 交接 checklist

- [x] 仓库已初始化并推送 `git@github.com:snsnnd/word.git`（`main` 分支）
- [x] `python3 -m pytest tests/ -q` → 全绿（228 passed / 2 xfailed）
- [x] CLI 端到端产出 `.docx` + `.pdf` 实测通过
- [x] `schemas/rules.schema.json` 能校验 `tf/parser/rules_min.json`（jsonschema 4.26.0 实测）
- [x] `docs/00` → `docs/09` 文档与实现一致（`05 §6` 引擎矩阵已按实现更新）
- [ ] Windows + Word 实机回归（本机无环境，留给接手的同学）
- [ ] `.doc` 输入 / LLM 层 / 校对台 按 §8 排期

## 10. 任务状态

本轮任务到此**停止**：主链路与收尾已交付并全绿，剩余事项（§6、§8）不在本轮范围内。
后续从那一条开始，直接按 §8 的优先级推进即可。
