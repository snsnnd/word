# ⑤ 渲染与后处理引擎（Render Engine）

> 把带标签的块，按顺序写进母版 `.docx`，套用母版预设样式，执行字体分离、清多余空行、三线表转换、图片统一缩放；最后做 **Word 原生收尾**（刷新目录域、导出 PDF）。

---

## 1. 渲染总流程

```
render(blocks, template: TemplateSpec, out_path)
  │
  ├─ 1. 载入母版为「工作副本」（渲染在工作副本上进行，源母版只读）
  ├─ 2. 清空母版正文占位内容（保留样式、页眉页脚、分节、TOC 域）
  ├─ 3. 逐块渲染：
  │     ├─ 保护块（image/table/formula/footnote）：定位并保留，仅按规则改外框
  │     ├─ 普通段落：创建段落 → 套样式 → 字体分离 → 应用段距/缩进
  │     └─ 标题段落：套 Heading N
  ├─ 4. 后处理：
  │     ├─ 清多余空行
  │     ├─ 合并被拆散的 run
  │     ├─ 三线表转换
  │     ├─ 图片统一缩放 + 居中
  │     └─ 中英文间空格/标点（可选）
  ├─ 5. 分节与页码：按 page.sections 设置页码格式、起始页
  ├─ 6. 保存工作副本为 out_path（.docx）
  └─ 7. Word 原生收尾（finalization，可选）：更新域 + 导出 PDF
```

---

## 2. 内容灌入（写入母版）

### 2.1 为什么是「灌入」而非「新建」
母版已含样式、页眉页脚、分节、目录域。直接新建文档会丢这些。**灌入 = 在母版副本的正文区按块顺序重建内容**，同时保留 `sectPr` 与 `TOC` 域。

### 2.2 样式应用
- 每个块 → `paragraph.style = spec.master_styles[tag]`（或规则表 `resolve(tag)`）。
- 若母版缺对应样式 → 回退 `Normal` 并按规则表「代码式」补设字体/字号/段距（兜底，同时告警）。

### 2.3 字体分离（中英文分别刷字体）
核心步骤。对每个 run：
- 判断字符是中文还是西文（Unicode 区间 + `w:rFonts` 的 `w:eastAsia` / `w:ascii` / `w:hAnsi`）。
- 中文走 `font_cn`，西文走 `font_en`。
- **上下标 run 不改 vertAlign**；**公式 run（OMML）不改字体**。
- 若一个 run 混合中英（如 `中文abc`），按 `merge_split_runs` 拆成两段分别设字体，再合回（或保留拆分）。

### 2.4 保护块处理
- **图片**：只改 `wp:extent`（按 `max_width_cm` 等比缩放）、`wp:anchor` → `inline`（若规则要求）、居中。
- **表格**：转三线表（见 §4）、设字号、表头加粗、跨页重复表头。
- **公式**：保留 OMML 树；仅调整公式段行距/编号右对齐；MathType 对象**只告警不改**。

---

## 3. 后处理（Post-process）

| 处理 | 说明 | 风险 |
|------|------|------|
| 清多余空行 | 删除连续空段落（保留分节/分页处） | 误删需人工校对台防呆 |
| 合并被拆散 run | 同一段内连续同格式 run 合并，修复字体错乱 | 低 |
| 图片统一缩放 | 超 `max_width_cm` 等比缩小，居中 | 低 |
| 三线表转换 | 见 §4 | 中（合并单元格需小心） |
| 中英文空格 | 中英文之间加空格（可选） | 低 |
| 标点转换 | 保守默认关闭 | 中 |

所有后处理均可在规则表 `postprocess` 中开关。

---

## 4. 三线表转换细节

三线表 = 顶线、表头下线、底线，无竖线：

```
━━━━━━━━━━━━━━━━━━━━━━━━  ← 顶线（1.5pt）
   表头1   表头2   表头3
────────────────────────  ← 表头下线（0.75pt）
   数据    数据    数据
━━━━━━━━━━━━━━━━━━━━━━━━  ← 底线（1.5pt）
```

实现要点（OOXML）：
- 清除所有 `w:tcBorders` 的左右/内部竖线。
- 顶线：表格 `w:tblBorders` 的 `top` = single, sz=12（1.5pt）。
- 表头下线：首行 `w:tcBorders` 的 `bottom` = single, sz=6（0.75pt）。
- 底线：表格 `w:tblBorders` 的 `bottom` = single, sz=12。
- 保留合并单元格（`gridSpan`/`vMerge`）。
- 若 `table.style = keep_as_is` 则跳过。

---

## 5. 分节与页码

按 `page.sections` 配置：
- 封面：无页码。
- 摘要/目录：罗马数字（`w:pgNumType w:fmt="upperRoman"`）。
- 正文：阿拉伯数字，起始 1（`w:pgNumType w:start="1" w:fmt="decimal"`）。
- 参考文献：续前节（`w:fmt="decimal"`，不重置 start）。

分节符对应 Ingestion 的 `section_break` 块，渲染时映射到 `w:sectPr`。

---

## 6. Word 原生收尾（Finalization）

> Python 生态库（如 `python-docx`）**无法自动更新目录页码和交叉引用**。这必须由 Word 自己完成。

### 6.1 引擎选择（`finalization.engine`）
| 引擎 | 条件 | 说明 |
|------|------|------|
| `win32com` | Windows + 装有 Word | `pywin32` 静默调 Word，刷新域 + 导出 PDF |
| `vba` | Windows + Word | 执行内嵌 VBA 宏刷新域 |
| `auto` | 自动探测（实现按此执行） | `word` → `libreoffice` → `none`，逐级降级 |
| `none` | Linux 且无 LibreOffice | 仅产出 docx，提示用户手动更新域 |

实现落地（`tf/render/engine.py::_detect_finalizer`，顺序探测，不抛异常）：

| 探测顺序 | 引擎 | 探测方式 | 目录/域 | PDF |
|---------|------|---------|---------|-----|
| 1 | `word` | `sys.platform` 为 win32 且 `import win32com.client` 成功 | ✅ 刷新 | ✅ 导出 |
| 2 | `libreoffice` | PATH 上有 `soffice` / `libreoffice` | ❌ 不刷新（命令行不更新域） | ✅ 导出 |
| 3 | `none` | 以上都没有 | ❌ | ❌ |

`FinalizeResult.degraded` 在 `libreoffice` / `none` 下为 `True`，`warnings` 说明具体原因；
`engine` / `toc_updated` / `fields_updated` / `pdf_path` 供 CLI 与质检报告直接引用。

### 6.2 win32com 收尾流程（Windows）
```python
word = win32.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(str(out_path))
# 刷新目录
for toc in doc.TablesOfContents:
    toc.Update()
# 刷新所有域（交叉引用、页码）
doc.Fields.Update()
# 导出 PDF
doc.SaveAs(str(pdf_path), FileFormat=17)
doc.Close(False)
word.Quit()
```

实现为 `DispatchEx`（独立 Word 进程，避免附着到用户已打开的文档），
导出走 `ExportAsFixedFormat(ExportFormat=17, OptimizeFor=0)`；
`try/except` 内捕获失败 → `degraded=True`，`finally` 关闭文档并 `Quit()`，
异常路径不会留下常驻 `WINWORD` 进程。

### 6.3 降级（Linux/无 Word）
- `libreoffice`（有 `soffice`/`libreoffice` 时）：headless 转 PDF 作预览，**不刷新目录域**
  （命令行不更新 TOC/交叉引用），`toc_updated=False` 并给出提示：
  `目录/交叉引用未刷新（LibreOffice 命令行不更新域）。请在 Word 中按 Ctrl+A → F9 后重新导出。`
  - 使用独立 `-env:UserInstallation`（系统临时目录）避免与常驻 soffice 抢用户配置；
    结束后删除该临时配置目录，不在用户产物目录留下垃圾。
- `none`（都不可用）：
  `未检测到 Word/WPS 或 LibreOffice。请在 Word 中按 Ctrl+A → F9 更新域后重新导出 PDF。`
- 两种情况都产出 `.docx`，退出码仍为 `0`（降级不是失败），告警进 `warnings` → CLI stderr。

---

## 7. 质检清单（日志输出）

渲染结束打印结构化日志（同时写入 `report.json` / `.md`）：

```
========== TF 质检报告 ==========
文档：毕业论文_v3.docx
母版：thesis_master.docx
规则：本科规范 v1

[结构]
  - 标题：一级 6 / 二级 24 / 三级 41
  - 正文段落：312
  - 图：14 张（已统一缩放至 ≤12cm）
  - 表格：5 个（已转三线表）
  - 公式：23 个 OMML；疑似 MathType 3 处 —— ⚠ 请人工核对行距
  - 参考文献：38 条

[告警]
  - 母版缺样式 "Caption"，已回退规则表兜底
  - 块 b0032 置信度低(0.62)，已按 body 渲染
  - 疑似公式块 b0210/b0211/b0212 为 MathType 对象，未改动

[收尾]
  - 引擎：win32com —— 已刷新目录域 + 导出 PDF
  - 输出：out/毕业论文_v3.docx , out/毕业论文_v3.pdf
==================================
```

---

## 8. 接口契约（详见 §06）

```python
class RenderEngine(Protocol):
    def render(self, blocks: list[Block], spec: TemplateSpec, out_path: Path) -> RenderResult: ...
    def finalize(self, docx_path: Path, spec: TemplateSpec) -> FinalizeResult: ...
```

---

## 9. 竞品参考

- 三线表/OMML/目录引用处理参考 [docx-formatter-skill](https://github.com/CP-here/docx-formatter-skill)。
- 纯 Python 更新域的局限与 LibreOffice 替代方案：见 CSDN 相关讨论（Linux 下绕开 win32com）。
- 公式 OMML 原生处理参考 [dsh-formula-docx](https://github.com/Red-Rayux/dsh-formula-docx)。