"""`tf/cli.py` 端到端测试（docs/07 §2 测试清单 · test_cli）。

直接调 `main(argv)`（返回 int，不触发 SystemExit），用 capsys 断言：
- 告警/进度写 stderr，成功行写 stdout；
- 缺参数时 argparse 抛 SystemExit(2)，usage 写 stderr。

覆盖：
- `tf render <docx> --out <file>`（无 --template）→ 0，落 docx，stderr 提示「未指定 --template」
  + finalize 提示「[finalize:none] … 未检测到 Word」
- `tf run <docx> --template <docx|目录> --out <file>` → 0，stdout 打印 26/26 块，
  产物含封面文字
- `tf ingest` / `tf parse` → 块 JSON（stdout 或 --out 文件）
- 缺参 / 未知子命令 → SystemExit 2 + usage

约定：
- 产物一律落在 pytest 的 tmp_path，绝不写仓库目录（除非显式验证 _default_out 的
  命名规则，那时先把源文件复制到 tmp_path）。
- 本机有 Word/WPS 才会出现 finalize 成功；CI/无 Office 环境恒为 engine="none"。
"""
from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

import docx
import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tf.cli import main  # noqa: E402

#: 仓库自带样例（接入 + 分类后恒为 26 个块）
FIXTURE = ROOT / "tests" / "fixtures" / "source_template.docx"

#: 样例文档的块数（ingest/classify 的既成事实，供断言引用）
EXPECTED_BLOCKS = 26

#: 母版标记文字：用于验证 --template 真的生效
MASTER_MARKER = "MASTER-MARKER-7Q"


# --------------------------------------------------------------------------- #
# 辅助
# --------------------------------------------------------------------------- #
@pytest.fixture(autouse=True)
def _no_office(monkeypatch):
    """CLI 测试不依赖本机是否装了 Office。

    finalize 的引擎探测（word > libreoffice > none）会随宿主环境变化：本机装了
    LibreOffice 时 `render` 会真的跑一次无头转换，既慢又让 `[finalize:…]` 的断言
    不稳定。这里统一钉死为 `none`，LibreOffice 真导出的路径由 test_render.py
    的 `test_finalize_uses_libreoffice_to_export_pdf` 单独覆盖。
    """
    from tf.render import engine as engine_mod

    monkeypatch.setattr(engine_mod, "_detect_finalizer", lambda: "none")
def _all_text(document) -> str:
    """把 docx 的段落/表格文本拼成一个字符串（去空格，便于跨 Word 版本比对）。"""
    chunks = [p.text for p in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                chunks.append(cell.text)
    return "".join(c.replace(" ", "").replace("\u3000", "") for c in chunks)


def _docx_xml(path: Path) -> str:
    """读取 docx 包内全部 XML 的文本（含页眉页脚）。

    渲染层会 `_clear_body` 清空母版正文，只保留样式 / 页眉页脚 / sectPr，
    所以「母版是否生效」的标记必须写在页眉里，并在这里跨 XML 查找。
    """
    with zipfile.ZipFile(path) as archive:
        blobs = [
            archive.read(name)
            for name in archive.namelist()
            if name.startswith("word/") and name.endswith(".xml")
        ]
    return b"\n".join(blobs).decode("utf-8", errors="ignore")


def _make_master(path: Path, marker: str = MASTER_MARKER) -> Path:
    """造一个带可识别标记的母版 docx（标记写在页眉，正文会被渲染层清空）。"""
    document = docx.Document()
    document.add_heading("Master Heading", level=1)
    document.add_paragraph(marker)
    document.sections[0].header.paragraphs[0].text = marker
    document.save(str(path))
    return path


def _blocks_from_stdout(captured) -> list[dict]:
    """stdout 必须是纯 JSON 数组（告警走 stderr，不污染 stdout）。"""
    payload = json.loads(captured.out)
    assert isinstance(payload, list)
    return payload


# --------------------------------------------------------------------------- #
# render：无 --template
# --------------------------------------------------------------------------- #
def test_render_without_template_returns_zero_and_writes_docx(tmp_path, capsys):
    """render（无 --template）：退出码 0、落 docx、可被 python-docx 打开。"""
    out = tmp_path / "render.docx"

    rc = main(["render", str(FIXTURE), "--out", str(out)])
    captured = capsys.readouterr()

    assert rc == 0
    assert out.is_file() and out.stat().st_size > 0
    document = docx.Document(str(out))  # 不抛异常即说明是合法 docx
    assert document.paragraphs


def test_render_without_template_warns_on_stderr(tmp_path, capsys):
    """缺 --template：提示写 stderr，不写 stdout。"""
    out = tmp_path / "render.docx"

    main(["render", str(FIXTURE), "--out", str(out)])
    captured = capsys.readouterr()

    assert "未指定 --template" in captured.err
    assert "未指定 --template" not in captured.out


def test_render_prints_finalize_notice_on_stderr(tmp_path, capsys):
    """render 末尾必须打印 finalize 行（无 Office 环境即 `[finalize:none] … 未检测到 Word`）。"""
    out = tmp_path / "render.docx"

    main(["render", str(FIXTURE), "--out", str(out)])
    captured = capsys.readouterr()

    assert "[finalize:none]" in captured.err
    assert "未检测到 Word" in captured.err


def test_render_prints_success_line_on_stdout(tmp_path, capsys):
    """成功行（`[render] 已生成 …（26/26 块）`）走 stdout。"""
    out = tmp_path / "render.docx"

    main(["render", str(FIXTURE), "--out", str(out)])
    captured = capsys.readouterr()

    assert "[render] 已生成" in captured.out
    assert str(out) in captured.out
    assert f"{EXPECTED_BLOCKS}/{EXPECTED_BLOCKS} 块" in captured.out


def test_render_with_nonexistent_template_warns(tmp_path, capsys):
    """--template 指向不存在的路径：静默退化为空白文档，并打印同一条告警。"""
    out = tmp_path / "render.docx"

    rc = main([
        "render", str(FIXTURE),
        "--template", str(tmp_path / "does-not-exist.docx"),
        "--out", str(out),
    ])
    captured = capsys.readouterr()

    assert rc == 0
    assert "未指定 --template" in captured.err
    assert out.is_file()


def test_render_default_out_path_is_stem_formatted(tmp_path, capsys):
    """未给 --out 时产物 = 源文件同目录的 <stem>_formatted.docx（_default_out）。"""
    source = tmp_path / "sample.docx"
    source.write_bytes(FIXTURE.read_bytes())

    rc = main(["render", str(source)])
    capsys.readouterr()

    assert rc == 0
    produced = tmp_path / "sample_formatted.docx"
    assert produced.is_file()
    assert produced.read_bytes() != source.read_bytes()


# --------------------------------------------------------------------------- #
# run：--template 生效
# --------------------------------------------------------------------------- #
def test_run_with_template_docx_renders_cover(tmp_path, capsys):
    """run --template <母版 docx>：26/26 块，产物含封面文字，且不再有「未指定 --template」告警。"""
    template = tmp_path / "master.docx"
    _make_master(template)
    out = tmp_path / "run.docx"

    rc = main(["run", str(FIXTURE), "--template", str(template), "--out", str(out)])
    captured = capsys.readouterr()

    assert rc == 0
    assert f"{EXPECTED_BLOCKS}/{EXPECTED_BLOCKS} 块" in captured.out
    assert "[run] 已生成" in captured.out
    assert "未指定 --template" not in captured.err
    assert out.is_file()

    text = _all_text(docx.Document(str(out)))
    assert "实验报告" in text           # 封面主标题
    assert "数字图像处理课程实验报告" in text
    assert "MASTER-MARKER-7Q" in _docx_xml(out)   # 母版页眉确实被带入


def test_run_with_template_directory_resolves_master_docx(tmp_path, capsys):
    """--template 传目录时，按 master.docx → template.docx → *.docx 顺序解析。"""
    template_dir = tmp_path / "tpl"
    template_dir.mkdir()
    _make_master(template_dir / "master.docx")
    out = tmp_path / "run_dir.docx"

    rc = main(["run", str(FIXTURE), "--template", str(template_dir), "--out", str(out)])
    captured = capsys.readouterr()

    assert rc == 0
    assert "[run] 已生成" in captured.out
    assert "未指定 --template" not in captured.err
    assert "MASTER-MARKER-7Q" in _docx_xml(out)


def test_run_template_directory_prefers_master_over_other_docx(tmp_path):
    """目录里同时有 master.docx 与别的 docx 时，必须优先取 master.docx。"""
    template_dir = tmp_path / "tpl"
    template_dir.mkdir()
    _make_master(template_dir / "master.docx", marker="FROM-MASTER-DOCX")
    _make_master(template_dir / "aaa_first.docx", marker="FROM-AAA-DOCX")
    out = tmp_path / "run_pref.docx"

    rc = main(["run", str(FIXTURE), "--template", str(template_dir), "--out", str(out)])

    assert rc == 0
    xml = _docx_xml(out)
    assert "FROM-MASTER-DOCX" in xml
    assert "FROM-AAA-DOCX" not in xml


def test_run_template_directory_without_docx_falls_back(tmp_path, capsys):
    """空模板目录 → 解析失败 → 退回空白文档 + 「未指定 --template」告警。"""
    empty_dir = tmp_path / "empty_tpl"
    empty_dir.mkdir()
    out = tmp_path / "run_empty.docx"

    rc = main(["run", str(FIXTURE), "--template", str(empty_dir), "--out", str(out)])
    captured = capsys.readouterr()

    assert rc == 0
    assert "未指定 --template" in captured.err
    assert out.is_file()


# --------------------------------------------------------------------------- #
# ingest / parse
# --------------------------------------------------------------------------- #
def test_ingest_dumps_blocks_to_stdout(capsys):
    """ingest（无 --out）：stdout 是块 JSON 数组，退出码 0。"""
    rc = main(["ingest", str(FIXTURE)])
    captured = capsys.readouterr()

    assert rc == 0
    blocks = _blocks_from_stdout(captured)
    assert len(blocks) == EXPECTED_BLOCKS
    assert blocks[0]["block_id"]
    assert blocks[0]["index"] == 0


def test_ingest_writes_out_file(tmp_path, capsys):
    """ingest --out：写 JSON 文件，stdout 打印 `[ok] 写出 26 个块 → …`。"""
    out = tmp_path / "blocks.json"

    rc = main(["ingest", str(FIXTURE), "--out", str(out)])
    captured = capsys.readouterr()

    assert rc == 0
    assert f"[ok] 写出 {EXPECTED_BLOCKS} 个块" in captured.out
    assert str(out) in captured.out
    data = json.loads(out.read_text(encoding="utf-8"))
    assert len(data) == EXPECTED_BLOCKS


def test_parse_assigns_semantic_tags(capsys):
    """parse：每个块都有非空 semantic_tag（分类已跑）。"""
    rc = main(["parse", str(FIXTURE)])
    captured = capsys.readouterr()

    assert rc == 0
    blocks = _blocks_from_stdout(captured)
    assert len(blocks) == EXPECTED_BLOCKS
    for block in blocks:
        assert block.get("semantic_tag")
        assert block["confidence"] >= 0.0


def test_parse_writes_out_file(tmp_path, capsys):
    """parse --out：写带语义标签的 JSON 文件。"""
    out = tmp_path / "parsed.json"

    rc = main(["parse", str(FIXTURE), "--out", str(out)])
    captured = capsys.readouterr()

    assert rc == 0
    assert f"[ok] 写出 {EXPECTED_BLOCKS} 个块" in captured.out
    data = json.loads(out.read_text(encoding="utf-8"))
    assert all("semantic_tag" in item for item in data)


# --------------------------------------------------------------------------- #
# 参数错误
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    "argv",
    [
        [],
        ["render"],
        ["run"],
        ["ingest"],
        ["parse"],
        ["render", str(FIXTURE), "--unknown-flag"],
        ["no-such-subcommand", str(FIXTURE)],
    ],
    ids=[
        "no-args",
        "render-missing-docx",
        "run-missing-docx",
        "ingest-missing-docx",
        "parse-missing-docx",
        "unknown-flag",
        "unknown-subcommand",
    ],
)
def test_bad_args_exit_with_usage(argv, capsys):
    """缺参/未知子命令 → SystemExit(2)，usage 写 stderr。"""
    with pytest.raises(SystemExit) as excinfo:
        main(argv)
    captured = capsys.readouterr()

    assert excinfo.value.code == 2
    assert "usage:" in captured.err


def test_usage_mentions_all_subcommands(capsys):
    """usage 行必须列出全部子命令（ingest/parse/run/render）。"""
    with pytest.raises(SystemExit):
        main([])

    err = capsys.readouterr().err
    for name in ("ingest", "parse", "run", "render"):
        assert name in err