"""命令行入口（见 docs/06 §4、docs/07 §2）。

用法：
  tf ingest  <docx> [--out FILE]              解析为块，输出 JSON
  tf parse   <docx> [--rules FILE] [--out FILE]  接入+分类，输出带语义标签的 JSON
  tf run     <docx> [--template DIR] [--out FILE]  完整流水线（接入+分类+渲染 → docx）
  tf render  <docx> [--template DIR] [--out FILE]  渲染 → docx
  tf report  <dir>                            汇总报告（M3）
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from tf.ingestion.engine import ingest
from tf.parser.regex_parser import classify
from tf.render import finalize, render
from tf.template.rules import default_rules_path, load_rules
from tf.template.master import extract_master_styles
from tf.core.models import TemplateSpec


def _cmd_ingest(args) -> int:
    blocks = ingest(args.docx)
    _dump_blocks(blocks, args.out)
    return 0


def _cmd_parse(args) -> int:
    blocks = ingest(args.docx)
    classify(blocks)
    _dump_blocks(blocks, args.out)
    return 0


def _resolve_master(template):
    """把 --template（目录或母版 docx）解析为母版 docx 路径；无则 None。"""
    if not template:
        return None
    p = Path(template)
    if p.is_dir():
        for name in ("master.docx", "template.docx"):
            cand = p / name
            if cand.exists():
                return str(cand)
        docs = sorted(p.glob("*.docx"))
        if docs:
            return str(docs[0])
        return None
    return str(p) if p.exists() else None


def _build_spec(args):
    """构建 TemplateSpec：rules.json（缺省用内置）+ 母版 docx + 母版样式快照。"""
    master = _resolve_master(args.template)
    rules = load_rules(args.rules) if getattr(args, "rules", None) else load_rules(default_rules_path())
    snapshot = extract_master_styles(master) if master else {}
    return TemplateSpec(name="cli", version="0.1.0", master_path=master,
                        rules=rules, style_snapshot=snapshot)


def _default_out(docx_path) -> str:
    """未指定 --out 时，产物落在源文件同目录的 <stem>_formatted.docx。"""
    p = Path(docx_path)
    return str(p.with_name(f"{p.stem}_formatted.docx"))


def _run_pipeline(args) -> int:
    """接入 → 分类 → 灌母版渲染 → finalize → docx（run/render 共用）。"""
    blocks = ingest(args.docx)
    classify(blocks)
    spec = _build_spec(args)
    out = args.out or _default_out(args.docx)
    if not spec.master_path:
        print("  [warn] 未指定 --template（或路径不存在），将用空白文档渲染，"
              "样式不会套用母版", file=sys.stderr)
    result = render(blocks, spec, out)
    stats = result.qa_report.get("stats", {})
    tag = "run" if args.cmd == "run" else "render"
    print(f"[{tag}] 已生成 {result.docx_path}"
          f"（{stats.get('rendered', 0)}/{stats.get('blocks', 0)} 块）")
    for w in stats.get("warnings", []):
        print(f"  [warn] {w}", file=sys.stderr)
    fin = finalize(result.docx_path, spec)
    for w in fin.warnings:
        print(f"  [finalize:{fin.engine}] {w}", file=sys.stderr)
    if fin.pdf_path:
        print(f"[pdf] 已导出 {fin.pdf_path}（引擎：{fin.engine}）")
    return 0


def _cmd_run(args) -> int:
    return _run_pipeline(args)


def _cmd_render(args) -> int:
    return _run_pipeline(args)


def _dump_blocks(blocks, out_path) -> None:
    data = [b.to_dict() for b in blocks]
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if out_path:
        Path(out_path).write_text(text, encoding="utf-8")
        print(f"[ok] 写出 {len(blocks)} 个块 → {out_path}")
    else:
        print(text)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="tf", description="thesis-formatter 论文/报告自动排版工具")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("ingest", help="解析为块")
    s.add_argument("docx")
    s.add_argument("--out", default=None)
    s.set_defaults(func=_cmd_ingest)

    s = sub.add_parser("parse", help="接入+分类")
    s.add_argument("docx")
    s.add_argument("--rules", default=None)
    s.add_argument("--out", default=None)
    s.set_defaults(func=_cmd_parse)

    s = sub.add_parser("run", help="完整流水线")
    s.add_argument("docx")
    s.add_argument("--template", default=None, help="模板目录或母版 docx（缺省用内置规则 + 空白文档）")
    s.add_argument("--rules", default=None)
    s.add_argument("--out", default=None)
    s.set_defaults(func=_cmd_run)

    s = sub.add_parser("render", help="渲染")
    s.add_argument("docx")
    s.add_argument("--template", default=None, help="模板目录或母版 docx（缺省不灌母版）")
    s.add_argument("--out", default=None)
    s.set_defaults(func=_cmd_render)

    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())