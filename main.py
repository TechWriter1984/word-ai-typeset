"""一键启动入口：步骤2 → 步骤3 → 步骤4(占位)。

用法：
    python main.py <前处理后的 docx 路径>
    python main.py <path> --skip-typeset          # 跳过 LLM 分类，直接合并
    python main.py <path> --output D:\\out.docx    # 指定最终输出路径
    python main.py <path> --config D:\\cfg.yaml    # 指定配置文件
"""
from __future__ import annotations

import argparse
import os
import sys

from config import load_config
from typeset import typeset_doc
from merge import merge_template


def parse_args():
    p = argparse.ArgumentParser(description="Word AI 自动排版 + 模板合并")
    p.add_argument("input", help="用户前处理后的 docx 路径（步骤2输入）")
    p.add_argument("--config", default=None, help="config.yaml 路径（默认项目根）")
    p.add_argument("--output", default=None,
                   help="最终输出 docx 路径（默认 output/<base>_final.docx）")
    p.add_argument("--skip-typeset", action="store_true",
                   help="跳过步骤2，直接对输入文件做合并（输入需已排版）")
    return p.parse_args()


def main():
    args = parse_args()
    cfg = load_config(args.config)

    if not os.path.exists(args.input):
        print(f"输入文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    # ---- 步骤2：LLM 分类 + 自动排版 ----
    if args.skip_typeset:
        typeset_path = args.input
        print(f"[main] 跳过步骤2，直接使用输入文件: {typeset_path}")
    else:
        print(f"[main] 步骤2：自动排版 {args.input}")
        typeset_path = typeset_doc(cfg, args.input)

    # ---- 步骤3：合并模板，替换 BodyText 节 ----
    if args.output:
        out_path = args.output
    else:
        base = os.path.splitext(os.path.basename(args.input))[0]
        out_path = os.path.join(cfg.output_dir, f"{base}_final.docx")
    print(f"[main] 步骤3：合并模板 -> {out_path}")
    merge_template(cfg, typeset_path, out_path)

    # ---- 步骤4：后续检查逻辑（占位，待实现） ----
    # 预留：例如校验标题层级连续性、表格样式覆盖率、图片题注配对等
    print("[main] 步骤4：后续检查逻辑（占位，待实现）")

    print(f"\n[main] 全部完成！最终文件: {out_path}")


if __name__ == "__main__":
    main()
