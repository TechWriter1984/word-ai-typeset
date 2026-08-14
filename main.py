"""一键启动入口：步骤2 → 步骤3 → 步骤4（post_process，含删标记/占位符/术语表/图片表格收尾）。

用法：
    python main.py <前处理后的 docx 路径>
    python main.py <path> --skip-typeset          # 跳过 LLM 分类，直接合并
    python main.py <path> --output D:\\out.docx    # 指定最终输出路径
    python main.py <path> --config D:\\cfg.yaml    # 指定配置文件
    python main.py <path> --skip-post-process      # 跳过步骤4（只出合并稿，不做占位符等收尾）
"""
from __future__ import annotations

import argparse
import os
import sys

from config import load_config
from typeset import typeset_doc
from merge import merge_template
from post_process import run_post_process


def parse_args():
    p = argparse.ArgumentParser(description="Word AI 自动排版 + 模板合并 + 后处理")
    p.add_argument("input", help="用户前处理后的 docx 路径（步骤2输入）")
    p.add_argument("--config", default=None, help="config.yaml 路径（默认项目根）")
    p.add_argument("--output", default=None,
                   help="最终输出 docx 路径（默认 output/<base>_final.docx）")
    p.add_argument("--skip-typeset", action="store_true",
                   help="跳过步骤2，直接对输入文件做合并（输入需已排版）")
    p.add_argument("--skip-post-process", action="store_true",
                   help="跳过步骤4后处理（仅合并稿，不做删标记/占位符/术语表/图片表格收尾）")
    return p.parse_args()


def main():
    args = parse_args()
    cfg = load_config(args.config)

    if not os.path.exists(args.input):
        print(f"输入文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    base = os.path.splitext(os.path.basename(args.input))[0]

    # ---- 步骤2：LLM 分类 + 自动排版 ----
    if args.skip_typeset:
        typeset_path = args.input
        print(f"[main] 跳过步骤2，直接使用输入文件: {typeset_path}")
    else:
        print(f"[main] 步骤2：自动排版 {args.input}")
        typeset_path = typeset_doc(cfg, args.input)

    # ---- 步骤3：合并模板，替换 BodyText 节 ----
    if args.output:
        merged_path = os.path.join(cfg.output_dir, f"{base}_merged.docx")
        final_path = args.output
    else:
        merged_path = os.path.join(cfg.output_dir, f"{base}_merged.docx")
        final_path = os.path.join(cfg.output_dir, f"{base}_final.docx")

    print(f"[main] 步骤3：合并模板 -> {merged_path}")
    merge_template(cfg, typeset_path, merged_path)

    # ---- 步骤4：后处理（删标记/占位符/术语表/图片表格收尾）----
    if args.skip_post_process:
        if args.output:
            # 用户指定输出路径，把 merged 直接 copy 过去
            import shutil
            os.makedirs(os.path.dirname(os.path.abspath(final_path)), exist_ok=True)
            shutil.copyfile(merged_path, final_path)
        else:
            final_path = merged_path
        print(f"[main] 跳过步骤4后处理（--skip-post-process）")
    else:
        print(f"[main] 步骤4：后处理 -> {final_path}")
        run_post_process(cfg, merged_path, final_path)

    print(f"\n[main] 全部完成！最终文件: {final_path}")


if __name__ == "__main__":
    main()
