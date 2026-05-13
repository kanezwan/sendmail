#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票剪贴板助手：按目录顺序，把"标题文本 + 各张发票图"依次推到 macOS 剪贴板，
每按一次回车进入下一步，配合微信 Cmd+V 快速粘贴发送。

用法:
    python send_invoice.py [invoice_dir]
默认 invoice_dir = ./invoice
"""
from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path

IMAGE_EXTS = {".jpg", ".jpeg", ".png"}


def copy_text(text: str) -> None:
    """文字 → 剪贴板（pbcopy）"""
    p = subprocess.Popen(["pbcopy"], stdin=subprocess.PIPE)
    p.communicate(text.encode("utf-8"))


def copy_files(paths: list[Path]) -> None:
    """
    把多个本地文件以 "Finder 多选 + Cmd+C" 的方式塞进剪贴板。
    必须用 NSPasteboard 写 public.file-url + NSFilenamesPboardType，
    微信 Mac 客户端才会把它当作多张图片粘贴。
    （AppleScript 只能写出 list/furl 单项，微信不识别。）
    """
    if not paths:
        return
    try:
        from AppKit import NSPasteboard, NSURL  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "缺少 pyobjc，请运行: pip install pyobjc-framework-Cocoa"
        ) from e
    urls = [NSURL.fileURLWithPath_(str(p.resolve())) for p in paths]
    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    if not pb.writeObjects_(urls):
        raise RuntimeError("NSPasteboard 写入文件 URL 失败")


def dir_to_title(dir_name: str) -> str:
    """周典斌-上海 → 周典斌 - 上海"""
    if "-" in dir_name:
        name, city = dir_name.split("-", 1)
        return f"{name.strip()} - {city.strip()}"
    return dir_name


def list_invoice_dirs(root: Path) -> list[Path]:
    """列出 root 下所有含图片的子目录，按目录名排序（隐藏目录跳过）"""
    dirs = []
    for d in sorted(root.iterdir()):
        if not d.is_dir() or d.name.startswith("."):
            continue
        imgs = [f for f in d.iterdir() if f.suffix.lower() in IMAGE_EXTS]
        if imgs:
            dirs.append(d)
    return dirs


def list_images(d: Path) -> list[Path]:
    """目录下图片，按文件名排序"""
    return sorted(
        (f for f in d.iterdir() if f.suffix.lower() in IMAGE_EXTS),
        key=lambda x: x.name,
    )


def wait_enter(prompt: str) -> None:
    """阻塞等用户按回车（Ctrl+C 退出）"""
    try:
        input(prompt)
    except (EOFError, KeyboardInterrupt):
        print("\n已中止")
        sys.exit(0)


def main() -> None:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "invoice").resolve()
    if not root.is_dir():
        print(f"目录不存在: {root}")
        sys.exit(1)

    dirs = list_invoice_dirs(root)
    if not dirs:
        print(f"{root} 下没有可处理的目录")
        sys.exit(0)

    # 1. 列出所有目录
    print(f"\n📁 {root}")
    print(f"共 {len(dirs)} 个目录:")
    for i, d in enumerate(dirs, 1):
        imgs = list_images(d)
        print(f"  {i:>2}. {d.name}  ({len(imgs)} 张)")
    print()

    # 2. 顺序处理
    for i, d in enumerate(dirs, 1):
        title = dir_to_title(d.name)
        imgs = list_images(d)

        print(f"\n━━━ [{i}/{len(dirs)}] {d.name}  ({len(imgs)} 张图) ━━━")

        # 2.1 文本（末尾加换行，方便微信粘贴后直接接图片预览）
        wait_enter(f"  回车复制文本「{title}」… ")
        copy_text(title + "\n")
        print(f"  ✅ [1/2] 文本已复制 → 切到微信 Cmd+V")

        # 2.2 一次性复制全部图片（多文件，等同 Finder 多选 Cmd+C）
        wait_enter(f"  回车复制全部 {len(imgs)} 张图… ")
        copy_files(imgs)
        names = ", ".join(p.name for p in imgs)
        print(f"  ✅ [2/2] 已复制 {len(imgs)} 张图 → Cmd+V")
        print(f"        {names}")

    print("\n🎉 全部目录处理完成，退出")


if __name__ == "__main__":
    main()
