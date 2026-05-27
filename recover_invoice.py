#!/usr/bin/env python3
"""
发票恢复工具：从「已删除」文件夹回溯今天 00:00 起的发票邮件。

背景：早前因模板占位符 %fphm% 未被运营平台渲染，多封邮件被错误落盘为
      `invoice/<姓名>-<地区>/%fphm%.jpg`，互相覆盖造成丢失。
      由于这些邮件被处理后仍移动到「已删除」文件夹（默认保留 30 天），
      只要在保留期内即可重新抽取并产出正确文件名 `{发票号}.jpg`。

设计：
- 复用 main.py 的全部抽取/下载/落盘逻辑（process_invoice_mail / collect_pdf_from_mail / ...）
- 自适应识别 163 / QQ 等邮箱的「已删除」文件夹（复用 find_trash_folder）
- 时间过滤双保险：IMAP `SEARCH SINCE <today>` + 本地 `INTERNALDATE >= 今天 00:00`
- 只读模式：用 BODY.PEEK[] 避免标已读，不移动 / 不删除任何邮件
- 跳过冲突：现有 `{发票号}.jpg` 落盘逻辑会用文件名去重；占位 `%fphm%.jpg` 与新文件天然不撞名

用法：
  python recover_invoice.py             # 真实恢复
  python recover_invoice.py --dry-run   # 只打印抽取结果，不落盘
  python recover_invoice.py --list-folders  # 列出每个邮箱的文件夹名（调试用）
"""

from __future__ import annotations

import argparse
import email
import imaplib
import re
import sys
from datetime import datetime, time
from email.utils import parsedate_to_datetime
from pathlib import Path

# 复用 main.py 已有逻辑（不修改 main.py）
from main import (
    BASE_DIR,
    connect_imap,
    decode_mime_words,
    find_trash_folder,
    parse_accounts,
    process_invoice_mail,
)


# ── IMAP 工具 ────────────────────────────────────────


def list_folders(mail: imaplib.IMAP4_SSL) -> list[str]:
    """列出当前账号所有文件夹原始名（调试用）。"""
    status, folders = mail.list()
    if status != "OK" or not folders:
        return []
    names = []
    for raw in folders:
        if not raw:
            continue
        line = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw
        m = re.search(r'"([^"]+)"\s*$', line)
        if m:
            names.append(m.group(1))
        else:
            parts = line.rsplit(" ", 1)
            if len(parts) == 2:
                names.append(parts[1].strip())
    return names


def select_mailbox(mail: imaplib.IMAP4_SSL, folder: str) -> int:
    """
    选中指定文件夹（只读），返回邮件总数；失败返回 -1。
    folder 可能含空格（如 "Deleted Messages"），需带引号。
    """
    quoted = f'"{folder}"' if " " in folder or any(ord(c) > 127 for c in folder) else folder
    status, data = mail.select(quoted, readonly=True)
    if status != "OK":
        # 部分服务器需要不带引号
        status, data = mail.select(folder, readonly=True)
        if status != "OK":
            return -1
    try:
        return int(data[0].decode())
    except Exception:
        return 0


def search_since_today(mail: imaplib.IMAP4_SSL, since_date=None) -> list[bytes]:
    """
    IMAP SEARCH SINCE <date> —— 按日（不含时分秒）粗筛。
    本地再用 INTERNALDATE 精确二次过滤到 >= 起点 00:00:00。
    """
    if since_date is None:
        since_date = datetime.now().date()
    # IMAP SINCE 接受 DD-Mon-YYYY 形式
    since_str = since_date.strftime("%d-%b-%Y")
    status, data = mail.uid("search", None, "SINCE", since_str)
    if status != "OK" or not data or not data[0]:
        return []
    return data[0].split()


def fetch_internaldate(mail: imaplib.IMAP4_SSL, uid: bytes) -> datetime | None:
    """获取邮件 INTERNALDATE（服务端收件时间）。"""
    status, data = mail.uid("fetch", uid, "(INTERNALDATE)")
    if status != "OK" or not data or not data[0]:
        return None
    raw = data[0]
    if isinstance(raw, tuple):
        raw = raw[0]
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")
    m = re.search(r'INTERNALDATE "([^"]+)"', raw)
    if not m:
        return None
    try:
        # IMAP INTERNALDATE 形如 "23-May-2026 00:54:12 +0800"
        return parsedate_to_datetime(m.group(1))
    except Exception:
        return None


def fetch_message(mail: imaplib.IMAP4_SSL, uid: bytes):
    """只读获取邮件正文（BODY.PEEK[] 不会标已读）。"""
    status, data = mail.uid("fetch", uid, "(BODY.PEEK[])")
    if status != "OK" or not data or data[0] is None:
        return None
    raw_email = data[0][1]
    return email.message_from_bytes(raw_email)


# ── 主流程 ────────────────────────────────────────────


def recover_from_account(
    account: dict,
    dry_run: bool = False,
    list_only: bool = False,
    since_date=None,
) -> tuple[int, int, int]:
    """
    在单个账号的「已删除」文件夹中回溯指定日期起的发票邮件。
    返回 (匹配发票邮件数, 成功落盘数, 跳过数)
    """
    if since_date is None:
        since_date = datetime.now().date()
    mail = connect_imap(account["server"], account["user"], account["password"])
    if mail is None:
        return (0, 0, 0)

    try:
        if list_only:
            print(f"  📂 {account['user']} 文件夹列表:")
            for name in list_folders(mail):
                print(f"     - {name}")
            return (0, 0, 0)

        trash = find_trash_folder(mail)
        if not trash:
            print(f"  ⚠️  未找到「已删除」文件夹，跳过 {account['user']}")
            print(f"     可用文件夹: {list_folders(mail)}")
            return (0, 0, 0)
        print(f"  🗑️  使用「已删除」文件夹: {trash}")

        total = select_mailbox(mail, trash)
        if total < 0:
            print(f"  ❌ 无法打开 {trash}")
            return (0, 0, 0)
        print(f"     共 {total} 封邮件")

        uids = search_since_today(mail, since_date=since_date)
        if not uids:
            print(f"  📭 起点日期之后没有邮件")
            return (0, 0, 0)
        print(f"  📬 IMAP SINCE 粗筛得到 {len(uids)} 封")

        # 本地精确过滤：INTERNALDATE >= 起点 00:00
        today_start = datetime.combine(since_date, time.min)

        matched, ok_cnt, skip_cnt = 0, 0, 0
        for uid in uids:
            idate = fetch_internaldate(mail, uid)
            if idate is not None:
                # IMAP 返回带时区，统一去 tz 比对（按本地时间近似即可）
                idate_naive = idate.replace(tzinfo=None) if idate.tzinfo else idate
                if idate_naive < today_start:
                    continue

            msg = fetch_message(mail, uid)
            if msg is None:
                continue

            subject = decode_mime_words(msg.get("Subject", ""))
            sender = decode_mime_words(msg.get("From", ""))

            if "发票" not in subject and "电子报销凭证" not in subject:
                continue

            matched += 1
            print(f"\n  ── [{idate}] UID={uid.decode() if isinstance(uid, bytes) else uid}")

            if dry_run:
                # dry-run：只跑字段抽取一段，不下载、不落盘
                from main import (
                    extract_invoice_fields,
                    get_email_html_body,
                    get_email_text_body,
                )

                html = get_email_html_body(msg)
                text = get_email_text_body(msg)
                fields = extract_invoice_fields(subject, html, text)
                print(f"     主题: {subject}")
                print(f"     发件: {sender}")
                print(f"     抽取: {fields}")
                continue

            try:
                ok = process_invoice_mail(msg, subject, sender)
            except Exception as e:
                print(f"  ❌ 处理异常: {e}")
                ok = False

            if ok:
                ok_cnt += 1
            else:
                skip_cnt += 1

        return (matched, ok_cnt, skip_cnt)
    finally:
        try:
            mail.logout()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(description="从已删除邮箱回溯指定日期起的发票")
    parser.add_argument("--dry-run", action="store_true", help="只打印抽取结果，不下载、不落盘")
    parser.add_argument("--list-folders", action="store_true", help="列出每个账号的文件夹名（调试）")
    parser.add_argument("--since", type=str, default=None,
                        help="回溯起点日期 YYYY-MM-DD（含当日 00:00 起），默认今天")
    args = parser.parse_args()

    accounts = parse_accounts()
    if not accounts:
        print("❌ 未配置邮箱账号，请在 .env 文件中设置 EMAIL_ACCOUNTS")
        sys.exit(1)

    if args.since:
        try:
            since_date = datetime.strptime(args.since, "%Y-%m-%d").date()
        except ValueError:
            print(f"❌ --since 格式错误，应为 YYYY-MM-DD，收到: {args.since}")
            sys.exit(1)
    else:
        since_date = datetime.now().date()

    print(f"🕒 回溯时间起点: {since_date.strftime('%Y-%m-%d')} 00:00:00 (本地)")
    print(f"📁 输出目录: {BASE_DIR.resolve()}")
    print(f"🔧 模式: " + (
        "list-folders" if args.list_folders else ("dry-run（只读探查）" if args.dry_run else "实际恢复")
    ))
    print()

    grand_matched, grand_ok, grand_skip = 0, 0, 0
    for i, account in enumerate(accounts):
        if i > 0:
            print("\n" + "=" * 50 + "\n")
        m, ok, sk = recover_from_account(
            account, dry_run=args.dry_run, list_only=args.list_folders,
            since_date=since_date,
        )
        grand_matched += m
        grand_ok += ok
        grand_skip += sk

    if not args.list_folders:
        print("\n" + "=" * 50)
        print(f"✅ 完成！匹配发票邮件 {grand_matched} 封，"
              f"成功落盘 {grand_ok} 封，失败/跳过 {grand_skip} 封")
        if args.dry_run:
            print("ℹ️  当前为 dry-run，未真正落盘。确认抽取结果后去掉 --dry-run 重跑。")


if __name__ == "__main__":
    main()
