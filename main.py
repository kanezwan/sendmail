#!/usr/bin/env python3
"""
发票邮件处理工具
- 登录多个邮箱 (IMAP)
- 下载发票邮件中的 PDF 附件
- 将 PDF 转换为 JPG 图片
- 按 {名称}-{地区} 目录保存
"""

import imaplib
import email
from email.header import decode_header
import os
import re
import shutil
import sys
import pypdfium2 as pdfium
from PIL import Image
from pathlib import Path


# ── 配置 ─────────────────────────────────────────────
def load_env(env_path: str = ".env"):
    """从 .env 文件加载环境变量"""
    env_file = Path(env_path)
    if not env_file.exists():
        return
    with open(env_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, value = line.split("=", 1)
                os.environ[key.strip()] = value.strip()


load_env()

BASE_DIR = Path(__file__).parent / "invoice"
IMAP_PORT = 993


def parse_accounts() -> list[dict]:
    """
    解析邮箱账号配置。
    格式: 用户名:密码:IMAP服务器;用户名:密码:IMAP服务器
    """
    raw = os.environ.get("EMAIL_ACCOUNTS", "")
    if not raw:
        return []
    accounts = []
    for entry in raw.split(";"):
        entry = entry.strip()
        if not entry:
            continue
        parts = entry.split(":")
        if len(parts) >= 3:
            accounts.append({
                "user": parts[0].strip(),
                "password": parts[1].strip(),
                "server": parts[2].strip(),
            })
    return accounts


# ── 工具函数 ──────────────────────────────────────────
def decode_mime_words(s: str) -> str:
    """解码 MIME 编码的邮件头"""
    if s is None:
        return ""
    decoded_parts = []
    for part, charset in decode_header(s):
        if isinstance(part, bytes):
            decoded_parts.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded_parts.append(part)
    return "".join(decoded_parts)


def parse_subject(subject: str) -> tuple[str, str] | None:
    """
    从邮件主题解析公司名称和地区。
    主题格式：您收到一张来自{地区}{公司名称}的电子发票【发票金额：xxx】
    返回：(名称, 地区) 或 None
    """
    pattern = r"来自(.+?)的电子发票"
    match = re.search(pattern, subject)
    if not match:
        return None

    full_name = match.group(1).strip()

    # 尝试拆分地区和公司名称
    # 地区通常以 省/市/区/县/州 结尾
    region_pattern = r"^(.+?[省市区县州])(.*)"
    region_match = re.match(region_pattern, full_name)
    if region_match:
        region = region_match.group(1)
        company = region_match.group(2)
    else:
        # 没有明确行政区后缀，取前 2 个字作为地区
        if len(full_name) > 2:
            region = full_name[:2]
            company = full_name[2:]
        else:
            region = "未知地区"
            company = full_name

    return (company.strip(), region.strip())


def pdf_to_jpg(pdf_path: Path, jpg_path: Path, dpi: int = 200):
    """将 PDF 文件转换为 JPG 图片（取第一页）"""
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page = pdf[0]
        scale = dpi / 72
        bitmap = page.render(scale=scale)
        pil_image = bitmap.to_pil()
        if pil_image.mode == "RGBA":
            background = Image.new("RGB", pil_image.size, (255, 255, 255))
            background.paste(pil_image, mask=pil_image.split()[3])
            pil_image = background
        pil_image.save(str(jpg_path), "JPEG", quality=95)
        print(f"  ✅ 已转换: {jpg_path.name}")
    finally:
        pdf.close()


def connect_imap(server: str, user: str, password: str) -> imaplib.IMAP4_SSL | None:
    """连接并登录 IMAP 服务器，失败返回 None"""
    print(f"📧 正在连接邮箱 {user} ({server}) ...")
    try:
        # 注册 ID 命令（163邮箱需要）
        imaplib.Commands["ID"] = ("AUTH",)

        mail = imaplib.IMAP4_SSL(server, IMAP_PORT)
        mail.login(user, password)

        # 发送 IMAP ID 命令声明客户端身份（解决163邮箱 Unsafe Login）
        args = (
            '"name" "InvoiceTool" '
            '"contact" "invoice@tool.com" '
            '"version" "1.0.0" '
            '"vendor" "InvoiceTool"'
        )
        mail._simple_command("ID", f"({args})")
    except imaplib.IMAP4.error as e:
        print(f"  ❌ 登录失败: {e}")
        return None
    except Exception as e:
        print(f"  ❌ 连接失败: {e}")
        return None
    print("  ✅ 登录成功")
    return mail


def process_mailbox(mail: imaplib.IMAP4_SSL, account_label: str) -> tuple[int, int]:
    """
    处理单个邮箱的所有发票邮件。
    返回 (已处理数, 跳过数)
    """
    processed_count = 0
    skipped_count = 0
    delete_ids = []  # 收集待删除的邮件 ID

    # 选择收件箱（读写模式，处理完后删除邮件）
    status, data = mail.select("INBOX")
    if status != "OK":
        print(f"  ❌ 无法打开收件箱: {data}")
        return (0, 0)
    total = data[0].decode()
    print(f"  📂 收件箱打开成功，共 {total} 封邮件")

    # 搜索所有邮件（使用 UID 避免序列号漂移导致删除错误）
    status, messages = mail.uid("search", None, "ALL")
    if status != "OK" or not messages[0]:
        print("  📭 收件箱中没有邮件")
        return (0, 0)

    msg_uids = messages[0].split()
    print(f"  📬 找到 {len(msg_uids)} 封邮件，开始处理...\n")

    for msg_uid in msg_uids:
        status, msg_data = mail.uid("fetch", msg_uid, "(RFC822)")
        if status != "OK" or msg_data[0] is None:
            continue

        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = decode_mime_words(msg.get("Subject", ""))
        sender = decode_mime_words(msg.get("From", ""))

        # 解析主题提取名称和地区
        parsed = parse_subject(subject)
        if not parsed:
            skipped_count += 1
            continue

        company, region = parsed

        # 遍历附件，找到 PDF 文件
        pdf_found = False
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" not in content_disposition:
                continue

            filename = part.get_filename()
            if filename:
                filename = decode_mime_words(filename)

            if not filename or not filename.lower().endswith(".pdf"):
                continue

            pdf_data = part.get_payload(decode=True)
            if not pdf_data:
                continue

            pdf_found = True

            # 从文件名提取姓名：格式为 xxx_xxx_姓名.pdf
            stem = Path(filename).stem  # 如 264420000036_92402671_彭翠华
            parts = stem.rsplit("_", 1)
            if len(parts) == 2:
                name_part = parts[1]       # 彭翠华
                number_part = parts[0]     # 264420000036_92402671
            else:
                name_part = "未知"
                number_part = stem

            # 目录: 姓名-地区
            save_dir = BASE_DIR / f"{name_part}-{region}"
            save_dir.mkdir(parents=True, exist_ok=True)

            # 文件名: 去掉姓名，只保留编号
            jpg_filename = number_part + ".jpg"
            jpg_path = save_dir / jpg_filename

            print(f"── 邮件: {subject}")
            print(f"   姓名: {name_part} | 地区: {region}")

            # 跳过已存在的文件
            if jpg_path.exists():
                print(f"  ⏭️  已存在，跳过: {jpg_filename}")
                continue

            pdf_path = save_dir / filename

            # 保存 PDF（临时）
            with open(pdf_path, "wb") as f:
                f.write(pdf_data)
            print(f"  📥 已下载: {filename}")

            # 转换为 JPG
            try:
                pdf_to_jpg(pdf_path, jpg_path)
            except Exception as e:
                print(f"  ❌ PDF 转换失败: {e}")
            finally:
                pdf_path.unlink(missing_ok=True)

        if pdf_found:
            processed_count += 1
            delete_ids.append(msg_uid)
            print()
        else:
            skipped_count += 1

    # 统一标记删除并执行（使用 UID 确保删除正确的邮件）
    if delete_ids:
        for uid in delete_ids:
            mail.uid("store", uid, "+FLAGS", "\\Deleted")
        mail.expunge()
        print(f"  🗑️  已从服务器删除 {len(delete_ids)} 封已处理的发票邮件")

    return (processed_count, skipped_count)


def clean_invoice_dir():
    """清空 invoice/ 目录下所有文件和子目录"""
    if BASE_DIR.exists():
        shutil.rmtree(BASE_DIR)
        print(f"🧹 已清空目录: {BASE_DIR}")
    BASE_DIR.mkdir(parents=True, exist_ok=True)


def main():
    """主入口：遍历所有配置的邮箱"""
    accounts = parse_accounts()
    if not accounts:
        print("❌ 未配置邮箱账号，请在 .env 文件中设置 EMAIL_ACCOUNTS")
        print("   格式: 用户名:密码:IMAP服务器;用户名:密码:IMAP服务器")
        sys.exit(1)

    # 启动前清空 invoice/ 目录
    clean_invoice_dir()

    total_processed = 0
    total_skipped = 0

    for i, account in enumerate(accounts):
        if i > 0:
            print("\n" + "=" * 50 + "\n")

        mail = connect_imap(account["server"], account["user"], account["password"])
        if mail is None:
            print(f"  ⚠️  跳过此邮箱\n")
            continue

        try:
            processed, skipped = process_mailbox(mail, account["user"])
            total_processed += processed
            total_skipped += skipped
        finally:
            mail.logout()
            print(f"📧 已断开 {account['user']}")

    print("\n" + "=" * 50)
    print(f"✅ 全部完成！共处理 {total_processed} 封发票邮件，跳过 {total_skipped} 封非发票邮件")
    print(f"📁 发票图片保存在: {BASE_DIR.resolve()}")


if __name__ == "__main__":
    main()
