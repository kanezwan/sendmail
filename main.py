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
from urllib.parse import unquote
import requests
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


# ── 财云通（bigfintax）类型发票 ────────────────────────
BIGFINTAX_URL_RE = re.compile(
    r"https?://gateway\.bigfintax\.com/scanning-invoice/checkInvoice\?id=(\d+)"
)
BIGFINTAX_DOWNLOAD_API = (
    "https://gateway.bigfintax.com/xxApi/api/v2/electronInvoice/invoiceBatchDownload"
)
BIGFINTAX_UA = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Referer": "https://gateway.bigfintax.com/scanning-invoice/",
}


def extract_region(full_name: str) -> tuple[str, str]:
    """
    从销方名称中拆出地区 + 公司名。
    - 优先按 省/市/区/县/州 行政区后缀切分
    - 否则取前 2 个字作为地区
    """
    full_name = full_name.strip()
    m = re.match(r"^(.+?[省市区县州])(.*)", full_name)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    if len(full_name) > 2:
        return full_name[:2], full_name[2:]
    return "未知地区", full_name


def get_email_html_body(msg) -> str:
    """提取邮件中的 HTML 正文（不含附件）"""
    for part in msg.walk():
        ctype = part.get_content_type()
        disp = str(part.get("Content-Disposition", ""))
        if "attachment" in disp:
            continue
        if ctype == "text/html":
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            try:
                return payload.decode(charset, errors="replace")
            except Exception:
                return payload.decode("utf-8", errors="replace")
    return ""


def parse_bigfintax_mail(msg, sender: str, subject: str) -> dict | None:
    """
    判断并解析"财云通"类型发票邮件。
    识别条件：发件人含"财云通"或正文中含 bigfintax checkInvoice 链接
    返回：{"inv_id": str, "buyer": str | None, "region": str | None} 或 None
    """
    html = get_email_html_body(msg)
    m = BIGFINTAX_URL_RE.search(html) if html else None

    is_bigfintax = (
        "财云通" in sender
        or "ysb@szyh.com" in sender.lower()
        or m is not None
    )
    if not is_bigfintax or not m:
        return None

    inv_id = m.group(1)

    # 从 HTML 正文解析购方姓名、销方名称（作为 fallback，主路径是 Content-Disposition）
    buyer = None
    seller = None
    bm = re.search(r"购方名称[：:]\s*([^\s<]+)", html)
    if bm:
        buyer = bm.group(1).strip()
    sm = re.search(r"销方名称[：:]\s*([^\s<]+)", html)
    if sm:
        seller = sm.group(1).strip()

    region = None
    if seller:
        region, _ = extract_region(seller)

    return {"inv_id": inv_id, "buyer": buyer, "region": region}


def parse_content_disposition_filename(header: str) -> str | None:
    """
    解析 Content-Disposition 里的 filename 字段（已做 percent-decode）。
    兼容 filename="xxx" / filename=xxx / filename*=UTF-8''xxx 格式。
    """
    if not header:
        return None
    # RFC 5987: filename*=UTF-8''xxx
    m = re.search(r"filename\*\s*=\s*[^']*''([^;]+)", header, re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip().strip('"'))
    # 普通 filename=xxx
    m = re.search(r"filename\s*=\s*\"?([^\";]+)\"?", header, re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip())
    return None


def download_bigfintax_pdf(inv_id: str) -> tuple[bytes, str | None] | None:
    """
    调用财云通下载接口获取 PDF 二进制。
    返回 (pdf_bytes, filename) 或 None。
    filename 来自 Content-Disposition，格式通常为：
        电子发票（普通发票）_{发票号}_{销方名称}_{购方姓名}_{开票日期}.pdf
    """
    try:
        r = requests.get(
            BIGFINTAX_DOWNLOAD_API,
            params={"id": inv_id, "downloadType": "1"},
            headers=BIGFINTAX_UA,
            timeout=30,
        )
    except Exception as e:
        print(f"  ❌ 下载请求失败: {e}")
        return None

    if r.status_code != 200:
        print(f"  ❌ 下载失败: HTTP {r.status_code}")
        return None

    if not r.content.startswith(b"%PDF"):
        print(f"  ❌ 返回非 PDF 内容: {r.content[:80]!r}")
        return None

    filename = parse_content_disposition_filename(r.headers.get("Content-Disposition", ""))
    return r.content, filename


def parse_bigfintax_filename(filename: str) -> dict:
    """
    解析财云通下载接口返回的文件名：
        电子发票（普通发票）_{发票号}_{销方名称}_{购方姓名}_{开票日期}.pdf
    返回 {"invoice_no", "seller", "buyer", "date"}，字段缺失则为 None。
    """
    stem = Path(filename).stem
    parts = stem.split("_")
    result = {"invoice_no": None, "seller": None, "buyer": None, "date": None}
    # 形态 1：前缀(电子发票...) + 4 段
    if len(parts) >= 5:
        result["invoice_no"] = parts[-4]
        result["seller"] = parts[-3]
        result["buyer"] = parts[-2]
        result["date"] = parts[-1]
    elif len(parts) == 4:
        result["invoice_no"], result["seller"], result["buyer"], result["date"] = parts
    return result


def process_bigfintax_invoice(
    msg, sender: str, subject: str
) -> bool:
    """
    处理财云通类型发票邮件：
    1. 解析邮件找到 inv_id
    2. 调用下载 API 获取 PDF
    3. 从返回的文件名解析姓名/地区/发票号
    4. PDF → JPG，保存至 invoice/{姓名}-{地区}/{发票号}.jpg

    成功返回 True，非该类型或失败返回 False。
    """
    parsed = parse_bigfintax_mail(msg, sender, subject)
    if not parsed:
        return False

    inv_id = parsed["inv_id"]
    print(f"── 邮件: {subject}")
    print(f"   类型: 财云通(bigfintax) | inv_id={inv_id}")

    # 下载 PDF
    result = download_bigfintax_pdf(inv_id)
    if not result:
        return False
    pdf_bytes, dl_filename = result

    # 优先用 Content-Disposition 的文件名解析信息
    info = parse_bigfintax_filename(dl_filename) if dl_filename else {}
    invoice_no = info.get("invoice_no")
    buyer = info.get("buyer") or parsed.get("buyer")
    seller = info.get("seller")

    # 若销方已知，用其推地区；否则用邮件正文解析到的 region
    if seller:
        region, _ = extract_region(seller)
    else:
        region = parsed.get("region")

    # 发票号兜底：从主题取
    if not invoice_no:
        sm = re.search(r"发票号码[：:\s]*(\d+)", subject)
        if sm:
            invoice_no = sm.group(1)

    if not buyer:
        buyer = "未知"
    if not region:
        region = "未知地区"
    if not invoice_no:
        invoice_no = inv_id  # 实在没拿到就用 id 兜底

    print(f"   姓名: {buyer} | 地区: {region} | 发票号: {invoice_no}")

    # 保存路径
    save_dir = BASE_DIR / f"{buyer}-{region}"
    save_dir.mkdir(parents=True, exist_ok=True)
    jpg_path = save_dir / f"{invoice_no}.jpg"

    if jpg_path.exists():
        print(f"  ⏭️  已存在，跳过: {jpg_path.name}")
        return True

    # 写入临时 PDF 后转 JPG
    pdf_path = save_dir / f"{invoice_no}.pdf"
    try:
        pdf_path.write_bytes(pdf_bytes)
        print(f"  📥 已下载: {pdf_path.name} ({len(pdf_bytes)} bytes)")
        try:
            pdf_to_jpg(pdf_path, jpg_path)
        except Exception as e:
            print(f"  ❌ PDF 转换失败: {e}")
            return False
    finally:
        pdf_path.unlink(missing_ok=True)

    return True


# ── 朴朴超市（pupumall）类型发票 ──────────────────────
PUPU_PDF_URL_RE = re.compile(
    r"""href=['"](https://finance-files\.pupumall\.com/[^'"]+\.pdf)['"]"""
)


def extract_pdf_text(pdf_path: Path) -> str:
    """使用 pypdfium2 提取 PDF 第一页文本"""
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page = pdf[0]
        tp = page.get_textpage()
        return tp.get_text_bounded()
    finally:
        pdf.close()


def parse_pupu_pdf_text(text: str) -> dict:
    """
    从朴朴发票 PDF 提取的文本中解析字段。
    返回 {"invoice_no", "buyer", "seller"}，未识别字段为 None。
    """
    result = {"invoice_no": None, "buyer": None, "seller": None}

    # 发票号码：电子发票号码固定 20 位数字
    m = re.search(r"\b(\d{20})\b", text)
    if m:
        result["invoice_no"] = m.group(1)

    # 销方名称：优先匹配 "X市XXX有限公司"，兜底匹配任意 "XXX有限公司"
    m = re.search(r"([\u4e00-\u9fa5]+市[\u4e00-\u9fa5]+?有限公司)", text)
    if not m:
        m = re.search(r"([\u4e00-\u9fa5]{2,}?有限公司)", text)
    if m:
        result["seller"] = m.group(1)

    # 购买方姓名：销方公司名所在行的前面那段中文（2-4 字），格式 "{姓名} {销方}"
    if result["seller"]:
        m = re.search(
            r"([\u4e00-\u9fa5]{2,4})\s+" + re.escape(result["seller"]),
            text,
        )
        if m:
            result["buyer"] = m.group(1)

    return result


def process_pupu_invoice(msg, sender: str, subject: str) -> bool:
    """
    处理朴朴超市类型发票邮件：
    1. 识别类型（发件人 pupumall.net 或 主题含"朴朴超市-电子发票通知"）
    2. 从 HTML 正文提取 PDF 链接
    3. requests.get 下载 PDF
    4. pypdfium2 提取文本，解析 发票号/销方/购方姓名
    5. PDF → JPG，保存到 invoice/{姓名}-{地区}/{发票号}.jpg
    """
    is_pupu = (
        "pupumall.net" in sender.lower()
        or "朴朴超市" in sender
        or "朴朴超市-电子发票" in subject
    )
    if not is_pupu:
        return False

    html = get_email_html_body(msg)
    if not html:
        return False

    m = PUPU_PDF_URL_RE.search(html)
    if not m:
        print(f"── 邮件: {subject}")
        print("   类型: 朴朴超市 | ❌ 正文未找到 PDF 链接")
        return False

    pdf_url = m.group(1)
    print(f"── 邮件: {subject}")
    print(f"   类型: 朴朴超市(pupumall) | url={pdf_url[:80]}...")

    # 下载 PDF
    try:
        r = requests.get(
            pdf_url,
            headers={"User-Agent": "Mozilla/5.0 Chrome/124.0"},
            timeout=30,
        )
    except Exception as e:
        print(f"  ❌ 下载请求失败: {e}")
        return False

    if r.status_code != 200 or not r.content.startswith(b"%PDF"):
        print(f"  ❌ 下载失败: HTTP {r.status_code}, 首字节 {r.content[:8]!r}")
        return False

    pdf_bytes = r.content
    print(f"  📥 已下载: {len(pdf_bytes)} bytes")

    # 先写入临时 PDF（在临时目录，解析完再决定最终保存路径）
    tmp_pdf = BASE_DIR / f".pupu_tmp_{os.getpid()}.pdf"
    tmp_pdf.parent.mkdir(parents=True, exist_ok=True)
    tmp_pdf.write_bytes(pdf_bytes)

    try:
        # 提取文本解析字段
        try:
            text = extract_pdf_text(tmp_pdf)
        except Exception as e:
            print(f"  ❌ PDF 文本提取失败: {e}")
            return False

        info = parse_pupu_pdf_text(text)
        invoice_no = info["invoice_no"]
        buyer = info["buyer"]
        seller = info["seller"]

        # 兜底
        if not invoice_no:
            # 用 URL 里的 hash 末 12 位兜底
            url_hash = re.search(r"/([0-9a-f]{20,})\.pdf$", pdf_url)
            invoice_no = url_hash.group(1)[-12:] if url_hash else "unknown"
            print(f"  ⚠️  未解析到发票号，使用兜底: {invoice_no}")

        if not buyer:
            buyer = "未知"
            print("  ⚠️  未解析到购买方姓名")

        if seller:
            region, _ = extract_region(seller)
        else:
            region = "未知地区"
            print("  ⚠️  未解析到销方名称")

        print(f"   姓名: {buyer} | 地区: {region} | 发票号: {invoice_no}")

        # 保存路径
        save_dir = BASE_DIR / f"{buyer}-{region}"
        save_dir.mkdir(parents=True, exist_ok=True)
        jpg_path = save_dir / f"{invoice_no}.jpg"

        if jpg_path.exists():
            print(f"  ⏭️  已存在，跳过: {jpg_path.name}")
            return True

        # PDF → JPG
        final_pdf = save_dir / f"{invoice_no}.pdf"
        try:
            shutil.move(str(tmp_pdf), str(final_pdf))
            try:
                pdf_to_jpg(final_pdf, jpg_path)
            except Exception as e:
                print(f"  ❌ PDF 转换失败: {e}")
                return False
        finally:
            final_pdf.unlink(missing_ok=True)

        return True
    finally:
        tmp_pdf.unlink(missing_ok=True)


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

        # 解析主题提取名称和地区（旧类型：票通）
        parsed = parse_subject(subject)
        if not parsed:
            # 尝试财云通（bigfintax）类型
            if process_bigfintax_invoice(msg, sender, subject):
                processed_count += 1
                delete_ids.append(msg_uid)
                print()
            # 尝试朴朴超市（pupumall）类型
            elif process_pupu_invoice(msg, sender, subject):
                processed_count += 1
                delete_ids.append(msg_uid)
                print()
            else:
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

    # 统一将已处理的发票邮件移动到回收站（不彻底删除）
    if delete_ids:
        trash_folder = find_trash_folder(mail)
        if trash_folder:
            moved = move_to_trash(mail, delete_ids, trash_folder)
            print(f"  🗑️  已将 {moved} 封发票邮件移动到回收站: {trash_folder}")
        else:
            print(f"  ⚠️  未找到回收站文件夹，邮件未删除")

    return (processed_count, skipped_count)


# 常见的回收站文件夹名称（按优先级排序）
TRASH_CANDIDATES = [
    "&XfJT0ZAB-",           # 163 "已删除" (Modified UTF-7 编码)
    "Deleted Messages",     # QQ 邮箱
    "&XfJSIJZk-",           # 部分邮箱的"垃圾邮件"（非回收站，仅备用）
    "Trash",
    "Deleted",
    "Deleted Items",
    "已删除",
    "回收站",
    "垃圾箱",
    "INBOX.Trash",
]


def find_trash_folder(mail: imaplib.IMAP4_SSL) -> str | None:
    """查找邮箱的回收站文件夹名称"""
    status, folders = mail.list()
    if status != "OK" or not folders:
        return None

    folder_names = []
    for raw in folders:
        if not raw:
            continue
        line = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw
        # LIST 响应格式: (\HasNoChildren) "/" "INBOX"
        # 提取最后一段带引号的文件夹名
        m = re.search(r'"([^"]+)"\s*$', line)
        if m:
            folder_names.append(m.group(1))
        else:
            # 末尾没有引号，取最后一个空格分隔的部分
            parts = line.rsplit(" ", 1)
            if len(parts) == 2:
                folder_names.append(parts[1].strip())

    # 优先匹配 TRASH_CANDIDATES 中的名称
    for candidate in TRASH_CANDIDATES:
        for name in folder_names:
            if name == candidate:
                return name

    # 再匹配包含关键字的文件夹
    keywords = ["trash", "deleted", "已删除", "回收站", "垃圾箱"]
    for name in folder_names:
        low = name.lower()
        if any(k in low or k in name for k in keywords):
            return name

    return None


def move_to_trash(mail: imaplib.IMAP4_SSL, uids: list, trash_folder: str) -> int:
    """
    将邮件移动到回收站。
    优先使用 IMAP MOVE 扩展（RFC 6851），不支持则回退到 COPY + DELETE + EXPUNGE。
    返回成功移动的邮件数。
    """
    moved = 0
    # 文件夹名含空格时 IMAP 要求用双引号包起来
    quoted_folder = f'"{trash_folder}"' if " " in trash_folder else trash_folder
    # 尝试使用 UID MOVE（原子操作，速度快）
    for uid in uids:
        try:
            status, _ = mail.uid("MOVE", uid, quoted_folder)
            if status == "OK":
                moved += 1
                continue
        except Exception:
            pass

        # 回退：COPY + STORE \Deleted + EXPUNGE
        try:
            status, _ = mail.uid("COPY", uid, quoted_folder)
            if status == "OK":
                mail.uid("STORE", uid, "+FLAGS", "\\Deleted")
                moved += 1
        except Exception as e:
            print(f"  ⚠️  移动邮件 UID={uid.decode() if isinstance(uid, bytes) else uid} 失败: {e}")

    # 对于 COPY+DELETE 的邮件，执行 expunge 清除原位置
    try:
        mail.expunge()
    except Exception:
        pass

    return moved


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
