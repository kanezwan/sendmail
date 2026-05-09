#!/usr/bin/env python3
"""
发票邮件处理工具（通用版）

入口：邮件主题包含"发票"即视为发票邮件，进入统一处理流程。

通用流程：
1. 解析邮件正文（HTML/纯文本）+ 主题，按内置 label 表抽取
   姓名（发票抬头/购方名称/...）、发票号（数电号码/发票号码/...）、销方
2. 收集 PDF 候选源（按优先级）：
   a. 邮件附件 .pdf
   b. 正文 <a href> 中的 *.pdf 直链
   c. 正文中"查看发票"类链接 → 进入中转页 → 在 HTML/JS 里搜 .pdf 链接或下载 API
3. 拿到 PDF → 转 JPG → 保存至 invoice/{姓名}-{地区}/{发票号}.jpg
"""

import imaplib
import email
from email.header import decode_header
import os
import re
import shutil
import sys
from urllib.parse import unquote, urljoin, urlparse
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

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
}

# label 映射表（从邮件正文/主题里抽取字段时用）
NAME_LABELS = ["发票抬头", "购方名称", "购买方名称", "购买方", "抬头"]
INVOICE_LABELS = ["数电号码", "发票号码", "发票号"]
SELLER_LABELS = ["销方名称", "销售方名称", "销方", "销售方"]


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


# ── 通用工具 ──────────────────────────────────────────
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


def get_email_text_body(msg) -> str:
    """提取邮件中的 text/plain 正文（不含附件）"""
    for part in msg.walk():
        ctype = part.get_content_type()
        disp = str(part.get("Content-Disposition", ""))
        if "attachment" in disp:
            continue
        if ctype == "text/plain":
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            try:
                return payload.decode(charset, errors="replace")
            except Exception:
                return payload.decode("utf-8", errors="replace")
    return ""


def html_to_text(html: str) -> str:
    """粗略地把 HTML 转成纯文本（去标签 + 反转义实体）"""
    if not html:
        return ""
    # 去掉 <script>/<style> 整段
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.S | re.I)
    # 替换 <br> / </p> / </div> / </tr> 为换行，便于按行解析
    text = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", text)
    text = re.sub(r"(?i)</\s*(p|div|tr|li|h[1-6])\s*>", "\n", text)
    # 去掉所有标签
    text = re.sub(r"<[^>]+>", "", text)
    # HTML 实体反转义
    text = (
        text.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&#39;", "'")
        .replace("&yen;", "¥")
    )
    # 全角冒号统一
    text = text.replace("：", ":")
    return text


def extract_region(full_name: str) -> tuple[str, str]:
    """
    从销方名称中拆出地区 + 公司名。
    - 优先按 省/市/区/县/州 行政区后缀切分
    - 否则取前 2 个字作为地区
    """
    full_name = (full_name or "").strip()
    if not full_name:
        return "未知地区", ""
    m = re.match(r"^(.+?[省市区县州])(.*)", full_name)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    if len(full_name) > 2:
        return full_name[:2], full_name[2:]
    return "未知地区", full_name


def parse_content_disposition_filename(header: str) -> str | None:
    """
    解析 Content-Disposition 里的 filename 字段（已做 percent-decode）。
    兼容 filename="xxx" / filename=xxx / filename*=UTF-8''xxx 格式。
    """
    if not header:
        return None
    m = re.search(r"filename\*\s*=\s*[^']*''([^;]+)", header, re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip().strip('"'))
    m = re.search(r"filename\s*=\s*\"?([^\";]+)\"?", header, re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip())
    return None


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


def extract_pdf_text(pdf_path: Path) -> str:
    """使用 pypdfium2 提取 PDF 第一页文本"""
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page = pdf[0]
        tp = page.get_textpage()
        return tp.get_text_bounded()
    finally:
        pdf.close()


# ── 通用字段抽取 ──────────────────────────────────────
def _search_label(text: str, labels: list[str]) -> str | None:
    """
    依次按 label 在文本里查找 `{label}[：:]\\s*([^\\s<]+)`，找到即返回。
    text 应已统一全角冒号为半角。
    """
    if not text:
        return None
    for label in labels:
        m = re.search(rf"{re.escape(label)}\s*[:：]\s*([^\s<>]+)", text)
        if m:
            return m.group(1).strip().rstrip("，,；;。.")
    return None


def _seller_from_subject(subject: str) -> str | None:
    """
    从主题里抽销方：优先取【...】中包含"公司"或"店"的方括号内容。
    """
    if not subject:
        return None
    # 兼容【】和[]两种括号
    for m in re.finditer(r"[【\[]([^】\]]+)[】\]]", subject):
        s = m.group(1).strip()
        if any(k in s for k in ["公司", "店", "厂", "中心", "集团"]):
            return s
    return None


def extract_invoice_fields(subject: str, html: str, text: str) -> dict:
    """
    从主题 + HTML + 纯文本里抽取通用字段。
    返回：{"name": str|None, "invoice_no": str|None, "seller": str|None}
    """
    text_subject = (subject or "").replace("：", ":")
    text_body = text or ""
    # 把 HTML 转纯文本一起加进来（应对正文只有 HTML 的情况）
    text_html_plain = html_to_text(html) if html else ""

    haystacks = [text_subject, text_body, text_html_plain]
    combined = "\n".join(h for h in haystacks if h)

    name = _search_label(combined, NAME_LABELS)
    invoice_no = _search_label(combined, INVOICE_LABELS)
    seller = _search_label(combined, SELLER_LABELS) or _seller_from_subject(subject)

    # 发票号兜底：主题里 "发票号码：xxx" 已被 _search_label 覆盖；
    # 再兜底搜 12-20 位连续数字（避免误命中过短/过长的数字串）
    if not invoice_no:
        m = re.search(r"\b(\d{12,20})\b", combined)
        if m:
            invoice_no = m.group(1)

    return {"name": name, "invoice_no": invoice_no, "seller": seller}


# ── PDF 候选源收集 + 下载 ─────────────────────────────
PDF_DIRECT_RE = re.compile(r"""href=['"]([^'"]+\.pdf(?:\?[^'"]*)?)['"]""", re.IGNORECASE)
ANY_HREF_RE = re.compile(r"""href=['"](https?://[^'"<>]+)['"]""", re.IGNORECASE)
PDF_URL_IN_TEXT_RE = re.compile(r"""https?://[^\s'"<>]+\.pdf(?:\?[^\s'"<>]*)?""", re.IGNORECASE)


def _looks_like_pdf(content: bytes) -> bool:
    return content[:4] == b"%PDF"


def _fetch(url: str, session: requests.Session, **kwargs) -> requests.Response | None:
    """简单封装：失败返回 None，不抛异常"""
    try:
        return session.get(url, headers=DEFAULT_HEADERS, timeout=30, **kwargs)
    except Exception as e:
        print(f"  ⚠️  请求失败 {url[:80]}: {e}")
        return None


def _try_download_pdf(url: str, session: requests.Session, referer: str | None = None) -> bytes | None:
    """直接 GET，若返回是 PDF 则返回字节，否则 None"""
    headers = dict(DEFAULT_HEADERS)
    if referer:
        headers["Referer"] = referer
    try:
        r = session.get(url, headers=headers, timeout=30, allow_redirects=True)
    except Exception as e:
        print(f"  ⚠️  下载失败 {url[:80]}: {e}")
        return None
    if r.status_code != 200:
        return None
    ctype = r.headers.get("Content-Type", "").lower()
    if "pdf" in ctype or _looks_like_pdf(r.content):
        return r.content
    return None


# ── Landing-page 适配器 ───────────────────────────────
# 部分站点的中转页是 SPA（HTML 里没有 PDF 链接），需要按站点特定的 API 拉取。
# 适配器约定：输入 (final_url, session) → 返回 (pdf_bytes, fields_dict) 或 None。
# fields_dict 可包含 name / invoice_no / seller，用于补全主流程未拿到的字段。
def _adapter_nuonuo(final_url: str, session: requests.Session) -> tuple[bytes, dict] | None:
    """
    诺诺网（nnfp.jss.com.cn）短链中转页：
    1) 短链跳转后 URL 含 paramList / shortLinkSource / aliView 参数
    2) POST /scan2/getIvcDetailShow.do → 返回 invoiceSimpleVo
       - buyername / saleName / fphm / url(PDF 直链)
    3) GET vo.url 即得 PDF
    """
    from urllib.parse import urlparse, parse_qs
    qs = parse_qs(urlparse(final_url).query)
    param_list = qs.get("paramList", [""])[0]
    if not param_list:
        return None
    api = f"{urlparse(final_url).scheme}://{urlparse(final_url).netloc}/scan2/getIvcDetailShow.do"
    headers = dict(DEFAULT_HEADERS)
    headers["Referer"] = final_url
    headers["Content-Type"] = "application/x-www-form-urlencoded"
    data = {
        "paramList": param_list,
        "shortLinkSource": qs.get("shortLinkSource", [""])[0],
        "aliView": qs.get("aliView", [""])[0],
    }
    try:
        r = session.post(api, data=data, headers=headers, timeout=30)
        j = r.json()
    except Exception as e:
        print(f"  ⚠️  诺诺网 API 调用失败: {e}")
        return None
    if j.get("status") != "0000":
        return None
    vo = (j.get("data") or {}).get("invoiceSimpleVo") or {}
    pdf_url = vo.get("url")
    if not pdf_url:
        return None
    pdf = _try_download_pdf(pdf_url, session, referer=final_url)
    if not pdf:
        return None
    return pdf, {
        "name": vo.get("buyername") or vo.get("buyerName"),
        "invoice_no": vo.get("fphm") or vo.get("invoiceNo"),
        "seller": vo.get("saleName") or vo.get("taxPayerName"),
    }


def _adapter_bigfintax(final_url: str, session: requests.Session) -> tuple[bytes, dict] | None:
    """
    财云通（gateway.bigfintax.com）SPA 中转页：
    1) 邮件链接形如 .../scanning-invoice/checkInvoice?id=<id>
    2) 真实 PDF 下载：
       https://gateway.bigfintax.com/sopinv/invoice/out/fusion/templateDownload/1/<id>saas
       （type=1=PDF, type=2=OFD, type=3=XML；apply_serial_no = id + "saas"）
    3) 详情接口（AES-CBC 加密响应，可选用于补全字段）：
       GET /xxApi/api/v2/electronInvoice/getInvoiceInfo?id=<id>
    """
    from urllib.parse import urlparse, parse_qs
    parsed = urlparse(final_url)
    qs = parse_qs(parsed.query)
    inv_id = qs.get("id", [""])[0]
    if not inv_id:
        return None

    extra: dict = {}
    apply_serial_no = inv_id + "saas"

    # 先尝试详情 API 补全字段（AES 解密响应，失败也不影响 PDF 下载）
    try:
        from base64 import b64decode
        from json import loads as json_loads
        from Crypto.Cipher import AES
        from Crypto.Util.Padding import unpad

        info_api = f"{parsed.scheme}://{parsed.netloc}/xxApi/api/v2/electronInvoice/getInvoiceInfo"
        headers = dict(DEFAULT_HEADERS)
        headers["Referer"] = final_url
        headers["X-Requested-With"] = "XMLHttpRequest"
        rj = session.get(info_api, params={"id": inv_id}, headers=headers, timeout=30).json()
        if rj.get("result") == "success" and rj.get("data") and rj.get("key"):
            raw = b64decode(rj["data"])
            iv, ct = raw[:16], raw[16:]
            key = b64decode(rj["key"])
            plain = unpad(AES.new(key, AES.MODE_CBC, iv).decrypt(ct), AES.block_size)
            obj = json_loads(plain.decode("utf-8"))
            extra = {
                "name": obj.get("buyer"),
                "invoice_no": obj.get("invoiceNumeric"),
                "seller": obj.get("seller") or obj.get("orgName"),
            }
            # downloadUrl 形如 .../sopinv/invoicePreview.html?apply_serial_no=<id>saas
            dl_url = obj.get("downloadUrl") or ""
            m = re.search(r"apply_serial_no=([A-Za-z0-9]+)", dl_url)
            if m:
                apply_serial_no = m.group(1)
    except Exception as e:
        print(f"  ⚠️  财云通详情 API 失败（不影响 PDF 下载）: {e}")

    pdf_url = f"{parsed.scheme}://{parsed.netloc}/sopinv/invoice/out/fusion/templateDownload/1/{apply_serial_no}"
    pdf = _try_download_pdf(pdf_url, session, referer=final_url)
    if not pdf:
        return None
    return pdf, extra


# host 关键字 → 适配器
LANDING_ADAPTERS = [
    ("nnfp.jss.com.cn", _adapter_nuonuo),
    ("gateway.bigfintax.com", _adapter_bigfintax),
    # 后续新站点只在此处加一行
]


def _match_landing_adapter(url: str):
    low = url.lower()
    for host_kw, adapter in LANDING_ADAPTERS:
        if host_kw in low:
            return adapter
    return None


def _resolve_pdf_from_landing_page(
    url: str, session: requests.Session
) -> tuple[bytes, dict] | None:
    """
    访问中转页，按以下顺序尝试：
    1) 站点适配器（命中 LANDING_ADAPTERS）
    2) 通用 HTML/JS 探测：找 .pdf 直链
    3) 通用 API 探测：含 download/preview/invoice 等关键字的 URL
    4) 兜底：附加 ?download=1
    返回 (pdf_bytes, extra_fields_dict)，extra_fields 可能为空 dict。
    """
    r = _fetch(url, session)
    if r is None or r.status_code != 200:
        return None

    # 有些跳转页直接 302 到 PDF
    if _looks_like_pdf(r.content):
        return r.content, {}

    final_url = r.url

    # 1) host 适配器
    adapter = _match_landing_adapter(final_url)
    if adapter:
        result = adapter(final_url, session)
        if result:
            return result

    text = r.text or ""

    # 1) 找页面里出现的 .pdf 直链（href 或 JS 字符串里的）
    candidates = []
    for m in PDF_DIRECT_RE.finditer(text):
        candidates.append(m.group(1))
    for m in PDF_URL_IN_TEXT_RE.finditer(text):
        candidates.append(m.group(0))
    # JS 里常见 'xxx.pdf' 单/双引号包裹
    for m in re.finditer(r"""['"]([^'"\s]+\.pdf(?:\?[^'"\s]*)?)['"]""", text, re.IGNORECASE):
        candidates.append(m.group(1))

    seen = set()
    for cand in candidates:
        full = urljoin(r.url, cand)
        if full in seen:
            continue
        seen.add(full)
        pdf = _try_download_pdf(full, session, referer=r.url)
        if pdf:
            return pdf, {}

    # 2) 找下载/预览类 API（含 download/preview/invoice/file 关键词的 http 链接）
    api_candidates = []
    for m in re.finditer(
        r"""['"]?(https?://[^\s'"<>]*(?:download|preview|invoice|getFile|fileDownload|getPdf)[^\s'"<>]*)['"]?""",
        text,
        re.IGNORECASE,
    ):
        api_candidates.append(m.group(1))
    for cand in api_candidates:
        if cand in seen:
            continue
        seen.add(cand)
        pdf = _try_download_pdf(cand, session, referer=r.url)
        if pdf:
            return pdf, {}

    # 3) 兜底：尝试常见的 ?download=1 / ?downloadType=1 参数
    parsed = urlparse(r.url)
    if parsed.scheme and parsed.netloc:
        for suffix in ("?download=1", "?downloadType=1", "&download=1"):
            try_url = r.url + (suffix if "?" not in r.url or suffix.startswith("&") else suffix)
            pdf = _try_download_pdf(try_url, session, referer=r.url)
            if pdf:
                return pdf, {}

    return None


def collect_pdf_from_mail(msg, html: str) -> tuple[bytes, str | None, dict] | None:
    """
    按优先级取 PDF。返回 (pdf_bytes, dl_filename, extra_fields) 或 None。
    extra_fields 来自 landing 适配器（如有），可补全 name/invoice_no/seller。
    1) 邮件附件 .pdf
    2) 正文 <a href> 中的 *.pdf 直链
    3) 正文中所有 http(s) 链接 → 中转页探测
    """
    # 1) 附件
    for part in msg.walk():
        disp = str(part.get("Content-Disposition", ""))
        if "attachment" not in disp.lower():
            continue
        filename = part.get_filename()
        if filename:
            filename = decode_mime_words(filename)
        if not filename or not filename.lower().endswith(".pdf"):
            continue
        data = part.get_payload(decode=True)
        if data and _looks_like_pdf(data):
            return data, filename, {}

    if not html:
        return None

    session = requests.Session()

    # 2) 正文 .pdf 直链
    pdf_urls: list[str] = []
    for m in PDF_DIRECT_RE.finditer(html):
        pdf_urls.append(m.group(1))
    for m in PDF_URL_IN_TEXT_RE.finditer(html):
        pdf_urls.append(m.group(0))
    seen = set()
    for url in pdf_urls:
        if url in seen:
            continue
        seen.add(url)
        pdf = _try_download_pdf(url, session)
        if pdf:
            return pdf, None, {}

    # 3) 正文里其它 http 链接 → 中转页探测
    other_urls: list[str] = []
    for m in ANY_HREF_RE.finditer(html):
        u = m.group(1)
        if u in seen:
            continue
        # 跳过明显的非发票链接
        if any(skip in u.lower() for skip in [
            "unsubscribe", "mailto:", "weibo", "weixin", "qq.com/cgi", "/help",
        ]):
            continue
        other_urls.append(u)

    for url in other_urls:
        if url in seen:
            continue
        seen.add(url)
        result = _resolve_pdf_from_landing_page(url, session)
        if result:
            pdf_bytes, extra = result
            return pdf_bytes, None, extra

    return None


# ── 主处理流程 ────────────────────────────────────────
def process_invoice_mail(msg, subject: str, sender: str) -> bool:
    """
    通用发票邮件处理：
    主题包含"发票"即调用本函数；本函数自行判断能否成功取到 PDF。
    成功保存返回 True，否则 False（不影响其它邮件）。
    """
    print(f"── 邮件: {subject}")

    html = get_email_html_body(msg)
    text = get_email_text_body(msg)
    fields = extract_invoice_fields(subject, html, text)

    # 取 PDF
    got = collect_pdf_from_mail(msg, html)
    if not got:
        print("  ❌ 未能从邮件中提取到 PDF")
        return False
    pdf_bytes, dl_filename, extra = got
    print(f"  📥 已下载 PDF: {len(pdf_bytes)} bytes" + (f" (filename={dl_filename})" if dl_filename else ""))

    # 字段补全：若主题/正文没拿到，尝试用 landing 适配器结果 / 下载文件名 / PDF 文本兜底
    name = fields.get("name")
    invoice_no = fields.get("invoice_no")
    seller = fields.get("seller")

    # landing 适配器返回的字段优先级高于正文（因为来自结构化 API）
    if extra:
        name = extra.get("name") or name
        invoice_no = extra.get("invoice_no") or invoice_no
        seller = extra.get("seller") or seller

    # 下载文件名形如 "电子发票（普通发票）_{发票号}_{销方}_{购方}_{日期}.pdf"
    if dl_filename:
        stem = Path(dl_filename).stem
        parts = stem.split("_")
        if len(parts) >= 4:
            invoice_no = invoice_no or parts[-4]
            seller = seller or parts[-3]
            name = name or parts[-2]

    # 用 PDF 文本兜底
    if not invoice_no or not name or not seller:
        tmp_dir = BASE_DIR / ".tmp"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        tmp_pdf = tmp_dir / f"_probe_{os.getpid()}.pdf"
        try:
            tmp_pdf.write_bytes(pdf_bytes)
            try:
                pdf_text = extract_pdf_text(tmp_pdf)
            except Exception:
                pdf_text = ""
            if pdf_text:
                pdf_text_norm = pdf_text.replace("：", ":")
                if not invoice_no:
                    m = re.search(r"\b(\d{12,20})\b", pdf_text_norm)
                    if m:
                        invoice_no = m.group(1)
                if not seller:
                    m = re.search(r"([\u4e00-\u9fa5]+市[\u4e00-\u9fa5]+?有限公司)", pdf_text_norm)
                    if not m:
                        m = re.search(r"([\u4e00-\u9fa5]{2,}?有限公司)", pdf_text_norm)
                    if m:
                        seller = m.group(1)
                if not name and seller:
                    m = re.search(
                        r"([\u4e00-\u9fa5]{2,4})\s+" + re.escape(seller),
                        pdf_text_norm,
                    )
                    if m:
                        name = m.group(1)
        finally:
            tmp_pdf.unlink(missing_ok=True)

    # 兜底默认值
    if not name:
        name = "未知"
        print("  ⚠️  未识别到姓名/抬头")
    if not invoice_no:
        print("  ❌ 未识别到发票号，跳过保存")
        return False
    if seller:
        region, _ = extract_region(seller)
    else:
        region = "未知地区"

    print(f"   姓名: {name} | 地区: {region} | 发票号: {invoice_no}")

    # 保存
    save_dir = BASE_DIR / f"{name}-{region}"
    save_dir.mkdir(parents=True, exist_ok=True)
    jpg_path = save_dir / f"{invoice_no}.jpg"

    if jpg_path.exists():
        print(f"  ⏭️  已存在，跳过: {jpg_path.name}")
        return True

    pdf_path = save_dir / f"{invoice_no}.pdf"
    try:
        pdf_path.write_bytes(pdf_bytes)
        try:
            pdf_to_jpg(pdf_path, jpg_path)
        except Exception as e:
            print(f"  ❌ PDF 转换失败: {e}")
            return False
    finally:
        pdf_path.unlink(missing_ok=True)

    return True


# ── IMAP / 邮箱处理 ──────────────────────────────────
def connect_imap(server: str, user: str, password: str) -> imaplib.IMAP4_SSL | None:
    """连接并登录 IMAP 服务器，失败返回 None"""
    print(f"📧 正在连接邮箱 {user} ({server}) ...")
    try:
        imaplib.Commands["ID"] = ("AUTH",)
        mail = imaplib.IMAP4_SSL(server, IMAP_PORT)
        mail.login(user, password)
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
    处理单个邮箱：主题含"发票"即进入通用流程。
    返回 (已处理数, 跳过数)
    """
    processed_count = 0
    skipped_count = 0
    delete_ids = []

    status, data = mail.select("INBOX")
    if status != "OK":
        print(f"  ❌ 无法打开收件箱: {data}")
        return (0, 0)
    total = data[0].decode()
    print(f"  📂 收件箱打开成功，共 {total} 封邮件")

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

        # 唯一入口规则：主题含"发票"
        if "发票" not in subject:
            skipped_count += 1
            continue

        ok = False
        try:
            ok = process_invoice_mail(msg, subject, sender)
        except Exception as e:
            print(f"  ❌ 处理异常: {e}")
        print()

        if ok:
            processed_count += 1
            delete_ids.append(msg_uid)
        else:
            skipped_count += 1

    # 把已处理成功的发票邮件移到回收站
    if delete_ids:
        trash_folder = find_trash_folder(mail)
        if trash_folder:
            moved = move_to_trash(mail, delete_ids, trash_folder)
            print(f"  🗑️  已将 {moved} 封发票邮件移动到回收站: {trash_folder}")
        else:
            print(f"  ⚠️  未找到回收站文件夹，邮件未删除")

    return (processed_count, skipped_count)


# ── 回收站迁移 ────────────────────────────────────────
TRASH_CANDIDATES = [
    "&XfJT0ZAB-",           # 163 "已删除"
    "Deleted Messages",     # QQ 邮箱
    "&XfJSIJZk-",
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
        m = re.search(r'"([^"]+)"\s*$', line)
        if m:
            folder_names.append(m.group(1))
        else:
            parts = line.rsplit(" ", 1)
            if len(parts) == 2:
                folder_names.append(parts[1].strip())

    for candidate in TRASH_CANDIDATES:
        for name in folder_names:
            if name == candidate:
                return name

    keywords = ["trash", "deleted", "已删除", "回收站", "垃圾箱"]
    for name in folder_names:
        low = name.lower()
        if any(k in low or k in name for k in keywords):
            return name

    return None


def move_to_trash(mail: imaplib.IMAP4_SSL, uids: list, trash_folder: str) -> int:
    """
    将邮件移动到回收站。
    优先使用 IMAP MOVE（RFC 6851），不支持则回退到 COPY + DELETE + EXPUNGE。
    """
    moved = 0
    quoted_folder = f'"{trash_folder}"' if " " in trash_folder else trash_folder
    for uid in uids:
        try:
            status, _ = mail.uid("MOVE", uid, quoted_folder)
            if status == "OK":
                moved += 1
                continue
        except Exception:
            pass
        try:
            status, _ = mail.uid("COPY", uid, quoted_folder)
            if status == "OK":
                mail.uid("STORE", uid, "+FLAGS", "\\Deleted")
                moved += 1
        except Exception as e:
            print(f"  ⚠️  移动邮件 UID={uid.decode() if isinstance(uid, bytes) else uid} 失败: {e}")

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
