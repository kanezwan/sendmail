# 📧 发票邮件自动处理工具

自动从邮箱中提取电子发票邮件，下载 PDF 并转换为 JPG 图片，按 `{姓名}-{地区}` 分目录保存。

## 功能概述

1. **多邮箱支持**：同时处理多个邮箱（163、QQ 等）
2. **多发票类型支持**：
   - **"票通"类型**：附件直接带 PDF
   - **"财云通"(bigfintax) 类型**：邮件正文仅提供下载链接，程序自动调用接口拉取 PDF
   - **"朴朴超市"(pupumall) 类型**：邮件正文含 PDF 直链，下载后从 PDF 内容提取字段
3. **PDF → JPG**：自动将发票 PDF 转换为高清 JPG 图片
4. **智能跳过**：已下载的发票不会重复处理
5. **自动清理**：发票邮件处理完成后自动移动到邮箱回收站（可恢复），非发票邮件保留不动
6. **启动清空**：每次运行会先清空 `invoice/` 目录，保证输出全新

## 目录结构

```
sendinvoice/
├── .env                 # 邮箱账号配置（需自行填写）
├── .gitignore
├── requirements.txt     # Python 依赖
├── main.py              # 主程序
├── doc/                 # 需求文档
└── invoice/             # 发票图片输出目录（运行时自动创建/清空）
    ├── 张三-广州/
    │   ├── 264420000036_92402671.jpg
    │   └── 264420000036_92723731.jpg
    ├── 周琬-北京/
    │   └── 26112000001770567046.jpg
    └── ...
```

## 安装

### 1. 安装 Python 依赖

```bash
python3 -m venv ~/.venv
source ~/.venv/bin/activate
pip install -r requirements.txt
```

依赖包：
- `pypdfium2` — PDF 渲染引擎（纯 Python，无需系统依赖）
- `Pillow` — 图片处理
- `requests` — 调用财云通下载接口

### 2. 配置邮箱账号

编辑项目根目录下的 `.env` 文件：

```env
# 多个账号用分号（;）分隔
# 格式: 用户名:密码(授权码):IMAP服务器
EMAIL_ACCOUNTS=你的邮箱:授权码:imap服务器地址
```

**示例：**

```env
# 单个邮箱
EMAIL_ACCOUNTS=example@163.com:YOUR_AUTH_CODE:imap.163.com

# 多个邮箱
EMAIL_ACCOUNTS=user1@163.com:AUTH_CODE_1:imap.163.com;user2@qq.com:AUTH_CODE_2:imap.qq.com
```

### 3. 获取邮箱授权码

| 邮箱 | IMAP 服务器 | 授权码获取方式 |
|------|------------|--------------|
| 163 邮箱 | `imap.163.com` | 登录 [mail.163.com](https://mail.163.com) → 设置 → POP3/SMTP/IMAP → 开启 IMAP 服务 → 获取授权码 |
| QQ 邮箱 | `imap.qq.com` | 登录 [mail.qq.com](https://mail.qq.com) → 设置 → 账户 → 开启 IMAP/SMTP 服务 → 获取授权码 |

> ⚠️ **注意**：IMAP 登录需要使用**授权码**，不是邮箱登录密码。

## 使用方法

```bash
python main.py
```

程序会自动执行以下流程：

```
清空 invoice/ → 连接邮箱 → 读取收件箱 → 识别发票类型 → 下载/提取 PDF → 转换为 JPG → 按目录保存 → 移动已处理邮件到回收站
```

### 运行示例

```
🧹 已清空目录: /path/to/invoice
📧 正在连接邮箱 example@qq.com (imap.qq.com) ...
  ✅ 登录成功
  📂 收件箱打开成功，共 22 封邮件
  📬 找到 22 封邮件，开始处理...

── 邮件: 您收到一张来自广州xx有限公司的电子发票【发票金额：58.00】
   姓名: 张三 | 地区: 广州
  📥 已下载: 264420000036_92402671_张三.pdf
  ✅ 已转换: 264420000036_92402671.jpg

── 邮件: 【电子发票】您有一张电子发票[发票号码 ：26112000001770567046]
   类型: 财云通(bigfintax) | inv_id=1859362401638842224
   姓名: 周琬 | 地区: 北京 | 发票号: 26112000001770567046
  📥 已下载: 26112000001770567046.pdf (142775 bytes)
  ✅ 已转换: 26112000001770567046.jpg

  🗑️  已将 9 封发票邮件移动到回收站: Deleted Messages
==================================================
✅ 全部完成！共处理 9 封发票邮件，跳过 26 封非发票邮件
📁 发票图片保存在: /path/to/sendinvoice/invoice
```

## 支持的发票类型

### 类型一："票通"（带 PDF 附件）

- **邮件主题**：`您收到一张来自{地区}{公司名称}的电子发票【发票金额：xxx】`
- **PDF 附件名**：`{发票代码}_{发票号码}_{姓名}.pdf`
- **输出**：`invoice/{姓名}-{地区}/{发票代码}_{发票号码}.jpg`

### 类型二："财云通"（bigfintax，邮件仅含下载链接）

- **发件人**：`财云通 <ysb@szyh.com>`
- **邮件主题**：`【电子发票】您有一张电子发票[发票号码 ：xxx]`
- **邮件正文**：HTML 中包含链接 `https://gateway.bigfintax.com/scanning-invoice/checkInvoice?id={inv_id}`
- **处理方式**：
  1. 从正文提取 `inv_id`
  2. 调用 `https://gateway.bigfintax.com/xxApi/api/v2/electronInvoice/invoiceBatchDownload?id={inv_id}&downloadType=1` 直接下载 PDF
  3. 从响应头 `Content-Disposition` 的文件名解析发票号、销方、购方、开票日期
- **输出**：`invoice/{购方姓名}-{销方地区}/{发票号码}.jpg`

### 类型三："朴朴超市"（pupumall，邮件含 PDF 直链）

- **发件人**：`朴朴超市 <message@pupumall.net>`
- **邮件主题**：`朴朴超市-电子发票通知邮件`（完全固定）
- **邮件正文**：HTML 中包含 3 个 `<a href='...'>` 链接（XML / OFD / PDF），PDF 链接形如 `https://finance-files.pupumall.com/INVOICE_REQUEST/{年}/{月}/{日}/{uuid}/{hash}.pdf`
- **处理方式**：
  1. 从正文正则提取 `.pdf` 链接
  2. `requests.get` 直接下载（无需鉴权）
  3. 用 `pypdfium2` 提取 PDF 第一页文本，正则解析：
     - 发票号码（20 位数字）
     - 销售方名称（`X市XXX有限公司`）→ 地区
     - 购买方姓名（销方公司名同行前面的 2-4 字中文）
- **输出**：`invoice/{购买方姓名}-{销方地区}/{发票号码}.jpg`

## 邮件清理策略

处理成功的发票邮件会**移动到邮箱的"已删除/回收站"文件夹**（而非彻底删除），
可在邮箱的已删除邮件中随时恢复。非发票邮件保持原状。
