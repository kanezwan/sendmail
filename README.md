# 📧 发票邮件自动处理工具（通用版）

自动从邮箱中提取**主题包含「发票」**的邮件，按通用流程下载 PDF 并转换为 JPG 图片，按 `{姓名}-{地区}/{发票号}.jpg` 分目录保存。

> 重构目标：**新增发票来源不需要改业务代码**。99% 的发票邮件都能被通用流程处理；只有当中转页是 SPA（HTML 里没有 PDF 链接）时，才需要为该 host 加一个 1 函数的"适配器"。

## 通用处理流程

1. **粗筛**：邮件主题包含 `发票` 即视为发票邮件
2. **字段抽取**（按 label 表，从主题 + 邮件正文）：
   - 姓名：`发票抬头` / `购方名称` / `购买方名称` / `购买方` / `抬头`
   - 发票号：`数电号码` / `发票号码` / `发票号`
   - 销方：`销方名称` / `销售方名称` / 主题中【...】方括号
   - 地区：从销方名称按"省/市/区/县/州"切分
3. **PDF 获取**（按优先级，第一个成功就停）：
   1. 邮件**附件** `.pdf`
   2. 正文中的 `.pdf` **直链**
   3. 正文中其它链接 → **进入中转页**：
      - 命中 `LANDING_ADAPTERS` 中的 host → 调用专用适配器
      - 否则通用探测：HTML/JS 里搜 `.pdf` 链接 / 含 `download/preview/invoice` 关键字的 API
4. **PDF 字段兜底**：若上一步未拿全，从下载文件名 + PDF 第一页文本里再抽
5. **PDF → JPG**，保存至 `invoice/{姓名}-{地区}/{发票号}.jpg`
6. 处理成功的邮件**移动到回收站**（可恢复）

## 当前内置的"中转页适配器"

| Host 关键字 | 站点 | 处理方式 |
|---|---|---|
| `nnfp.jss.com.cn` | 诺诺网 | 短链跳转 → POST `/scan2/getIvcDetailShow.do` 获取 `invoiceSimpleVo` → GET `vo.url` |

> **如何新增一个适配器**：在 `main.py` 的 `LANDING_ADAPTERS` 列表里加一行 `(host_关键字, 适配器函数)` 即可，业务流程不变。

## 已验证覆盖的发票来源

| 来源 | 发件人示例 | 命中环节 |
|---|---|---|
| 票通 | 各税务平台 | 附件 `.pdf` |
| 朴朴超市 | `message@pupumall.net` | 正文 `.pdf` 直链 |
| 财云通 | `ysb@szyh.com` | 中转页通用 API 探测（含 `download` 关键字） |
| 诺诺网 | `invoice@info.nuonuo.com` | `nnfp.jss.com.cn` 适配器 |

## 目录结构

```
sendmail/
├── .env                 # 邮箱账号配置（需自行填写）
├── requirements.txt
├── main.py              # 主程序（通用流程 + 适配器列表）
├── doc/                 # 需求文档
└── invoice/             # 输出目录（运行时自动清空 + 重建）
    ├── 周琬-上海/
    │   ├── 26312000002739059056.jpg
    │   └── ...
    └── ...
```

## 安装

```bash
python3 -m venv ~/.venv
source ~/.venv/bin/activate
pip install -r requirements.txt
```

## 配置 `.env`

```env
# 多个账号用 ; 分隔；格式：用户名:授权码:IMAP服务器
EMAIL_ACCOUNTS=user1@163.com:AUTH_CODE_1:imap.163.com;user2@qq.com:AUTH_CODE_2:imap.qq.com
```

| 邮箱 | IMAP 服务器 | 授权码获取 |
|---|---|---|
| 163 | `imap.163.com` | 设置 → POP3/SMTP/IMAP → 开启 IMAP → 授权码 |
| QQ | `imap.qq.com` | 设置 → 账户 → 开启 IMAP → 授权码 |

## 运行

```bash
python main.py
```

## 何时需要改代码

- ✅ **不需要改**：邮件附件带 PDF、正文直接给 `.pdf` 链接、正文给的链接是普通 HTML 页且 PDF 链接出现在 HTML/JS 里
- ⚠️ **需要加 1 个适配器函数**：链接是 SPA 页（HTML 是空壳，PDF 在 JSON API 里），且通用 API 关键字探测拿不到。只在 `LANDING_ADAPTERS` 列表加一行，**不动主流程**。

## 邮件清理策略

成功处理的发票邮件 → **移动到"回收站/已删除"文件夹**（可恢复，非彻底删除）。非发票邮件保持不动。
