# 香港公司审计 · 银行流水进账统计

> 把 DBS 月结单 PDF 丢进来，自动按「月 × 币种」汇总进账流水，一键生成 Excel 报表。用于**审计报价**和**做账对凭证**。

提供两种版本：

| 版本 | 适合谁 | 怎么用 |
|---|---|---|
| 🖱️ **网页版**（`流水统计器.html`） | 会计师 / 客户 / 非技术人员 | 双击 html 文件 → 浏览器打开 → 拖文件进去 |
| ⌨️ **命令行版**（`bank_statement_analyzer.py`） | 开发者 / 批量处理 | `python3 bank_statement_analyzer.py <目录>` |

两个版本解析逻辑完全一致，数字一模一样。

---

## 🖱️ 网页版（推荐）

**特点**：纯静态 HTML 单文件，**所有 PDF 在本机浏览器内解析**、不联网上传、断网可用。

### 怎么启动
- macOS: 双击「发行包/一键打开.command」或「发行包/流水统计器.html」
- Windows: 双击「发行包/流水统计器.html」（Chrome/Edge 打开）
- 客户分发: 把 `香港审计-流水统计器 v1.0.zip` 发给客户，客户解压双击即可

### 功能
- 支持直接拖入 **ZIP 压缩包**、**整个文件夹**、**多个 PDF** 一起
- 非月结单（Word / Excel / 图片）自动跳过
- 实时显示处理过程树状图，每个文件看到独立状态（等待 / 读取中 / 完成 / 跳过）
- 完成后生成同 CLI 版一致的 4-sheet Excel，本地浏览器直接下载

### 支持浏览器
Chrome / Edge / Safari 现代版本均可。首次打开需联网（从 CDN 下载 pdf.js / SheetJS / JSZip 三个开源库），之后可断网使用。

---

## ⌨️ 命令行版

### 安装

```bash
pip3 install pdfplumber openpyxl
```

### 最常用

```bash
python3 bank_statement_analyzer.py "/path/to/月结单文件夹"
```

### 完整参数

```bash
python3 bank_statement_analyzer.py <pdf目录> \
    -o "进账流水统计-客户名.xlsx" \
    --company "Miheng Trading Limited" \
    --start 2024-07-23 \
    --end 2025-12-31 \
    --rates USD=7.78,EUR=8.85,JPY=0.04987,SGD=5.416,CNY=1.25
```

| 参数 | 说明 |
|---|---|
| `input_dir` | 放 PDF 的文件夹，递归扫描 |
| `-o / --output` | 输出 xlsx 路径，默认写在 `input_dir` 的同级目录 |
| `--company` | 公司名。缺省自动从 PDF 首页识别（抓 `XXX LIMITED/LTD` 行） |
| `--start` / `--end` | 审计账期起止（`YYYY-MM-DD`），不填则按 PDF 覆盖区间 |
| `--rates` | 折算成 HKD 的汇率，格式 `CCY=rate,CCY=rate` |

默认汇率（→HKD）：`USD=7.78, EUR=8.85, JPY=0.04987, SGD=5.416, CNY=1.25, GBP=10.2, AUD=5.1`

---

## 📊 输出 · Excel 4 sheets

| Sheet | 用途 |
|---|---|
| **进账汇总** | 月 × 币种的进账金额矩阵 + 合计行 + 汇率行 —— **审计报价**按这个算 |
| **进账明细** | 每笔进账的日期、账号、币种、金额、交易后余额、对手方、源文件 —— **做账对凭证** |
| **全部交易** | 包含支出，用于与银行余额做 Completeness Test |
| **源文件** | 每张 PDF 覆盖哪几个月、抓到几笔进 / 出 —— 核查有无缺月 |

---

## 🧠 识别逻辑（供审计师核对）

判断「进账 / 支出」用**余额差法**：每笔交易行后都有「交易后余额」，拿本行 − 上一行：

- `delta > 0` → 进账（Credit）
- `delta < 0` → 支出（Debit）

比只看摘要里 `REMITTANCE IN` 关键字更稳 —— 利息存入、利息税、手续费等边角情况都能正确归类。

**交易行识别**：以 `DD-MM-YYYY` 开头、末尾带 `金额 + 余额` 两个数字的行；下一行若不以日期开头，视为上一笔的摘要续行。

---

## 🏦 目前支持的银行

- ✅ **DBS 星展银行**（中国 / 香港）综合月结单 Combined Statement

### 接入其他银行的做法

在 `bank_statement_analyzer.py` 里：

1. 加一套新的 `_SECTION_RE`（识别账户×币种行）
2. 确认交易行格式（日期位置、金额 / 余额顺序）
3. 在 `parse_pdf` 里按 PDF 首页关键字分派到对应 parser

代码里解析和导出已解耦，新增一家银行 ≈ 新增一个 `parse_xxx_pdf()` 函数。

---

## 🔒 数据安全

- 网页版：所有 PDF 只在你本机浏览器里解析，**不会发送到任何服务器**；关掉页面即清除内存
- 命令行版：只读 PDF，只写入 `-o` 指定的本地 xlsx
- 仓库本身**不含**任何真实客户数据；调试用的样本请放入被 `.gitignore` 过滤的目录

---

## 📁 仓库结构

```
.
├── 流水统计器.html              ← 网页版（主产品）
├── bank_statement_analyzer.py  ← 命令行版
├── 一键生成.command             ← macOS 双击启动 CLI
├── 示例输出-Miheng.xlsx         ← 参考：真实数据跑出来的 Excel
├── 发行包/                      ← 发客户的打包版（html + 说明 + 启动器）
│   ├── 流水统计器.html
│   ├── 使用说明.txt
│   └── 一键打开.command
└── README.md
```

---

## 📝 许可

为内部使用定制。如想改造或二次分发，请联系作者。
