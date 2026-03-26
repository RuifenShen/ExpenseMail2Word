# ExpenseMail2Word

**ExpenseMail2Word** 是一个智能 Agent 工程，用于自动从邮箱中提取、归集、整理报销凭证信息，并生成标准化、格式化的 Word 文档，用于报销申报与归档。

**你可以先试用本项目已有脚本快速开始，然后再通过优化SKILL.md来定制开发新功能。**

本项目的已有脚本支持 **12306 火车票**、**滴滴出行**、**高德打车**、**首汽约车** 等网约车发票与行程单，以及通用电子发票（餐饮、机票等）。已支持功能如下：

- **邮件搜索**：按发件人、主题、日期范围搜索（兼容 QQ 邮箱 IMAP）
- **附件下载**：自动下载 PDF/ZIP，解压 ZIP 内 PDF，支持 MIME 编码文件名
- **智能提取**：12306、滴滴、高德、首汽约车等行程单与发票信息解析
- **行程单配对**：行程单与发票按金额自动配对，发票继承行程单位置与日期
- **图片裁剪**：自动裁剪行程单广告、空白区域
- **汇总文档**：生成含表格与原图的 Word 文档，配对凭证紧凑排列

## 快速开始

### 1. 安装依赖

```bash
pip install pdfplumber PyMuPDF python-docx imapclient lxml

或者 

pip install -r requirements.txt
```

### 2. 配置

复制示例配置并填写邮箱信息：

```bash
cp examples/config_basic.json config.json
```

编辑 `config.json`，填入 `username`、`password`（QQ 邮箱需使用授权码），以及 `email_send_date_from`、`email_send_date_to` 等搜索条件。

### 3. 运行

```bash
python scripts/process_expense.py --config config.json --output ./output
```

输出目录下将生成：
- `intermediate/`：中间结果（邮件列表、附件列表、筛选前后提取结果）
- `renamed_pdfs/`：重命名后的 PDF
- `*.docx`：汇总 Word 文档

## 配置示例

```json
{
  "email": {
    "server": "imap.qq.com",
    "port": 993,
    "username": "your-email@qq.com",
    "password": "your-auth-code",
    "use_ssl": true,
    "timeout": 30
  },
  "search": {
    "senders": ["didifapiao@mailgate.xiaojukeji.com", "12306@rails.com.cn", "itinerary@ridesharing.amap.com"],
    "subjects": [["发票"]],
    "email_send_date_from": "2025-12-01",
    "email_send_date_to": "2026-03-01",
    "trip_date_from": "",
    "trip_date_to": "",
    "has_attachments": true
  }
}
```

## 命令行参数

| 参数 | 说明 |
|------|------|
| `--config`, `-c` | 配置文件路径（必填） |
| `--output`, `-o` | 输出目录，默认 `./output` |
| `--type`, `-t` | 仅处理指定类型：`12306`、`didi` 等 |
| `--verbose`, `-v` | 显示详细日志 |

## 平台支持

- 本项目已考虑多平台兼容，支持Windows / macOS / Linux，路径与配置文件使用 `pathlib` 处理
- 但是尽在Linux平台使用过

## 回归测试

已修复的 bug 有对应回归用例，修改代码后可运行：

```bash
pip install pytest
pytest tests/ -v
```

详见 [REGRESSION.md](REGRESSION.md)。

## 详细文档

更多说明（PDF 解析规则、重命名格式、高级配置、扩展开发）见 [SKILL.md](SKILL.md)。

## License

MIT
