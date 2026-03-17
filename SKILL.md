# ExpenseMail2Word

**从邮件自动处理报销PDF并生成汇总文档**

## 概述

这个Skill帮助用户从邮箱中自动搜索、下载和处理报销相关的PDF附件（如12306火车票、滴滴出行、高德打车发票等），提取关键信息，重命名文件，并生成包含汇总表格和原图的Word文档。

## 适用场景

- 定期处理12306火车票报销凭证
- 处理滴滴出行、高德打车等网约车发票和行程单
- 整理其他类型的电子发票（机票、酒店等，后续扩展）
- 自动化报销凭证整理和归档

## 核心功能

1. **邮件搜索** - 按发件人、主题、日期范围搜索邮件（兼容QQ邮箱IMAP）
2. **附件下载** - 自动下载PDF和ZIP附件，自动解压ZIP中的PDF，支持MIME编码文件名
3. **PDF处理** - 提取信息、重命名、转换为图片
4. **行程单配对** - 滴滴/高德行程单与发票按金额自动配对，发票继承行程单的起止地点信息
5. **图片裁剪** - 滴滴行程单裁剪广告头（检测剪切线）、高德行程单裁剪横幅和空白
6. **文档生成** - 创建包含汇总表格和原图的Word文档，配对凭证紧凑排列
7. **文件管理** - 自动整理和归档处理后的文件

## 快速开始

### 1. 安装依赖

```bash
pip install pdfplumber PyMuPDF python-docx imapclient lxml
```

### 2. 配置邮箱

创建配置文件 `config.json`：

```json
{
  "email": {
    "server": "imap.qq.com",
    "port": 993,
    "username": "your-email@qq.com",
    "password": "your-auth-code",
    "use_ssl": true
  },
  "search": {
    "senders": ["didifapiao@mailgate.xiaojukeji.com", "12306@rails.com.cn", "itinerary@ridesharing.amap.com"],
    "date_from": "2025-10-01",
    "date_to": "2026-03-16"
  },
  "processing": {
    "output_dir": "./processed",
    "rename_format": "{date}_{type}_{start}_{end}_{doc_type}.pdf",
    "image_dpi": 150
  }
}
```

### 3. 运行处理

```bash
python scripts/process_reimbursement.py --config config.json
```

### 跨平台说明

本 Skill 支持 **Windows、macOS、Linux**。所有路径在脚本内均通过 `pathlib.Path` 处理，配置文件中的路径可使用正斜杠 `/` 或反斜杠 `\`（Windows）。在 Windows 上建议在项目目录下打开终端（PowerShell 或 CMD）执行上述命令；在 macOS/Linux 下使用 Bash 即可。

## 详细说明

### 邮件搜索配置

支持以下搜索条件：
- `senders`: 发件人列表（支持多个）
- `subjects`: 主题关键词列表
- `date_from`: 开始日期（YYYY-MM-DD）
- `date_to`: 结束日期（YYYY-MM-DD）
- `has_attachments`: 是否只搜索有附件的邮件（检测 `.pdf` 和 `.zip` 附件）

#### QQ邮箱IMAP兼容性
QQ邮箱的IMAP服务器对 `OR` 和组合搜索条件支持有限，脚本采用以下策略绕过：
- 日期+主题条件构建 `base_criteria`，对每个 sender 单独执行 `FROM` 搜索
- 在 Python 端用集合运算（交集/并集）合并多个搜索结果
- 附件检测使用 MIME header 解码（`email.header.decode_header`）处理编码文件名
- ZIP 附件自动解压提取内含的 PDF 文件

### PDF处理规则

#### 12306火车票
- 重命名格式：`YYYYMMDD_12306_起点-终点_车次.pdf`
- 提取信息：日期、时间、起点站、终点站、车次、金额
- 提取方式：正则匹配（站名、车次格式 `G/D/C/T/K/Z + 数字`、日期 `YYYY年M月D日 HH:MM开`、票价 `￥金额\n票价`）

#### 滴滴出行
- 重命名格式：`YYYYMMDD_滴滴_起点_终点_行程单.pdf` / `YYYYMMDD_滴滴_起点_终点_发票.pdf`
- 起点/终点截取前10个字符
- **行程单提取**：使用 `pdfplumber.extract_tables()` 解析行程表格，提取日期、起止地点、金额
- **发票提取**：正则匹配开票日期和小写金额
- **自动配对**：行程单与发票按金额匹配（误差<0.01），发票继承行程单的日期、起点、终点
- 配对后的行程单和发票共享相同的文件名前缀

#### 滴滴行程单图片裁剪
- 自动检测PDF中的剪切线（横跨页面宽度>70%、高度<20pt的扁平图片）
- 裁剪掉剪切线以上的广告区域（滴滴Logo、优惠广告）
- 裁剪掉内容以下的空白和页脚（"页码：x/x"）

#### 高德打车（Amap）
- 发件人：`itinerary@ridesharing.amap.com`
- 重命名格式：`YYYYMMDD_高德_起点_终点_行程单.pdf` / `YYYYMMDD_高德_起点_终点_发票.pdf`
- 起点/终点截取前10个字符
- **行程单提取**：使用 pdfplumber 表格提取模式
  - 先通过 `extract_words()` 获取表头（序号/服务商/车型/上车时间/城市/起点/终点/金额）的精确 x 坐标
  - 计算相邻列中点作为 `explicit_vertical_lines`
  - 调用 `extract_tables(table_settings={"vertical_strategy": "explicit", ...})` 精确分列
  - 从数据行中按列名索引提取起点、终点
  - 日期从 `行程时间：YYYY-MM-DD` 正则提取，金额从 `合计XX元` 正则提取
- **发票提取**：pdfplumber 可能因 ¥ 字符编码问题失败，回退使用 fitz（PyMuPDF）全文提取
  - 日期从 `开票日期：YYYY年M月D日` 正则提取
  - 金额从 `（小写）¥XX.XX` 提取；若 fitz 文本中小写标记与金额被大写文字隔开，回退到 `[圆元角分整正]\s*¥XX.XX` 模式
- **自动配对**：同滴滴逻辑，行程单与发票按金额匹配，发票继承位置信息
- **行程单图片裁剪**：检测 "高德地图"/"AMAP" 文字块定位内容起始，跳过页码定位内容结束，去除顶部广告横幅和底部空白
- **检测防冲突**：高德文档可能含 "快车" 等滴滴关键词（如"特惠快车"），通过优先检查 "高德"/"AMAP" 关键词排除误判

### 输出文档

生成文件名格式：`报销开始日期_报销结束日期_出差地点_报销单生成日期.docx`
- 出差地点从12306车票的起终站自动提取城市名（去掉"南/北/西/东/虹桥/站"后缀）

文档内标题 = 文件名去掉最后的生成日期部分

生成的Word文档包含：
1. **汇总表格** - 所有报销条目按时间排列，表头：序号、条目(12306|滴滴|高德)、时间、起点、终点、金额
   - 已配对的发票不单独出行，金额不重复计算
   - 表格列宽自适应内容
   - 末行显示"合计"金额
2. **统计信息** - 总记录数、总金额
3. **PDF原图** - 所有PDF转换为图片插入文档
   - 配对的行程单和发票图片紧邻排列，使用 `keep_with_next` 尽量放在同一页
   - 滴滴/高德行程单图片已裁剪广告和空白
4. **页码** - 页脚居中显示"第 X / Y 页"

文档格式要求：
- 紧凑布局，去掉不必要的换行和图片标题
- 页边距 1 英寸
- 标题 16pt 居中加粗

## 脚本说明

### 主要脚本

| 脚本 | 功能 | 说明 |
|------|------|------|
| `process_reimbursement.py` | 完整流程 | 一键执行完整处理流程（推荐使用） |
| `search_emails.py` | 搜索邮件 | IMAP搜索，兼容QQ邮箱，单独FROM搜索+集合运算 |
| `download_attachments.py` | 下载附件 | 下载PDF/ZIP附件，MIME解码文件名，自动解压ZIP |
| `extract_pdf_info.py` | 提取PDF信息 | 12306正则、滴滴表格/正则、高德表格(explicit_vertical_lines)/正则/fitz回退 |
| `rename_files.py` | 重命名文件 | 按规则重命名，滴滴/高德地点截取10字符 |
| `create_summary_doc.py` | 创建汇总文档 | 生成Word，滴滴/高德行程单裁剪，配对紧凑排列，自适应表格，页码 |

### 工具函数

| 模块 | 功能 |
|------|------|
| `utils.py` | 通用工具函数 |
| `email_utils.py` | 邮件处理工具 |
| `pdf_utils.py` | PDF处理工具 |
| `doc_utils.py` | Word文档工具 |

## 配置示例

### 基本配置

```json
{
  "email": {
    "server": "imap.qq.com",
    "port": 993,
    "username": "your-email@qq.com",
    "password": "your-auth-code",
    "use_ssl": true
  },
  "search": {
    "senders": ["didifapiao@mailgate.xiaojukeji.com"],
    "date_from": "2025-10-05",
    "date_to": "2026-03-16"
  }
}
```

### 高级配置

```json
{
  "email": { ... },
  "search": {
    "senders": ["didifapiao@mailgate.xiaojukeji.com", "12306@rails.com.cn", "itinerary@ridesharing.amap.com"],
    "subjects": ["电子发票", "行程报销单", "火车票"],
    "date_from": "2025-01-01",
    "date_to": "2026-12-31",
    "has_attachments": true
  },
  "processing": {
    "output_dir": "/path/to/output",
    "rename_rules": {
      "12306": "{date}_{time}_{start}-{end}_{train}.pdf",
      "didi": "{date}_滴滴_{start}_{end}_{type}.pdf"
    },
    "image": {
      "dpi": 150,
      "width_inches": 6.0,
      "quality": 90
    }
  },
  "document": {
    "title": "报销凭证汇总",
    "include_table": true,
    "include_images": true,
    "sort_by": "date",
    "compact_layout": true
  }
}
```

## 使用示例

### 示例1：处理滴滴出行发票

将以下内容保存为 `didi_config.json`（可用任意文本编辑器创建，跨平台通用）：

```json
{
  "email": {
    "server": "imap.qq.com",
    "port": 993,
    "username": "your-email@qq.com",
    "password": "your-auth-code"
  },
  "search": {
    "senders": ["didifapiao@mailgate.xiaojukeji.com"],
    "date_from": "2026-01-01"
  }
}
```

然后运行（Windows PowerShell/CMD 或 macOS/Linux 终端均可）：

```bash
python scripts/process_reimbursement.py --config didi_config.json --type didi
```

### 示例2：处理12306火车票

```bash
python scripts/process_reimbursement.py --config config.json --type 12306
```

### 示例3：完整处理流程

推荐使用 `process_reimbursement.py` 一键处理，内部执行以下步骤：

```
步骤1: 搜索邮件（IMAP，兼容QQ邮箱）
步骤2: 下载附件（PDF/ZIP，自动解压）
步骤3: 提取PDF信息（12306正则 + 滴滴表格/正则 + 高德表格/正则/fitz）
步骤3.5: 行程单与发票配对（滴滴+高德，按金额匹配，发票继承行程单位置信息）
步骤4: 重命名文件（配对的行程单和发票共享前缀）
步骤5: 创建汇总文档（裁剪行程单图片、配对紧凑排列、自适应表格、页码）
```

也可以分步执行：

```bash
python scripts/search_emails.py --config config.json --output emails.json
python scripts/download_attachments.py --emails emails.json --output ./attachments
python scripts/extract_pdf_info.py --input ./attachments --output info.json
python scripts/rename_files.py --info info.json --input ./attachments --output ./renamed
python scripts/create_summary_doc.py --info info.json --pdfs ./renamed --output 报销汇总.docx
```

## 文件结构

```
ExpenseMail2Word/
├── SKILL.md                       # 技能主文档 / README
├── requirements.txt               # Python 依赖
├── .gitignore                     # Git 忽略规则
├── examples/                      # 示例配置
│   └── config_basic.json          # 基本配置模板（占位凭据）
├── references/                    # 参考文档
│   ├── file_formats.md            # 文件格式规范
│   └── search_12306_emails.py     # 邮件搜索参考实现
├── scripts/                       # 可执行脚本
│   ├── process_reimbursement.py   # 主处理脚本（一键运行完整流程）
│   ├── search_emails.py           # 邮件搜索（IMAP，QQ兼容）
│   ├── download_attachments.py    # 附件下载（PDF/ZIP，MIME解码）
│   ├── extract_pdf_info.py        # PDF信息提取（12306 + 滴滴 + 高德）
│   ├── rename_files.py            # 文件重命名
│   ├── create_summary_doc.py      # 文档生成（裁剪、配对、表格、页码）
│   └── utils/                     # 工具函数
│       ├── __init__.py
│       ├── email_utils.py
│       ├── pdf_utils.py
│       └── doc_utils.py
└── output/                        # 输出目录（自动创建，已 gitignore）
    ├── renamed_pdfs/
    └── *.docx
```

## 扩展开发

### 添加新的报销类型

1. 在 `pdf_utils.py` 中添加新的PDF解析器
2. 在配置文件中添加对应的重命名规则
3. 更新信息提取逻辑

### 自定义输出格式

修改 `doc_utils.py` 中的文档生成函数，支持：
- 不同的表格样式
- 自定义图片布局
- 多语言支持

## 注意事项

1. **跨平台**：支持 Windows / macOS / Linux，路径与命令行示例均可在各平台使用
2. **邮箱安全**：使用授权码而非密码，定期更新
3. **文件备份**：处理前备份原始文件
4. **错误处理**：脚本包含完整的错误处理和日志记录
5. **性能考虑**：大量PDF处理时注意内存使用

## 故障排除

### 常见问题

1. **邮箱连接失败**
   - 检查服务器地址和端口
   - 确认授权码正确
   - 检查SSL配置

2. **PDF提取失败**
   - 确认PDF格式支持
   - 检查PDF是否加密
   - 尝试不同的PDF解析库

3. **文档生成错误**
   - 检查Word依赖是否正确安装
   - 确认文件权限
   - 检查图片路径

### 日志查看

所有脚本都支持详细日志输出：

```bash
python scripts/process_reimbursement.py --config config.json --verbose --log-file process.log
```

## 版本历史

- v1.0 (2026-03-16): 初始版本，支持12306和滴滴出行
- v1.1 (2026-03-17): 完善邮件搜索（QQ IMAP兼容）、ZIP解压、MIME文件名解码
- v1.2 (2026-03-17): 修正PDF信息提取（12306正则、滴滴表格提取）、修正重命名和表头格式
- v1.3 (2026-03-17): 滴滴行程单+发票自动配对、输出文件名格式化、行程单图片裁剪广告、紧凑布局、表格自适应、页码
- v1.4 (2026-03-17): 新增高德打车支持（`itinerary@ridesharing.amap.com`），表格提取(explicit_vertical_lines)解析行程单、fitz回退提取发票金额、行程单图片裁剪横幅和空白、自动配对、滴滴/高德关键词防冲突
- 计划功能：支持机票、酒店等更多报销类型

## 作者

由OpenClaw Assistant根据实际需求开发。

## 许可证

MIT License
