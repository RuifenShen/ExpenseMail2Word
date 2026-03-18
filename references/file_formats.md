# 文件格式规范

## PDF文件命名规范

### 1. 12306火车票

**格式：** `YYYYMMDD_HHMM_起点-终点_车次.pdf`

**示例：**
- `20260301_1105_北京西-长沙南_G311.pdf`
- `20251105_1600_北京南-上海虹桥_G19.pdf`
- `20251122_2319_长沙南-北京西_D920.pdf`

**字段说明：**
- `YYYYMMDD`: 乘车日期（年月日）
- `HHMM`: 发车时间（时分）
- `起点`: 出发车站（中文）
- `终点`: 到达车站（中文）
- `车次`: 列车车次（如G311、D920）

### 2. 滴滴出行

**行程单格式：** `YYYYMMDD_滴滴_起点_终点_行程单.pdf`

**发票格式：** `YYYYMMDD_滴滴_起点_终点_发票.pdf`

**示例：**
- `20260127_滴滴_长沙南站-西广场-负_芙蓉北路_途虎养车工_行程单.pdf`
- `20260127_滴滴_长沙南站-西广场-负_芙蓉北路_途虎养车工_发票.pdf`
- `20251025_滴滴_莲花池_北京西站-南_田村_中国房子小区-_行程单.pdf`

**字段说明：**
- `YYYYMMDD`: 行程日期（年月日）
- `起点`: 上车地点（前10个字符）
- `终点`: 下车地点（前10个字符）

### 3. 通用规则

1. **字符限制**：只使用字母、数字、中文、下划线、连字符
2. **长度限制**：文件名不超过100字符
3. **编码**：UTF-8编码
4. **特殊字符**：避免使用 `\/:*?"<>|`

## 配置文件格式

### 基础配置（config.json）

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
    "senders": ["didifapiao@mailgate.xiaojukeji.com"],
    "subjects": ["电子发票", "行程报销单"],
    "email_send_date_from": "2025-10-01",
    "email_send_date_to": "2026-03-16",
    "trip_date_from": "",
    "trip_date_to": "",
    "has_attachments": true
  },
  "processing": {
    "output_dir": "./processed",
    "backup_original": true,
    "max_file_size_mb": 10
  }
}
```

### 高级配置

```json
{
  "email": { ... },
  "search": { ... },
  "processing": {
    "output_dir": "./output",
    "rename_rules": {
      "12306": {
        "format": "{date}_{time}_{start}-{end}_{train}.pdf",
        "max_length": 100
      },
      "didi": {
        "format": "{date}_滴滴_{start}_{end}_{type}.pdf",
        "start_length": 10,
        "end_length": 10
      }
    },
    "image_conversion": {
      "dpi": 150,
      "format": "png",
      "quality": 90,
      "max_width_px": 2000
    }
  },
  "document": {
    "title": "报销凭证汇总",
    "author": "自动生成",
    "template": null,
    "sections": ["summary", "images", "statistics"],
    "sort_by": "date_asc",
    "compact": true
  }
}
```

## 数据文件格式

### 1. 邮件搜索结果（emails.json）

```json
{
  "search_criteria": {
    "senders": ["didifapiao@mailgate.xiaojukeji.com"],
    "email_send_date_from": "2025-10-01",
    "email_send_date_to": "2026-03-16"
  },
  "results": [
    {
      "id": "12345",
      "subject": "滴滴出行电子发票及行程报销单",
      "from": "didifapiao@mailgate.xiaojukeji.com",
      "date": "2026-03-16 15:01:00",
      "attachments": [
        {
          "filename": "滴滴出行行程报销单.pdf",
          "size": 69884,
          "type": "application/pdf"
        },
        {
          "filename": "滴滴电子发票.pdf", 
          "size": 74375,
          "type": "application/pdf"
        }
      ]
    }
  ],
  "summary": {
    "total_emails": 32,
    "total_attachments": 43,
    "processing_time": "2026-03-16 16:30:00"
  }
}
```

### 2. PDF信息提取结果（info.json）

```json
{
  "files": [
    {
      "original_name": "滴滴出行行程报销单.pdf",
      "new_name": "20260127_滴滴_长沙南站-西广场-负_芙蓉北路_途虎养车工_行程单.pdf",
      "type": "didi_trip",
      "date": "2026-01-27",
      "start_location": "长沙南站-西广场-负",
      "end_location": "芙蓉北路_途虎养车工",
      "amount": 39.20,
      "distance_km": 20.4,
      "file_size": 69884,
      "extraction_method": "table"
    },
    {
      "original_name": "20260301_1105_北京西-长沙南_G311.pdf",
      "new_name": "20260301_1105_北京西-长沙南_G311.pdf",
      "type": "12306",
      "date": "2026-03-01",
      "time": "11:05",
      "start_location": "北京西",
      "end_location": "长沙南", 
      "train_number": "G311",
      "amount": 778.00,
      "file_size": 125000
    }
  ],
  "statistics": {
    "total_files": 15,
    "total_amount": 2575.16,
    "by_type": {
      "12306": 3,
      "didi": 12
    }
  }
}
```

### 3. 处理报告（report.json）

```json
{
  "process_id": "20260316_1630_12345",
  "start_time": "2026-03-16 16:30:00",
  "end_time": "2026-03-16 16:45:00",
  "config_file": "config.json",
  "steps": [
    {
      "name": "email_search",
      "status": "success",
      "duration_seconds": 12.5,
      "results": {
        "emails_found": 32,
        "attachments_found": 43
      }
    },
    {
      "name": "attachment_download",
      "status": "success", 
      "duration_seconds": 45.2,
      "results": {
        "files_downloaded": 24,
        "total_size_mb": 1.8
      }
    }
  ],
  "summary": {
    "total_time_seconds": 900,
    "success_rate": 100.0,
    "files_processed": 24,
    "errors": []
  }
}
```

## Word文档格式

### 文档结构

1. **封面页**
   - 标题：报销凭证汇总
   - 生成日期
   - 统计信息

2. **汇总表格**
   - 表格样式：网格线
   - 列：序号、报销条目、时间、出发地、目的地、价格
   - 排序：按时间升序
   - 总计行

3. **凭证原图**
   - 按时间顺序排列
   - 每个条目包含：
     - 标题（日期+类型）
     - 行程信息
     - PDF原图（转换的图片）
     - 文件说明

4. **统计页**
   - 分类统计
   - 金额汇总
   - 处理信息

### 样式规范

- **字体**：中文使用宋体，英文使用Times New Roman
- **字号**：正文10.5pt，标题适当加大
- **行距**：1.5倍行距
- **页边距**：上下2.54cm，左右3.17cm
- **图片宽度**：6英寸（约15.24cm）

## 日志文件格式

### 日志级别

```
[YYYY-MM-DD HH:MM:SS] [LEVEL] [MODULE] Message
```

### 示例

```
[2026-03-16 16:30:15] [INFO] [email_search] 开始搜索邮件
[2026-03-16 16:30:18] [INFO] [email_search] 连接到邮箱服务器: imap.qq.com
[2026-03-16 16:30:20] [INFO] [email_search] 登录成功
[2026-03-16 16:30:25] [INFO] [email_search] 找到32封邮件
[2026-03-16 16:30:30] [WARNING] [attachment_download] 跳过非PDF附件: image.jpg
[2026-03-16 16:30:35] [ERROR] [pdf_extract] PDF解析失败: 文件损坏
```

## 扩展文件格式

### 未来支持的报销类型

#### 机票
- 格式：`YYYYMMDD_航班_起点机场-终点机场_航班号.pdf`
- 示例：`20260320_航班_北京首都-上海虹桥_CA1501.pdf`

#### 酒店
- 格式：`YYYYMMDD_酒店_酒店名称_城市.pdf`
- 示例：`20260315_酒店_北京饭店_北京.pdf`

#### 餐饮
- 格式：`YYYYMMDD_餐饮_餐厅名称_金额.pdf`
- 示例：`20260310_餐饮_全聚德_北京_258.00.pdf`

## 验证规则

### 文件名验证

```python
def validate_filename(filename):
    # 长度检查
    if len(filename) > 100:
        return False
    
    # 字符检查
    invalid_chars = r'[\\/*?:"<>|]'
    if re.search(invalid_chars, filename):
        return False
    
    # 扩展名检查
    if not filename.lower().endswith('.pdf'):
        return False
    
    return True
```

### 金额验证

```python
def validate_amount(amount):
    # 必须是数字
    try:
        float(amount)
    except ValueError:
        return False
    
    # 金额范围
    if float(amount) <= 0 or float(amount) > 100000:
        return False
    
    return True
```

## 兼容性说明

### 支持的PDF版本
- PDF 1.4 及以上
- 支持文本提取（非扫描版）
- 支持加密PDF（需要密码）

### 支持的邮箱服务
- QQ邮箱（IMAP/SMTP）
- 163邮箱
- Gmail（需要应用专用密码）
- 企业邮箱（Exchange需额外配置）

### 系统要求
- Python 3.8+
- 内存：至少512MB
- 磁盘空间：根据PDF数量而定