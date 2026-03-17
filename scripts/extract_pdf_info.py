#!/usr/bin/env python3
"""
提取PDF信息 - 从PDF提取行程信息
"""

import json
import argparse
import os
import sys
import re
from pathlib import Path
import pdfplumber
import fitz  # PyMuPDF
import logging

def setup_logging(verbose=False):
    """设置日志记录"""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )

def extract_text_from_pdf(pdf_path):
    """使用pdfplumber提取PDF文本"""
    text_content = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    text_content += text + "\n"
    except Exception as e:
        logging.warning(f"使用pdfplumber读取PDF失败，尝试PyMuPDF: {str(e)}")

        # 回退到PyMuPDF
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                text_content += page.get_text() + "\n"
            doc.close()
        except Exception as e2:
            logging.error(f"无法读取PDF {pdf_path}: {str(e2)}")
            return ""

    return text_content

def extract_12306_info(pdf_path):
    """提取12306火车票信息"""
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return None
    if not any(kw in text for kw in ["12306", "铁路", "中国铁路", "客票"]):
        return None

    # 实际 PDF 文本格式：
    #   北京西 G311 长沙南
    #   站 站
    #   ...
    #   2026年03月01日 11:05开 12车10A号 二等座
    #   ￥778.00
    #   票价:

    station_match = re.search(
        r'(\S+)\s+((?:G|D|C|T|K|Z)\d+)\s+(\S+)\n站\s+站', text)

    dt_match = re.search(
        r'(\d{4}年\d{1,2}月\d{1,2}日)\s+(\d{2}:\d{2})开', text)

    seat_match = re.search(
        r'(一等座|二等座|商务座|特等座|硬座|软座|硬卧|软卧|高级软卧|动卧)', text)

    price_match = re.search(r'[￥¥](\d+\.\d{2})\n票价', text)

    date_raw = dt_match.group(1) if dt_match else ''
    date_str = date_raw.replace('年', '-').replace('月', '-').replace('日', '') if date_raw else ''

    return {
        'type': '12306',
        'original_filename': os.path.basename(pdf_path),
        'source': '12306',
        'date': date_str,
        'time': dt_match.group(2) if dt_match else '',
        'departure': station_match.group(1) if station_match else '',
        'destination': station_match.group(3) if station_match else '',
        'train_number': station_match.group(2) if station_match else '',
        'seat_class': seat_match.group(0) if seat_match else '',
        'amount': float(price_match.group(1)) if price_match else 0.0,
        'full_text': text[:500]
    }


def _extract_didi_trip(pdf_path, text):
    """提取滴滴行程报销单信息（表格式）"""
    # 行程起止日期
    date_match = re.search(r'行程起止日期[：:]\s*(\d{4}-\d{2}-\d{2})', text)
    date = date_match.group(1) if date_match else ''

    # 合计金额
    amount_match = re.search(r'合计(\d+\.?\d*)元', text)
    amount = float(amount_match.group(1)) if amount_match else 0.0

    # 用 pdfplumber 表格提取获取起点/终点
    start_location, end_location = '', ''
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if row and row[0] and row[0].strip().isdigit():
                            if len(row) >= 6:
                                start_location = (row[4] or '').strip().replace('\n', '')
                                end_location = (row[5] or '').strip().replace('\n', '')
                            break
                    if start_location:
                        break
    except Exception:
        pass

    return {
        'type': 'didi',
        'doc_type': '行程单',
        'original_filename': os.path.basename(pdf_path),
        'source': 'didi',
        'date': date,
        'start_location': start_location,
        'end_location': end_location,
        'amount': amount,
        'full_text': text[:500]
    }


def _extract_didi_invoice(pdf_path, text):
    """提取滴滴电子发票信息"""
    # 开票日期
    date_match = re.search(r'开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)', text)
    date_raw = date_match.group(1) if date_match else ''
    date = date_raw.replace('年', '-').replace('月', '-').replace('日', '') if date_raw else ''

    # 价税合计（小写）¥39.20
    amount_match = re.search(r'[（(]小写[）)][￥¥](\d+\.?\d*)', text)
    amount = float(amount_match.group(1)) if amount_match else 0.0

    return {
        'type': 'didi',
        'doc_type': '发票',
        'original_filename': os.path.basename(pdf_path),
        'source': 'didi',
        'date': date,
        'start_location': '',
        'end_location': '',
        'amount': amount,
        'full_text': text[:500]
    }


def extract_didi_info(pdf_path):
    """提取滴滴出行信息：区分行程单和发票"""
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return None
    # 高德文档可能含 "快车" 等关键词，优先排除
    if any(kw in text for kw in ['高德', 'AMAP', '高德地图']) or '高德' in os.path.basename(pdf_path):
        return None
    if not any(kw in text for kw in ["滴滴", "小桔科技", "行程报销", "快车", "专车", "出租车", "didi"]):
        return None

    if '行程单' in text or 'TRIP TABLE' in text:
        return _extract_didi_trip(pdf_path, text)
    elif '电子发票' in text:
        return _extract_didi_invoice(pdf_path, text)
    else:
        return _extract_didi_trip(pdf_path, text)


def _extract_amap_trip(pdf_path, text):
    """提取高德行程单信息（表格提取模式：根据表头动态计算列分隔线）"""
    date_match = re.search(r'行程时间[：:]\s*(\d{4}-\d{2}-\d{2})', text)
    date = date_match.group(1) if date_match else ''

    amount_match = re.search(r'合计(\d+\.?\d*)元', text)
    amount = float(amount_match.group(1)) if amount_match else 0.0

    start_location, end_location = '', ''
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            page = pdf.pages[0]
            words = page.extract_words()

            col_names = ['序号', '服务商', '车型', '上车时间', '城市', '起点', '终点', '金额']
            col_pos = {}
            for w in words:
                if w['text'] in col_names:
                    col_pos[w['text']] = (w['x0'], w['x1'])

            if len(col_pos) >= len(col_names):
                ordered = [col_pos[c] for c in col_names]
                v_lines = [ordered[0][0] - 5]
                for i in range(len(ordered) - 1):
                    v_lines.append((ordered[i][1] + ordered[i + 1][0]) / 2)
                v_lines.append(ordered[-1][1] + 10)

                tables = page.extract_tables(table_settings={
                    'vertical_strategy': 'explicit',
                    'explicit_vertical_lines': v_lines,
                })
                for table in tables:
                    header = [c.strip() if c else '' for c in table[0]] if table else []
                    start_idx = header.index('起点') if '起点' in header else 5
                    end_idx = header.index('终点') if '终点' in header else 6
                    for row in table[1:]:
                        if row and row[0] and row[0].strip()[:1].isdigit():
                            start_location = (row[start_idx] or '').strip().replace('\n', ' ')
                            end_location = (row[end_idx] or '').strip().replace('\n', ' ')
                            break
                    if start_location:
                        break
    except Exception:
        pass

    return {
        'type': 'amap',
        'doc_type': '行程单',
        'original_filename': os.path.basename(pdf_path),
        'source': 'amap',
        'date': date,
        'start_location': start_location,
        'end_location': end_location,
        'amount': amount,
        'full_text': text[:500]
    }


def _extract_amap_invoice(pdf_path, text):
    """提取高德电子发票信息（pdfplumber 可能因 ¥ 字符失败，回退到 fitz）"""
    if not text or '小写' not in text:
        try:
            doc = fitz.open(str(pdf_path))
            text = ''
            for page in doc:
                text += page.get_text()
            doc.close()
        except Exception:
            pass

    date_match = re.search(r'开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)', text)
    date_raw = date_match.group(1) if date_match else ''
    date = date_raw.replace('年', '-').replace('月', '-').replace('日', '') if date_raw else ''

    amount = 0.0
    m = re.search(r'[（(]小写[）)]\s*[￥¥](\d+\.?\d*)', text)
    if m:
        amount = float(m.group(1))
    else:
        # fitz 文本中（小写）和 ¥ 金额可能被大写金额文字隔开
        m = re.search(r'[圆元角分整正]\s*[￥¥](\d+\.?\d*)', text)
        if m:
            amount = float(m.group(1))

    return {
        'type': 'amap',
        'doc_type': '发票',
        'original_filename': os.path.basename(pdf_path),
        'source': 'amap',
        'date': date,
        'start_location': '',
        'end_location': '',
        'amount': amount,
        'full_text': text[:500]
    }


def extract_amap_info(pdf_path):
    """提取高德出行信息：区分行程单和发票。
    行程单含 '高德'/'AMAP' 关键词；发票可能不含，但文件名含 '高德打车'。
    """
    filename = os.path.basename(pdf_path)
    text = extract_text_from_pdf(pdf_path)

    is_amap = False
    if text and any(kw in text for kw in ['高德', 'AMAP', 'AMAP ITINERARY', '高德地图']):
        is_amap = True
    if '高德' in filename:
        is_amap = True
    if not is_amap:
        return None

    if text and ('行程单' in text or 'AMAP ITINERARY' in text):
        return _extract_amap_trip(pdf_path, text)
    elif '发票' in filename or (text and '电子发票' in text):
        return _extract_amap_invoice(pdf_path, text)
    else:
        return None


def extract_pdf_info(pdf_path):
    """通用PDF信息提取函数"""
    # 尝试提取12306信息
    info = extract_12306_info(pdf_path)
    if info:
        return info

    # 尝试提取滴滴信息
    info = extract_didi_info(pdf_path)
    if info:
        return info

    # 尝试提取高德信息
    info = extract_amap_info(pdf_path)
    if info:
        return info

    # 如果都不是特定类型，返回基本信息
    return {
        'type': 'unknown',
        'original_filename': os.path.basename(pdf_path),
        'source': 'unknown',
        'date': '',
        'time': '',
        'departure': '' if '12306' in pdf_path else '',
        'destination': '' if '12306' in pdf_path else '',
        'start_location': '' if 'didi' in pdf_path else '',
        'end_location': '' if 'didi' in pdf_path else '',
        'train_number': '',
        'seat_class': '',
        'amount': 0.0,
        'full_text': extract_text_from_pdf(pdf_path)[:500] + "..." if len(extract_text_from_pdf(pdf_path)) > 500 else extract_text_from_pdf(pdf_path)
    }

def main():
    parser = argparse.ArgumentParser(description='从PDF提取行程信息')
    parser.add_argument('--input', '-i', required=True, help='输入PDF文件路径或包含PDF的目录')
    parser.add_argument('--output', '-o', required=True, help='输出文件路径（JSON格式）')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细日志')

    args = parser.parse_args()
    setup_logging(args.verbose)

    input_path = Path(args.input)
    pdf_files = []

    # 确定PDF文件列表
    if input_path.is_file() and input_path.suffix.lower() == '.pdf':
        pdf_files = [input_path]
    elif input_path.is_dir():
        pdf_files = list(input_path.rglob('*.pdf'))
    else:
        logging.error(f"输入路径无效: {args.input}")
        sys.exit(1)

    logging.info(f"找到 {len(pdf_files)} 个PDF文件")

    extracted_info = []

    for pdf_path in pdf_files:
        logging.info(f"正在处理: {pdf_path.name}")
        try:
            info = extract_pdf_info(pdf_path)
            if info:
                info['filepath'] = str(pdf_path)
                extracted_info.append(info)
                logging.info(f"成功提取信息: {info.get('departure', '')} -> {info.get('destination', info.get('end_location', ''))} | 金额: {info.get('amount', 0)}元")
        except Exception as e:
            logging.error(f"处理PDF {pdf_path} 时出错: {str(e)}")
            continue

    # 保存结果
    result = {
        'summary': {
            'total_files': len(pdf_files),
            'successful_extractions': len(extracted_info),
            'failed_extractions': len(pdf_files) - len(extracted_info)
        },
        'extractions': extracted_info
    }

    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    logging.info(f"信息提取完成，结果已保存到: {args.output}")
    logging.info(f"成功处理: {len(extracted_info)}/{len(pdf_files)} 个文件")

if __name__ == '__main__':
    main()