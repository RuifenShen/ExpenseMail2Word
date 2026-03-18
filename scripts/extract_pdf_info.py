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
    """提取滴滴行程报销单信息（表格式），支持多行程返回列表。
    表格列可能为：序号、车型、上车时间、城市、起点、终点、里程、金额 等。
    """
    fallback_date_match = re.search(r'行程起止日期[：:]\s*(\d{4}-\d{2}-\d{2})', text)
    fallback_date = fallback_date_match.group(1) if fallback_date_match else ''

    total_amount_match = re.search(r'合计(\d+\.?\d*)元', text)
    total_amount = float(total_amount_match.group(1)) if total_amount_match else 0.0

    trips = []
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    header = [c.strip() if c else '' for c in table[0]]
                    start_idx = next((i for i, h in enumerate(header) if '起点' in h), 4)
                    end_idx = next((i for i, h in enumerate(header) if '终点' in h), 5)
                    time_idx = next((i for i, h in enumerate(header) if '时间' in h or '用车' in h), 2)
                    amt_idx = next((i for i, h in enumerate(header) if '金额' in h and '里程' not in h), 6)
                    if amt_idx >= len(header):
                        amt_idx = 7 if len(header) > 7 else 6

                    for row in table[1:]:
                        if not (row and row[0] and str(row[0]).strip().isdigit()):
                            continue
                        if len(row) <= max(start_idx, end_idx):
                            continue

                        start = (row[start_idx] or '').strip().replace('\n', '')
                        end_loc = (row[end_idx] or '').strip().replace('\n', '')

                        row_date = fallback_date
                        time_cell = (row[time_idx] if time_idx < len(row) else '') or ''
                        time_cell = str(time_cell).strip()
                        dm = re.search(r'(\d{4})[-.年/](\d{1,2})[-.月/](\d{1,2})', time_cell)
                        if dm:
                            row_date = f"{dm.group(1)}-{int(dm.group(2)):02d}-{int(dm.group(3)):02d}"
                        elif fallback_date and re.search(r'^(\d{1,2})[-.月/](\d{1,2})', time_cell):
                            md = re.search(r'^(\d{1,2})[-.月/](\d{1,2})', time_cell)
                            year = fallback_date[:4] if len(fallback_date) >= 4 else ''
                            if year:
                                row_date = f"{year}-{int(md.group(1)):02d}-{int(md.group(2)):02d}"

                        row_amount = 0.0
                        for amt_col in [amt_idx, 7, 6]:
                            if amt_col < len(row) and row[amt_col]:
                                am = re.search(r'(\d+\.?\d*)', str(row[amt_col]).strip())
                                if am:
                                    row_amount = float(am.group(1))
                                    break

                        trips.append({
                            'type': 'didi',
                            'doc_type': '行程单',
                            'original_filename': os.path.basename(pdf_path),
                            'source': 'didi',
                            'date': row_date,
                            'start_location': start,
                            'end_location': end_loc,
                            'amount': row_amount,
                            'full_text': text[:500]
                        })
    except Exception:
        pass

    if not trips:
        return [{
            'type': 'didi',
            'doc_type': '行程单',
            'original_filename': os.path.basename(pdf_path),
            'source': 'didi',
            'date': fallback_date,
            'start_location': '',
            'end_location': '',
            'amount': total_amount,
            'full_text': text[:500]
        }]

    if len(trips) == 1 and trips[0]['amount'] == 0 and total_amount > 0:
        trips[0]['amount'] = total_amount

    return trips


def _extract_didi_invoice(pdf_path, text):
    """提取滴滴电子发票信息。优先使用行程日期（用于 trip_date 筛选），无则用开票日期。"""
    date = ''
    # 行程日期/行程起止日期/行程时间/服务日期/用车日期（报销筛选应以此为准，与邮件发送日期、开票日期区分）
    for pat in [r'行程(?:起止)?日期[：:]\s*(\d{4}-\d{2}-\d{2})', r'行程时间[：:]\s*(\d{4}-\d{2}-\d{2})',
                r'服务日期[：:]\s*(\d{4}-\d{2}-\d{2})', r'用车日期[：:]\s*(\d{4}-\d{2}-\d{2})']:
        trip_dm = re.search(pat, text)
        if trip_dm:
            date = trip_dm.group(1)
            break
    if not date:
        trip_dm = re.search(r'(?:服务|用车)日期[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日', text)
        if trip_dm:
            date = f"{trip_dm.group(1)}-{int(trip_dm.group(2)):02d}-{int(trip_dm.group(3)):02d}"
    if not date:
        trip_dm = re.search(r'行程(?:起止)?日期[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日', text)
        if trip_dm:
            date = f"{trip_dm.group(1)}-{int(trip_dm.group(2)):02d}-{int(trip_dm.group(3)):02d}"
    if not date:
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
    """提取滴滴出行信息：区分行程单和发票。返回列表或 None。"""
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
        return [_extract_didi_invoice(pdf_path, text)]
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
    """提取高德电子发票信息。优先使用行程日期（用于 trip_date 筛选），无则用开票日期。"""
    if not text or '小写' not in text:
        try:
            doc = fitz.open(str(pdf_path))
            text = ''
            for page in doc:
                text += page.get_text()
            doc.close()
        except Exception:
            pass

    date = ''
    for pat in [r'行程(?:起止)?日期[：:]\s*(\d{4}-\d{2}-\d{2})', r'行程时间[：:]\s*(\d{4}-\d{2}-\d{2})',
                r'服务日期[：:]\s*(\d{4}-\d{2}-\d{2})', r'用车日期[：:]\s*(\d{4}-\d{2}-\d{2})']:
        trip_dm = re.search(pat, text)
        if trip_dm:
            date = trip_dm.group(1)
            break
    if not date:
        trip_dm = re.search(r'(?:行程(?:起止)?日期|服务日期|用车日期)[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日', text)
        if trip_dm:
            date = f"{trip_dm.group(1)}-{int(trip_dm.group(2)):02d}-{int(trip_dm.group(3)):02d}"
    if not date:
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


def extract_dining_info(pdf_path):
    """从餐饮发票 PDF 中提取开票日期、销售方名称和金额（价税合计）

    Returns:
        dict: {'date': 'YYYY-MM-DD', 'amount': float, 'seller_name': str}
    """
    text = extract_text_from_pdf(pdf_path)
    date = ''
    amount = 0.0
    seller_name = ''

    if not text:
        return {'date': date, 'amount': amount, 'seller_name': seller_name}

    # 开票日期：2025年12月03日
    dm = re.search(r'开票日期[：:]\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', text)
    if dm:
        date = f"{dm.group(1)}-{int(dm.group(2)):02d}-{int(dm.group(3)):02d}"

    # 销售方名称（电子发票 PDF 中的销方信息）
    # 格式: "销售方" ... "名 称：XXX公司" 或 "销售方名称：XXX"
    sm = re.search(r'销\s*售?\s*方.*?名\s*称[：:]\s*(.+)', text, re.DOTALL)
    if sm:
        name = sm.group(1).strip().split('\n')[0].strip()
        name = re.sub(r'\s+', '', name)
        if name:
            seller_name = name

    # 价税合计（小写）¥123.00
    m = re.search(r'[（(]小写[）)]\s*[￥¥](\d+\.?\d*)', text)
    if m:
        amount = float(m.group(1))
    else:
        m = re.search(r'(?:合计|总计|金额|实收)[^\d]*(\d+\.?\d{0,2})', text)
        if m:
            amount = float(m.group(1))

    return {'date': date, 'amount': amount, 'seller_name': seller_name}


def extract_dining_amount(pdf_path):
    """从餐饮发票 PDF 中提取金额（价税合计），兼容旧调用"""
    return extract_dining_info(pdf_path)['amount']


def _map_project_to_invoice_type(project_name):
    """根据项目名称映射发票类型"""
    if not project_name:
        return '其他发票'
    p = project_name.strip()
    if '餐饮' in p:
        return '餐饮'
    if '住宿' in p:
        return '酒店'
    if '旅客运输' in p or '航空' in p or '铁路' in p or '运输' in p:
        return '机票'
    if '经纪' in p or '代理' in p:
        return '其他发票'
    return '其他发票'


def _extract_third_party_trip(pdf_path, text):
    """提取首汽约车等第三方网约车行程单（滴滴平台发送，格式类似滴滴行程单）。"""
    if not any(kw in text for kw in ['首汽约车', '第三方网约车']):
        return None
    if not any(kw in text for kw in ['行程报销单', '行程单']):
        return None
    if any(kw in text for kw in ['高德', 'AMAP', '曹操出行']):
        return None

    # 行程起止日期：2025-08-20 至 2025-08-21
    trip_dm = re.search(r'行程起止日期[：:]\s*(\d{4}-\d{2}-\d{2})', text)
    fallback_date = trip_dm.group(1) if trip_dm else ''
    total_amount_match = re.search(r'合计(\d+\.?\d*)元', text)
    total_amount = float(total_amount_match.group(1)) if total_amount_match else 0.0

    trips = []
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    header = [c.strip() if c else '' for c in table[0]]
                    if '起点' not in str(header) or '终点' not in str(header):
                        continue
                    start_idx = next((i for i, h in enumerate(header) if '起点' in h), 4)
                    end_idx = next((i for i, h in enumerate(header) if '终点' in h), 5)
                    time_idx = next((i for i, h in enumerate(header) if '时间' in h or '用车' in h), 2)
                    amt_idx = next((i for i, h in enumerate(header) if '金额' in h and '里程' not in h), 6)
                    if amt_idx >= len(header):
                        amt_idx = 7 if len(header) > 7 else 6

                    for row in table[1:]:
                        if not (row and row[0] and str(row[0]).strip().isdigit()):
                            continue
                        if len(row) <= max(start_idx, end_idx):
                            continue

                        start = (row[start_idx] or '').strip().replace('\n', '')
                        end_loc = (row[end_idx] or '').strip().replace('\n', '')

                        row_date = fallback_date
                        time_cell = (row[time_idx] if time_idx < len(row) else '') or ''
                        time_cell = str(time_cell).strip()
                        dm = re.search(r'(\d{4})[-.年/](\d{1,2})[-.月/](\d{1,2})', time_cell)
                        if dm:
                            row_date = f"{dm.group(1)}-{int(dm.group(2)):02d}-{int(dm.group(3)):02d}"
                        elif fallback_date and re.search(r'^(\d{1,2})[-.月/](\d{1,2})', time_cell):
                            md = re.search(r'^(\d{1,2})[-.月/](\d{1,2})', time_cell)
                            year = fallback_date[:4] if len(fallback_date) >= 4 else ''
                            if year:
                                row_date = f"{year}-{int(md.group(1)):02d}-{int(md.group(2)):02d}"

                        row_amount = 0.0
                        for amt_col in [amt_idx, 7, 6]:
                            if amt_col < len(row) and row[amt_col]:
                                am = re.search(r'(\d+\.?\d*)', str(row[amt_col]).strip())
                                if am:
                                    row_amount = float(am.group(1))
                                    break

                        trips.append({
                            'type': 'third_party',
                            'doc_type': '行程单',
                            'original_filename': os.path.basename(pdf_path),
                            'source': 'third_party',
                            'date': row_date,
                            'start_location': start,
                            'end_location': end_loc,
                            'amount': row_amount,
                            'full_text': text[:500]
                        })
    except Exception:
        pass

    if not trips:
        return [{
            'type': 'third_party',
            'doc_type': '行程单',
            'original_filename': os.path.basename(pdf_path),
            'source': 'third_party',
            'date': fallback_date,
            'start_location': '',
            'end_location': '',
            'amount': total_amount,
            'full_text': text[:500]
        }]

    if len(trips) == 1 and trips[0]['amount'] == 0 and total_amount > 0:
        trips[0]['amount'] = total_amount

    return trips


def extract_generic_invoice_info(pdf_path):
    """从通用电子发票 PDF 中提取：项目名称、销售方、开票日期、金额。
    用于餐饮、酒店、机票等非滴滴/12306/高德来源的发票。
    出行类（旅客运输等）优先使用行程/服务/用车日期。
    """
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return None
    if not any(kw in text for kw in ['销售方', '开票日期', '价税合计', '发票']):
        return None
    if any(kw in text for kw in ['滴滴', '12306', '铁路', '高德', 'AMAP']):
        return None

    date = ''
    amount = 0.0
    seller_name = ''
    project_name = ''

    # 出行类发票优先用行程日期（报销筛选应以此为准）
    for pat in [r'行程(?:起止)?日期[：:]\s*(\d{4}-\d{2}-\d{2})', r'行程时间[：:]\s*(\d{4}-\d{2}-\d{2})',
                r'服务日期[：:]\s*(\d{4}-\d{2}-\d{2})', r'用车日期[：:]\s*(\d{4}-\d{2}-\d{2})']:
        dm = re.search(pat, text)
        if dm:
            date = dm.group(1)
            break
    if not date:
        dm = re.search(r'(?:行程(?:起止)?日期|服务日期|用车日期)[：:]\s*(\d{4})年(\d{1,2})月(\d{1,2})日', text)
        if dm:
            date = f"{dm.group(1)}-{int(dm.group(2)):02d}-{int(dm.group(3)):02d}"
    if not date:
        dm = re.search(r'开票日期[：:]\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', text)
        if dm:
            date = f"{dm.group(1)}-{int(dm.group(2)):02d}-{int(dm.group(3)):02d}"

    sm = re.search(r'销\s*售?\s*方.*?名\s*称[：:]\s*(.+)', text, re.DOTALL)
    if sm:
        name = sm.group(1).strip().split('\n')[0].strip()
        name = re.sub(r'\s+', '', name)
        if name:
            seller_name = name

    pm = re.search(r'(?:项目名称|货物或应税劳务[、，]?服务名称)[：:]\s*([^\n]+)', text)
    if pm:
        project_name = pm.group(1).strip()
    if not project_name:
        pm = re.search(r'[\*＊]\s*([^\n]*服务[^\n]*)', text)
        if pm:
            project_name = pm.group(1).strip()

    m = re.search(r'[（(]小写[）)]\s*[￥¥](\d+\.?\d*)', text)
    if m:
        amount = float(m.group(1))
    else:
        m = re.search(r'(?:合计|总计|金额|实收)[^\d]*(\d+\.?\d{0,2})', text)
        if m:
            amount = float(m.group(1))

    inv_type = _map_project_to_invoice_type(project_name)
    return {
        'type': inv_type,
        'doc_type': '发票',
        'original_filename': os.path.basename(pdf_path),
        'source': inv_type,
        'date': date,
        'seller_name': seller_name,
        'restaurant_name': seller_name,
        'amount': amount,
        'full_text': text[:500]
    }


def extract_pdf_info(pdf_path):
    """通用PDF信息提取函数，返回信息列表（可能含多条，如多行程行程单）"""
    # 尝试提取12306信息
    info = extract_12306_info(pdf_path)
    if info:
        return [info]

    # 尝试提取滴滴信息（可能返回多条）
    infos = extract_didi_info(pdf_path)
    if infos:
        return infos

    # 尝试提取高德信息
    info = extract_amap_info(pdf_path)
    if info:
        return [info]

    # 尝试提取首汽约车等第三方网约车行程单（滴滴平台发送）
    infos = _extract_third_party_trip(pdf_path, extract_text_from_pdf(pdf_path) or '')
    if infos:
        return infos

    # 尝试提取通用电子发票（餐饮、酒店、机票等，根据项目名称判断）
    info = extract_generic_invoice_info(pdf_path)
    if info:
        return [info]

    # 如果都不是特定类型，返回基本信息
    text = extract_text_from_pdf(pdf_path)
    return [{
        'type': 'unknown',
        'original_filename': os.path.basename(pdf_path),
        'source': 'unknown',
        'date': '',
        'time': '',
        'departure': '',
        'destination': '',
        'start_location': '',
        'end_location': '',
        'train_number': '',
        'seat_class': '',
        'amount': 0.0,
        'full_text': text[:500]
    }]

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
            infos = extract_pdf_info(pdf_path)
            for info in infos:
                info['filepath'] = str(pdf_path)
                extracted_info.append(info)
                logging.info(f"成功提取信息: {info.get('departure', info.get('start_location', ''))} -> {info.get('destination', info.get('end_location', ''))} | 金额: {info.get('amount', 0)}元")
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