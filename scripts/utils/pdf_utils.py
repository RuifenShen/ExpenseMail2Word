"""
PDF处理工具函数
"""

import pdfplumber
import fitz  # PyMuPDF
import re
import os
from pathlib import Path
import logging


def extract_text_from_pdf(pdf_path):
    """使用多种方法提取PDF文本"""
    text_content = ""
    try:
        # 首先尝试使用pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    text_content += text + "\n"
    except Exception as e:
        logging.debug(f"使用pdfplumber读取PDF失败: {str(e)}")

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

    return text_content.strip()


def convert_pdf_to_images(pdf_path, output_dir, dpi=150, max_width_inches=6.0):
    """将PDF转换为图片"""
    try:
        doc = fitz.open(pdf_path)

        image_paths = []
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)

            # 设置缩放比例以达到所需DPI
            zoom = dpi / 72.0  # 72是默认DPI
            mat = fitz.Matrix(zoom, zoom)

            pix = page.get_pixmap(matrix=mat)

            # 生成图片文件名
            base_name = Path(pdf_path).stem
            img_filename = f"{base_name}_page_{page_num + 1}.png"
            img_path = Path(output_dir) / img_filename

            pix.save(str(img_path))
            image_paths.append(str(img_path))

        doc.close()
        return image_paths
    except Exception as e:
        logging.error(f"转换PDF到图片失败 {pdf_path}: {str(e)}")
        return []


def detect_document_type(text):
    """检测文档类型"""
    text_lower = text.lower()

    # 检测12306
    if any(keyword in text for keyword in ["12306", "铁路", "中国铁路", "客票", "乘车信息提示"]):
        return "12306"

    # 检测滴滴
    if any(keyword in text_lower for keyword in ["滴滴", "小桔科技", "行程报销", "快车", "专车", "出租车", "行程信息"]):
        return "didi"

    # 可以继续添加其他类型
    return "unknown"


def validate_pdf_file(pdf_path):
    """验证PDF文件是否有效"""
    if not os.path.exists(pdf_path):
        return False, "文件不存在"

    if not str(pdf_path).lower().endswith('.pdf'):
        return False, "不是PDF文件"

    try:
        # 尝试打开PDF以确认它是有效的
        with pdfplumber.open(pdf_path) as pdf:
            if pdf.pages:
                return True, "有效PDF文件"
            else:
                return False, "PDF文件为空或损坏"
    except Exception as e:
        return False, f"PDF文件可能已损坏: {str(e)}"


def clean_filename(filename):
    """清理文件名中的非法字符"""
    # 移除或替换非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(illegal_chars, '_', filename)

    # 限制文件名长度
    max_length = 200
    if len(cleaned) > max_length:
        name, ext = os.path.splitext(cleaned)
        cleaned = name[:max_length-len(ext)] + ext

    return cleaned


def extract_date_from_text(text):
    """从文本中提取日期"""
    # 匹配 YYYY-MM-DD 或 YYYY.MM.DD 或 YYYY/MM/DD 格式
    date_patterns = [
        r'(\d{4}[年\-/.]\d{1,2}[月\-/.]\d{1,2}日?)',  # 支持 年-月-日 或 年.月.日 或 年/月/日
        r'(\d{4}[年\-/.]\d{2}[月\-/.]\d{2})',  # 严格格式
        r'(\d{2}[年\-/.]\d{2}[月\-/.]\d{2})',  # 简短格式
    ]

    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            date_str = match.group(1)
            # 标准化日期格式
            date_str = re.sub(r'[年月]', '-', date_str)
            date_str = re.sub(r'[日/]', '', date_str)
            return date_str

    return ""


def extract_amount_from_text(text):
    """从文本中提取金额"""
    # 匹配各种金额格式
    amount_patterns = [
        r'[￥\$]?\s*(\d+\.?\d*)\s*元',  # ￥100元 或 100元
        r'[￥\$]\s*(\d+\.?\d*)',       # ￥100
        r'\b(\d+\.?\d+)\s*元\b',       # 100元
        r'金额[:：]\s*[￥\$]?\s*(\d+\.?\d*)',  # 金额:100
    ]

    for pattern in amount_patterns:
        matches = re.findall(pattern, text)
        if matches:
            try:
                # 返回最大的金额值（可能是单价+总计的情况）
                amounts = [float(match) for match in matches]
                return max(amounts)
            except ValueError:
                continue

    return 0.0