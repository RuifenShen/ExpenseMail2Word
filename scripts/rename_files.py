#!/usr/bin/env python3
"""
重命名文件 - 按规则重命名PDF文件
"""

import json
import argparse
import os
import sys
import re
from pathlib import Path
from datetime import datetime
import shutil
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

def generate_new_filename(info, processing_config=None):
    """根据信息和配置生成新文件名"""
    if processing_config is None:
        processing_config = {}

    rename_rules = processing_config.get('rename_rules', {})
    rename_format = processing_config.get('rename_format', '{date}_{type}_{start}_{end}_{doc_type}.pdf')

    # 根据不同类型生成文件名
    def _date_fallback(filepath):
        """无日期时用文件修改时间生成 YYYYMMDD"""
        try:
            return datetime.fromtimestamp(Path(filepath).stat().st_mtime).strftime('%Y%m%d')
        except OSError:
            return 'unknown'

    if info['type'] == '12306':
        # 12306火车票命名规则
        date_str = info.get('date', '').replace('-', '')
        if not date_str:
            date_str = _date_fallback(info['filepath'])

        start = info.get('departure', '未知').replace(' ', '').replace('-', '')
        end = info.get('destination', '未知').replace(' ', '').replace('-', '')
        train = info.get('train_number', '未知')

        new_name = f"{date_str}_12306_{start}-{end}_{train}.pdf"

    elif info['type'] == 'didi':
        date_str = info.get('date', '').replace('-', '').replace('.', '')
        if not date_str:
            date_str = _date_fallback(info['filepath'])

        start = info.get('start_location', '').replace(' ', '')[:10] or '未知'
        end = info.get('end_location', '').replace(' ', '')[:10] or '未知'

        doc_type = info.get('doc_type', '')
        if not doc_type:
            if '行程单' in info.get('original_filename', '') or '行程' in info.get('full_text', ''):
                doc_type = '行程单'
            elif '发票' in info.get('original_filename', '') or 'invoice' in str(info.get('full_text', '')).lower():
                doc_type = '发票'
            else:
                doc_type = '行程'

        new_name = f"{date_str}_滴滴_{start}_{end}_{doc_type}.pdf"

    elif info['type'] == 'third_party':
        date_str = info.get('date', '').replace('-', '').replace('.', '')
        if not date_str:
            date_str = _date_fallback(info['filepath'])

        start = info.get('start_location', '').replace(' ', '')[:10] or '未知'
        end = info.get('end_location', '').replace(' ', '')[:10] or '未知'

        new_name = f"{date_str}_首汽约车_{start}_{end}_行程单.pdf"

    elif info['type'] == 'amap':
        date_str = info.get('date', '').replace('-', '').replace('.', '')
        if not date_str:
            date_str = _date_fallback(info['filepath'])

        start = info.get('start_location', '').replace(' ', '')[:10] or '未知'
        end = info.get('end_location', '').replace(' ', '')[:10] or '未知'

        doc_type = info.get('doc_type', '')
        if not doc_type:
            if '行程单' in info.get('original_filename', ''):
                doc_type = '行程单'
            elif '发票' in info.get('original_filename', ''):
                doc_type = '发票'
            else:
                doc_type = '行程'

        new_name = f"{date_str}_高德_{start}_{end}_{doc_type}.pdf"

    elif info['type'] in ('dining', '餐饮', '酒店', '机票', '其他发票'):
        date_str = info.get('date', '').replace('-', '').replace('.', '')
        if not date_str:
            date_str = _date_fallback(info['filepath'])

        inv_type = info.get('type', '发票')
        if inv_type == 'dining':
            inv_type = '餐饮'
        if inv_type == '机票' and info.get('pair_id') and (info.get('start_location') or info.get('end_location')):
            start = (info.get('start_location', '') or '').replace(' ', '')[:10] or '未知'
            end = (info.get('end_location', '') or '').replace(' ', '')[:10] or '未知'
            new_name = f"{date_str}_首汽约车_{start}_{end}_发票.pdf"
        else:
            seller = (info.get('seller_name', '') or info.get('restaurant_name', '')).replace(' ', '')[:20] or '未知'
            new_name = f"{date_str}_{inv_type}_{seller}.pdf"
    else:
        # 对于未知类型的文件，使用基本命名规则
        date_str = info.get('date', '').replace('-', '')
        if not date_str:
            date_str = 'unknown'

        source = info.get('source', 'unknown')
        new_name = f"{date_str}_{source}_unknown.pdf"

    # 清理文件名中的非法字符
    new_name = re.sub(r'[<>:"/\\|?*]', '_', new_name)

    # 限制文件名长度
    if len(new_name) > 200:
        name, ext = os.path.splitext(new_name)
        new_name = name[:200-len(ext)] + ext

    return new_name

def rename_files(info_file, input_dir, output_dir, config=None):
    """根据提取的信息重命名文件"""
    if config is None:
        config = {}

    processing_config = config.get('processing', {})

    # 读取信息文件
    with open(info_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    extractions = data.get('extractions', [])

    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    renamed_files = []

    for info in extractions:
        original_path = Path(info['filepath'])

        # 检查原文件是否存在
        if not original_path.exists():
            logging.warning(f"原文件不存在: {original_path}")
            continue

        # 生成新文件名
        new_filename = generate_new_filename(info, processing_config)
        new_path = output_path / new_filename

        # 处理重复文件名
        counter = 1
        original_new_path = new_path
        while new_path.exists():
            stem = original_new_path.stem
            suffix = original_new_path.suffix
            new_path = output_path / f"{stem}_{counter}{suffix}"
            counter += 1

        # 复制文件（保留原文件）
        try:
            shutil.copy2(original_path, new_path)

            renamed_info = {
                'original_path': str(original_path),
                'new_path': str(new_path),
                'original_filename': original_path.name,
                'new_filename': new_path.name,
                'info': info
            }

            renamed_files.append(renamed_info)
            logging.info(f"已重命名: {original_path.name} -> {new_path.name}")
        except Exception as e:
            logging.error(f"重命名文件失败 {original_path}: {str(e)}")
            continue

    return renamed_files

def main():
    parser = argparse.ArgumentParser(description='根据提取的信息重命名PDF文件')
    parser.add_argument('--info', '-i', required=True, help='提取的信息文件路径（JSON格式）')
    parser.add_argument('--input', '-j', required=True, help='包含原PDF文件的目录')
    parser.add_argument('--output', '-o', required=True, help='重命名后文件的输出目录')
    parser.add_argument('--config', '-c', help='配置文件路径')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细日志')

    args = parser.parse_args()
    setup_logging(args.verbose)

    # 读取配置文件
    config = {}
    if args.config:
        try:
            with open(args.config, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except FileNotFoundError:
            logging.error(f"配置文件不存在: {args.config}")
        except json.JSONDecodeError:
            logging.error(f"配置文件格式错误: {args.config}")

    try:
        logging.info(f"开始重命名文件...")
        logging.info(f"输入目录: {args.input}")
        logging.info(f"输出目录: {args.output}")

        renamed_files = rename_files(args.info, args.input, args.output, config)

        # 保存重命名结果
        result = {
            'summary': {
                'total_files': len(renamed_files),
                'input_directory': args.input,
                'output_directory': args.output
            },
            'renamed_files': renamed_files
        }

        output_file = Path(args.output) / 'renamed_files.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        logging.info(f"重命名完成，共处理了 {len(renamed_files)} 个文件")
        logging.info(f"结果已保存到: {output_file}")

    except Exception as e:
        logging.error(f"程序执行出错: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()