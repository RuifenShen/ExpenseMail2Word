#!/usr/bin/env python3
"""
完整流程处理 - 执行完整的报销处理流程
"""

import json
import argparse
import os
import sys
from pathlib import Path
import logging
import tempfile

# 添加脚本目录到系统路径以便导入其他脚本（跨平台）
sys.path.insert(0, str(Path(__file__).resolve().parent))

# 导入各个模块
from search_emails import connect_to_email, search_emails
from download_attachments import download_attachments_from_list
from extract_pdf_info import extract_pdf_info
from rename_files import generate_new_filename
from create_summary_doc import create_summary_document

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

def run_complete_process(config, output_dir, processing_type=None):
    """运行完整处理流程"""
    # 创建输出目录
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # 创建临时目录用于存储中间文件
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # 步骤1: 搜索邮件
        logging.info("步骤1: 搜索邮件...")
        client = connect_to_email(config)
        emails = search_emails(client, config)
        client.logout()

        if not emails:
            logging.info("未找到符合条件的邮件")
            return

        # 保存邮件列表
        emails_file = temp_path / 'emails.json'
        with open(emails_file, 'w', encoding='utf-8') as f:
            json.dump(emails, f, ensure_ascii=False, indent=2)

        # 步骤2: 下载附件
        logging.info("步骤2: 下载附件...")
        attachments_dir = temp_path / 'attachments'
        attachments_dir.mkdir(parents=True, exist_ok=True)

        attachments = download_attachments_from_list(emails, config, str(attachments_dir))

        if not attachments:
            logging.info("未找到符合条件的附件")
            return

        # 步骤3: 提取PDF信息
        logging.info("步骤3: 提取PDF信息...")
        pdfs_dir = attachments_dir  # 附件已经下载到pdfs_dir

        # 获取所有PDF文件
        pdf_files = list(pdfs_dir.rglob('*.pdf'))

        extracted_info = []
        for pdf_path in pdf_files:
            try:
                info = extract_pdf_info(pdf_path)
                if info:
                    info['filepath'] = str(pdf_path)
                    extracted_info.append(info)
                    logging.info(f"成功提取信息: {info.get('departure', info.get('start_location', ''))} -> {info.get('destination', info.get('end_location', ''))} | 金额: {info.get('amount', 0)}元")
            except Exception as e:
                logging.error(f"处理PDF {pdf_path} 时出错: {str(e)}")
                continue

        # 步骤3.5: 配对行程单与发票（滴滴 + 高德，按金额匹配）
        for src in ('didi', 'amap'):
            trips = [i for i in extracted_info if i.get('type') == src and i.get('doc_type') == '行程单']
            invoices = [i for i in extracted_info if i.get('type') == src and i.get('doc_type') == '发票']
            paired_inv_idx = set()

            for trip in trips:
                for j, inv in enumerate(invoices):
                    if j not in paired_inv_idx and abs(trip.get('amount', 0) - inv.get('amount', 0)) < 0.01:
                        inv['start_location'] = trip.get('start_location', '')
                        inv['end_location'] = trip.get('end_location', '')
                        inv['date'] = trip.get('date', '')
                        pair_id = f"{src}_{trip.get('date', '')}_{trip.get('amount', 0)}"
                        trip['pair_id'] = pair_id
                        inv['pair_id'] = pair_id
                        paired_inv_idx.add(j)
                        logging.info(f"配对: {trip['original_filename']} <-> {inv['original_filename']}")
                        break

        def _sort_key(item):
            d = item.get('date', '')
            pair = item.get('pair_id', '')
            doc = 0 if item.get('doc_type', '') != '发票' else 1
            return (d, pair, doc)

        extracted_info.sort(key=_sort_key)

        # 保存提取的信息
        info_file = temp_path / 'info.json'
        result = {
            'summary': {
                'total_files': len(pdf_files),
                'successful_extractions': len(extracted_info),
                'failed_extractions': len(pdf_files) - len(extracted_info)
            },
            'extractions': extracted_info
        }
        with open(info_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        # 步骤4: 重命名文件
        logging.info("步骤4: 重命名文件...")
        renamed_dir = output_path / 'renamed_pdfs'
        renamed_dir.mkdir(parents=True, exist_ok=True)

        # 根据提取的信息重命名文件
        for info in extracted_info:
            original_path = Path(info['filepath'])

            # 生成新文件名
            new_filename = generate_new_filename(info, config.get('processing', {}))
            new_path = renamed_dir / new_filename

            # 处理重复文件名
            counter = 1
            original_new_path = new_path
            while new_path.exists():
                stem = original_new_path.stem
                suffix = original_new_path.suffix
                new_path = renamed_dir / f"{stem}_{counter}{suffix}"
                counter += 1

            # 复制文件
            import shutil
            shutil.copy2(original_path, new_path)

            # 更新info中的filepath
            info['filepath'] = str(new_path)
        # 更新并保存提取的信息
        result['extractions'] = extracted_info
        with open(info_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        # 步骤5: 创建汇总文档
        logging.info("步骤5: 创建汇总文档...")

        # 生成文件名: 报销开始日期_报销结束日期_出差地点_报销单生成日期.docx
        expense_items = [
            ex for ex in extracted_info
            if not (ex.get('doc_type') == '发票' and ex.get('pair_id'))
        ]
        dates = sorted(d for ex in expense_items if (d := ex.get('date', '')) and d != '')
        date_start = dates[0].replace('-', '') if dates else 'unknown'
        date_end = dates[-1].replace('-', '') if dates else 'unknown'

        def _extract_city(station_name):
            """从站名提取城市: 北京西→北京, 长沙南→长沙, 上海虹桥→上海"""
            import re
            city = re.sub(r'(虹桥|南站?|北站?|西站?|东站?|站)$', '', station_name)
            return city or station_name[:2]

        cities = set()
        for ex in expense_items:
            if ex.get('type') == '12306':
                for key in ('departure', 'destination'):
                    v = ex.get(key, '')
                    if v:
                        cities.add(_extract_city(v))
        location_str = '出差' + ''.join(sorted(cities)) if cities else '报销汇总'

        from datetime import datetime as _dt
        gen_date = _dt.now().strftime('%Y%m%d')

        doc_filename = f"{date_start}_{date_end}_{location_str}_{gen_date}.docx"
        doc_output = output_path / doc_filename
        create_summary_document(str(info_file), str(renamed_dir), str(doc_output), config)

        logging.info(f"处理完成！汇总文档已保存至: {doc_output}")
        logging.info(f"重命名的PDF文件保存在: {renamed_dir}")

def main():
    parser = argparse.ArgumentParser(description='执行完整的报销处理流程')
    parser.add_argument('--config', '-c', required=True, help='配置文件路径')
    parser.add_argument('--output', '-o', default='./output', help='输出目录路径')
    parser.add_argument('--type', '-t', help='处理类型 (12306, didi, 等)')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细日志')
    parser.add_argument('--log-file', help='日志文件路径')

    args = parser.parse_args()

    # 设置日志
    if args.log_file:
        # 同时记录到文件和控制台
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG if args.verbose else logging.INFO)

        # 控制台处理器
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.DEBUG if args.verbose else logging.INFO)
        ch.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(ch)

        # 文件处理器
        fh = logging.FileHandler(args.log_file, encoding='utf-8')
        fh.setLevel(logging.DEBUG if args.verbose else logging.INFO)
        fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(fh)
    else:
        setup_logging(args.verbose)

    # 读取配置文件
    try:
        with open(args.config, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        logging.error(f"配置文件不存在: {args.config}")
        sys.exit(1)
    except json.JSONDecodeError:
        logging.error(f"配置文件格式错误: {args.config}")
        sys.exit(1)

    # 若未指定 --output，则优先使用配置文件中的 output_dir（与 test_config 等兼容）
    output_dir = args.output
    if output_dir == './output':
        output_dir = (
            config.get('workflow', {}).get('output_dir')
            or config.get('processing', {}).get('output_dir')
            or output_dir
        )

    try:
        logging.info("开始执行完整报销处理流程...")
        run_complete_process(config, output_dir, args.type)

    except KeyboardInterrupt:
        logging.info("用户中断了操作")
        sys.exit(1)
    except Exception as e:
        logging.error(f"程序执行出错: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()