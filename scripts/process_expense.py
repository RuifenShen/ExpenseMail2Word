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
import re
from email.utils import parsedate_to_datetime

from search_emails import connect_to_email, search_emails
from download_attachments import download_attachments_from_list
from extract_pdf_info import extract_pdf_info
from rename_files import generate_new_filename
from create_summary_doc import create_summary_document


def _parse_email_date(date_str):
    """将邮件 Date 头解析为 YYYY-MM-DD"""
    try:
        dt = parsedate_to_datetime(date_str)
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return ''


def _extract_restaurant_name(subject):
    """从邮件主题中提取饭店/公司名称。
    常见格式：
    - 【电子发票】北京湘采隐小厨餐饮有限公司（发票金额：515.00元）
    - 来自 北京湘采隐小厨餐饮有限公司 的电子发票
    - 百望云/诺诺等平台: 含发票号码的主题
    """
    # 格式1: 】公司名（  或 】公司名(
    m = re.search(r'[】\]]\s*(.+?)\s*[（(]', subject)
    if m:
        name = m.group(1).strip()
        if name and not name.isdigit():
            return name
    # 格式2: "来自" + 公司名
    m = re.search(r'来自[：:\s]*(.+?)(?:[的）)（(,，]|$)', subject)
    if m:
        name = m.group(1).strip()
        if name and not name.isdigit():
            return name
    # 格式3: 发票 + 公司名（跳过纯数字发票号码）
    m = re.search(r'发票[】\]：:\s]*(.+?)(?:[（(]|$)', subject)
    if m:
        name = m.group(1).strip()
        name = re.sub(r'^[\d\s号码：:]+', '', name).strip()
        if name and not name.isdigit():
            return name
    # 兜底：取主题中最长的连续中文（排除常见干扰词）
    skip_words = {'电子发票', '发票', '百望云', '诺诺', '您', '一张', '待查收', '已开具', '餐饮'}
    parts = [p for p in re.findall(r'[\u4e00-\u9fff]+', subject) if p not in skip_words]
    return max(parts, key=len) if parts else '未知餐饮'


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

    # 中间结果目录（保留邮件、附件列表、筛选前后提取结果，便于排查）
    intermediate_dir = output_path / 'intermediate'
    intermediate_dir.mkdir(parents=True, exist_ok=True)

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

        # 保存邮件列表（临时 + 持久化到 intermediate 便于排查）
        emails_file = temp_path / 'emails.json'
        with open(emails_file, 'w', encoding='utf-8') as f:
            json.dump(emails, f, ensure_ascii=False, indent=2)
        with open(intermediate_dir / 'emails.json', 'w', encoding='utf-8') as f:
            json.dump(emails, f, ensure_ascii=False, indent=2)
        logging.info(f"中间结果已保存: {intermediate_dir / 'emails.json'}")

        # 步骤2: 下载附件
        logging.info("步骤2: 下载附件...")
        attachments_dir = temp_path / 'attachments'
        attachments_dir.mkdir(parents=True, exist_ok=True)

        attachments = download_attachments_from_list(emails, config, str(attachments_dir))

        if not attachments:
            logging.info("未找到符合条件的附件")
            return

        # 保存附件列表到 intermediate（含邮件主题、发送日期、发件人，便于排查）
        import shutil
        att_save = [{'filename': Path(a.get('filepath', '')).name, 'filepath': a.get('filepath', ''),
                     'email_subject': a.get('email_subject', ''), 'email_date': a.get('email_date', ''),
                     'email_from': a.get('email_from', '')} for a in attachments]
        with open(intermediate_dir / 'attachments.json', 'w', encoding='utf-8') as f:
            json.dump(att_save, f, ensure_ascii=False, indent=2)
        att_copy_dir = intermediate_dir / 'attachments'
        att_copy_dir.mkdir(parents=True, exist_ok=True)
        for a in attachments:
            src = Path(a.get('filepath', ''))
            if src.exists():
                shutil.copy2(src, att_copy_dir / src.name)
        logging.info(f"中间结果已保存: {intermediate_dir / 'attachments.json'}, {att_copy_dir}")

        # 构建 附件文件路径 → 邮件元数据 的映射
        file_email_map = {}
        for att in attachments:
            fp = att.get('filepath', '')
            if fp:
                file_email_map[Path(fp).name] = {
                    'subject': att.get('email_subject', ''),
                    'date': att.get('email_date', ''),
                    'from': att.get('email_from', ''),
                }

        # 步骤3: 提取PDF信息
        logging.info("步骤3: 提取PDF信息...")
        pdfs_dir = attachments_dir

        pdf_files = list(pdfs_dir.rglob('*.pdf'))

        extracted_info = []
        for pdf_path in pdf_files:
            try:
                email_meta = file_email_map.get(pdf_path.name, {})
                email_subject = email_meta.get('subject', '')
                infos = extract_pdf_info(pdf_path)
                for info in infos:
                    info['filepath'] = str(pdf_path)
                    if info.get('type') in ('餐饮', '酒店', '机票', '其他发票'):
                        if not info.get('date'):
                            info['date'] = _parse_email_date(email_meta.get('date', ''))
                        if not info.get('seller_name') and not info.get('restaurant_name'):
                            info['seller_name'] = _extract_restaurant_name(email_subject)
                            info['restaurant_name'] = info['seller_name']
                        label = info.get('type', '发票')
                        logging.info(f"{label}发票: {info.get('seller_name', info.get('restaurant_name', ''))} | 日期: {info.get('date', '')} | 金额: {info.get('amount', 0)}元")
                    else:
                        logging.info(f"成功提取信息: {info.get('departure', info.get('start_location', ''))} -> {info.get('destination', info.get('end_location', ''))} | 金额: {info.get('amount', 0)}元")
                    extracted_info.append(info)
            except Exception as e:
                logging.error(f"处理PDF {pdf_path} 时出错: {str(e)}")
                continue

        # 步骤3.5: 去重（重复邮件可能导致同一凭证被多次下载）
        def _dedup_key(item):
            t = item.get('type', '')
            doc = item.get('doc_type', '')
            d = item.get('date', '')
            amt = f"{item.get('amount', 0):.2f}"
            if t == '12306':
                return (t, d, item.get('train_number', ''), amt)
            if t in ('dining', '餐饮', '酒店', '机票', '其他发票'):
                return (t, d, item.get('seller_name', item.get('restaurant_name', '')), amt)
            if t in ('didi', 'amap', 'third_party'):
                start_loc = item.get('start_location', '') or item.get('departure', '')
                end_loc = item.get('end_location', '') or item.get('destination', '')
                return (t, doc, d, start_loc, end_loc, amt)
            return (t, doc, d, amt, item.get('original_filename', ''))

        seen_keys = set()
        deduped = []
        for item in extracted_info:
            key = _dedup_key(item)
            if key in seen_keys:
                logging.info(f"去重: 跳过重复凭证 {item.get('original_filename', '')} (key={key})")
                continue
            seen_keys.add(key)
            deduped.append(item)
        if len(deduped) < len(extracted_info):
            logging.info(f"去重完成: {len(extracted_info)} -> {len(deduped)} 条（去除 {len(extracted_info) - len(deduped)} 条重复）")
        extracted_info = deduped

        # 步骤3.6: 配对行程单与发票（滴滴 + 高德 + 第三方网约车，按金额匹配，支持多行程→单发票）
        from collections import defaultdict
        for src in ('didi', 'amap', 'third_party'):
            trips = [i for i in extracted_info if i.get('type') == src and i.get('doc_type') == '行程单']
            invoices = [i for i in extracted_info if i.get('type') == src and i.get('doc_type') == '发票']
            if src == 'third_party':
                invoices = [i for i in extracted_info if i.get('doc_type') == '发票' and i.get('type') == '机票' and not i.get('pair_id')]
            paired_inv_idx = set()

            trip_groups = defaultdict(list)
            for trip in trips:
                trip_groups[trip['original_filename']].append(trip)

            for filename, group in trip_groups.items():
                group_sum = sum(t.get('amount', 0) for t in group)
                for j, inv in enumerate(invoices):
                    if j not in paired_inv_idx and abs(group_sum - inv.get('amount', 0)) < 0.01:
                        pair_id = f"{src}_{group[0].get('date', '')}_{group_sum}"
                        for trip in group:
                            trip['pair_id'] = pair_id
                        inv['start_location'] = group[0].get('start_location', '')
                        inv['end_location'] = group[0].get('end_location', '')
                        inv['date'] = group[0].get('date', '')
                        inv['pair_id'] = pair_id
                        paired_inv_idx.add(j)
                        logging.info(f"配对: {filename} ({len(group)}条行程) <-> {inv['original_filename']}")
                        break

        def _sort_key(item):
            d = item.get('date', '')
            pair = item.get('pair_id', '')
            doc = 0 if item.get('doc_type', '') != '发票' else 1
            return (d, pair, doc)

        extracted_info.sort(key=_sort_key)

        # 行程日期筛选参数（单数字月/日自动补零，如 2025-8-12 -> 2025-08-12）
        def _norm_d(s):
            s = (config.get('search', {}).get(s) or '').strip()
            if not s:
                return s
            m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', s)
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else s
        trip_date_from = _norm_d('trip_date_from')
        trip_date_to = _norm_d('trip_date_to')

        # 保存筛选前的提取结果到 intermediate（便于排查被 trip_date 过滤掉的条目）
        # 为每条提取结果附加来源邮件信息，便于追溯
        extractions_with_email = []
        for e in extracted_info:
            rec = dict(e)
            fn = Path(e.get('filepath', '')).name
            meta = file_email_map.get(fn, {})
            rec['_email_subject'] = meta.get('subject', '')
            rec['_email_date'] = meta.get('date', '')
            rec['_email_from'] = meta.get('from', '')
            extractions_with_email.append(rec)
        info_before_filter = {
            'summary': {'total': len(extracted_info), 'trip_date_from': trip_date_from,
                        'trip_date_to': trip_date_to},
            'extractions': extractions_with_email
        }
        with open(intermediate_dir / 'info_before_filter.json', 'w', encoding='utf-8') as f:
            json.dump(info_before_filter, f, ensure_ascii=False, indent=2)
        logging.info(f"中间结果已保存: {intermediate_dir / 'info_before_filter.json'}")

        # 按行程起止日期筛选（可选）
        if trip_date_from or trip_date_to:
            def _in_trip_range(item):
                d = item.get('date', '')
                if not d:
                    return False
                if trip_date_from and d < trip_date_from:
                    return False
                if trip_date_to and d > trip_date_to:
                    return False
                return True
            before_count = len(extracted_info)
            extracted_info.sort(key=lambda x: x.get('date', ''))
            filtered = []
            for x in extracted_info:
                if trip_date_to and x.get('date', '') > trip_date_to:
                    break
                if _in_trip_range(x):
                    filtered.append(x)
            extracted_info = filtered
            extracted_info.sort(key=_sort_key)
            logging.info(f"按行程日期筛选: {trip_date_from or '不限'} ~ {trip_date_to or '不限'}，保留 {len(extracted_info)}/{before_count} 条")

        # 保存筛选后的提取结果到 intermediate
        with open(intermediate_dir / 'info_after_filter.json', 'w', encoding='utf-8') as f:
            json.dump({'extractions': [dict(e) for e in extracted_info]}, f, ensure_ascii=False, indent=2)

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

        # 根据提取的信息重命名文件（同一原始 PDF 只复制一次）
        import shutil
        renamed_originals = {}
        for info in extracted_info:
            original_path = Path(info['filepath'])
            orig_name = info.get('original_filename', original_path.name)

            if orig_name in renamed_originals:
                info['filepath'] = str(renamed_originals[orig_name])
                continue

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
            shutil.copy2(original_path, new_path)
            renamed_originals[orig_name] = new_path

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
        logging.info(f"中间结果已保存至: {intermediate_dir} (emails.json/attachments.json/info_before_filter.json/info_after_filter.json，可对比排查被过滤的条目)")

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

    # 日期配置校验（单数字月/日自动补零）
    def _norm_date(s):
        s = (s or '').strip()
        if not s:
            return s
        m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', s)
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else s

    search_cfg = config.get('search', {})
    email_from = _norm_date(search_cfg.get('email_send_date_from'))
    email_to = _norm_date(search_cfg.get('email_send_date_to'))
    trip_from = _norm_date(search_cfg.get('trip_date_from'))
    trip_to = _norm_date(search_cfg.get('trip_date_to'))
    if trip_from and trip_to and trip_to < trip_from:
        logging.error(f"配置错误: trip_date_to ({trip_to}) 应大于等于 trip_date_from ({trip_from})")
        sys.exit(1)
    if trip_from and email_from and email_from < trip_from:
        logging.error(f"配置错误: email_send_date_from ({email_from}) 应大于等于 trip_date_from ({trip_from})")
        sys.exit(1)
    if trip_to and email_to and email_to < trip_to:
        logging.error(f"配置错误: email_send_date_to ({email_to}) 应大于等于 trip_date_to ({trip_to})")
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