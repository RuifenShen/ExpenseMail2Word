#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
搜索邮件 - 按条件搜索邮件，返回邮件列表
支持发件人、主题、日期范围、是否有附件等条件
"""

import argparse
import json
import logging
import sys
from datetime import datetime, timedelta
from email import message_from_bytes
from email.header import decode_header
from pathlib import Path

# 添加脚本目录到系统路径
sys.path.insert(0, str(Path(__file__).resolve().parent))

from imapclient import IMAPClient


def connect_to_email(config):
    """连接到邮箱"""
    email_config = config.get('email', {})

    server = email_config.get('server')
    port = email_config.get('port', 993)
    username = email_config.get('username')
    password = email_config.get('password')
    use_ssl = email_config.get('use_ssl', True)

    if not all([server, username, password]):
        raise ValueError("邮箱配置不完整，请检查server、username和password字段")

    client = IMAPClient(server, port=port, use_uid=True, ssl=use_ssl)
    client.login(username, password)
    client.select_folder('INBOX')

    return client


def _decode_mime_words(s):
    """解码邮件头"""
    if s is None:
        return ""
    decoded = decode_header(s)
    return ''.join([
        t[0].decode(t[1] or 'utf-8', errors='ignore') if isinstance(t[0], bytes) else t[0]
        for t in decoded
    ])


def _build_date_criteria(search_config):
    """构建日期范围条件"""
    criteria = []
    date_from = search_config.get('date_from')
    date_to = search_config.get('date_to')

    if date_from:
        from_date = datetime.strptime(date_from, '%Y-%m-%d').strftime('%d-%b-%Y')
        criteria.extend(['SINCE', from_date])

    if date_to:
        next_day = datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1)
        criteria.extend(['BEFORE', next_day.strftime('%d-%b-%Y')])

    return criteria


def search_emails(client, config):
    """
    按配置条件搜索邮件，返回邮件列表。
    每个邮件包含: id, subject, from, date, has_attachments, pdf_attachments, total_attachments
    """
    search_config = config.get('search', {})
    has_attachments_filter = search_config.get('has_attachments', False)
    senders = search_config.get('senders', [])
    subjects = search_config.get('subjects', [])

    # QQ IMAP 的 FROM + SINCE/BEFORE 组合搜索有 bug（单独都正常，组合返回 0），
    # 解决方案：分别搜索日期和发件人，在 Python 中取交集
    # SUBJECT + 日期可以正常组合，放在 base_criteria 中
    base_criteria = _build_date_criteria(search_config)
    for subject in subjects:
        base_criteria.extend(['SUBJECT', subject])
    if not base_criteria:
        base_criteria = ['ALL']

    if senders:
        base_uids = set(client.search(base_criteria))
        sender_uids = set()
        for sender in senders:
            result = client.search(['FROM', sender])
            logging.debug(f"发件人 {sender}: {len(result)} 封")
            sender_uids.update(result)
        uids = sorted(base_uids & sender_uids)
    else:
        uids = client.search(base_criteria)

    if not uids:
        logging.info("未找到符合条件的邮件")
        return []

    logging.info(f"找到 {len(uids)} 封符合条件的邮件")

    emails_info = []
    for uid in uids:
        try:
            fetch_data = client.fetch([uid], ['RFC822'])
            if uid not in fetch_data:
                continue

            raw_msg = fetch_data[uid][b'RFC822']
            msg = message_from_bytes(raw_msg)

            subject = _decode_mime_words(msg.get('Subject', ''))
            from_addr = _decode_mime_words(msg.get('From', ''))
            date_str = msg.get('Date', '')

            has_attachments = False
            pdf_attachments = []

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue

                filename = part.get_filename()
                if filename:
                    filename = _decode_mime_words(filename)
                    lower_name = filename.lower()
                    if lower_name.endswith('.pdf') or lower_name.endswith('.zip'):
                        has_attachments = True
                    if lower_name.endswith('.pdf'):
                        pdf_attachments.append(filename)

            # 若配置要求只搜索有附件的邮件，则过滤
            if has_attachments_filter and not has_attachments:
                continue

            email_info = {
                'id': uid,
                'subject': subject,
                'from': from_addr,
                'date': date_str,
                'has_attachments': has_attachments,
                'pdf_attachments': pdf_attachments,
                'total_attachments': len(pdf_attachments)
            }
            emails_info.append(email_info)
            logging.debug(f"邮件 {uid}: {subject[:50]}... | 附件: {len(pdf_attachments)}")

        except Exception as e:
            logging.warning(f"解析邮件 {uid} 时出错: {e}")
            continue

    return emails_info


def main():
    parser = argparse.ArgumentParser(description='按条件搜索邮件，输出邮件列表')
    parser.add_argument('--config', '-c', required=True, help='配置文件路径（JSON格式）')
    parser.add_argument('--output', '-o', default='emails.json', help='输出文件路径（默认 emails.json）')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细日志')

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler(sys.stdout)]
    )

    try:
        with open(args.config, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        logging.error(f"配置文件不存在: {args.config}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        logging.error(f"配置文件格式错误: {e}")
        sys.exit(1)

    try:
        logging.info(f"连接到邮箱: {config.get('email', {}).get('username', '')}")
        client = connect_to_email(config)
        emails = search_emails(client, config)
        client.logout()

        if emails:
            total_pdfs = sum(e.get('total_attachments', 0) for e in emails)
            logging.info(f"搜索完成: {len(emails)} 封邮件, {total_pdfs} 个PDF附件")

            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(emails, f, ensure_ascii=False, indent=2)
            logging.info(f"结果已保存到: {args.output}")
        else:
            logging.info("未找到符合条件的邮件")

    except Exception as e:
        logging.error(f"执行出错: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
