"""
邮件处理工具函数
"""

import json
import logging
from datetime import datetime
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
    client.select_folder('INBOX')  # 选择收件箱

    return client


def format_search_criteria(search_config):
    """格式化搜索条件"""
    criteria = []

    # 日期范围
    date_from = search_config.get('date_from')
    date_to = search_config.get('date_to')

    if date_from:
        from_date = datetime.strptime(date_from, '%Y-%m-%d').strftime('%d-%b-%Y')
        criteria.append(f'SINCE {from_date}')

    if date_to:
        # BEFORE在IMAP中是排他的，所以我们要加上结束日期的后一天
        to_date_obj = datetime.strptime(date_to, '%Y-%m-%d')
        next_day = to_date_obj.replace(day=to_date_obj.day + 1)
        next_day_str = next_day.strftime('%d-%b-%Y')
        criteria.append(f'BEFORE {next_day_str}')

    # 发件人
    senders = search_config.get('senders', [])
    for sender in senders:
        criteria.append(f'FROM "{sender}"')

    # 主题关键词
    subjects = search_config.get('subjects', [])
    for subject in subjects:
        criteria.append(f'SUBJECT "{subject}"')

    # 如果没有指定条件，默认搜索所有邮件
    if not criteria:
        criteria = ['ALL']

    return criteria


def parse_email_envelope(envelope):
    """解析邮件信封信息"""
    from_part = envelope.from_[0] if envelope.from_ else None
    to_part = envelope.to[0] if envelope.to else None

    parsed = {
        'subject': envelope.subject.decode('utf-8', errors='ignore') if envelope.subject else '',
        'from': {
            'name': from_part.name.decode('utf-8', errors='ignore') if from_part and from_part.name else '',
            'mailbox': from_part.mailbox.decode('utf-8', errors='ignore') if from_part and from_part.mailbox else '',
            'host': from_part.host.decode('utf-8', errors='ignore') if from_part and from_part.host else ''
        } if from_part else {},
        'to': {
            'name': to_part.name.decode('utf-8', errors='ignore') if to_part and to_part.name else '',
            'mailbox': to_part.mailbox.decode('utf-8', errors='ignore') if to_part and to_part.mailbox else '',
            'host': to_part.host.decode('utf-8', errors='ignore') if to_part and to_part.host else ''
        } if to_part else {},
        'date': envelope.date.strftime('%Y-%m-%d %H:%M:%S') if envelope.date else ''
    }

    return parsed


def safe_decode(value):
    """安全解码字节串"""
    if isinstance(value, bytes):
        try:
            return value.decode('utf-8')
        except UnicodeDecodeError:
            try:
                return value.decode('gbk')
            except UnicodeDecodeError:
                return value.decode('utf-8', errors='ignore')
    return value