#!/usr/bin/env python3
"""
下载附件 - 下载邮件中的PDF/ZIP附件
"""

import json
import argparse
import os
import sys
import zipfile
from email import message_from_bytes
from email.header import decode_header
from imapclient import IMAPClient
import logging
from pathlib import Path

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

def _decode_mime_filename(raw_filename):
    """解码 MIME 编码的文件名（=?UTF-8?b?...?= 等）"""
    if raw_filename is None:
        return None
    decoded_parts = decode_header(raw_filename)
    return ''.join([
        part.decode(charset or 'utf-8', errors='ignore') if isinstance(part, bytes) else part
        for part, charset in decoded_parts
    ])


def _unique_filepath(output_path, filename):
    """确保唯一文件名（如果已存在则加后缀）"""
    filepath = output_path / filename
    counter = 1
    original = filepath
    while filepath.exists():
        stem, ext = os.path.splitext(original.name)
        filepath = output_path / f"{stem}_{counter}{ext}"
        counter += 1
    return filepath


def _extract_zip(zip_path, output_path):
    """解压 ZIP 文件中的 PDF，返回提取的附件信息列表"""
    extracted = []
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for member in zf.namelist():
                if member.lower().endswith('.pdf'):
                    member_name = Path(member).name
                    dest = _unique_filepath(output_path, member_name)
                    with zf.open(member) as src, open(dest, 'wb') as dst:
                        dst.write(src.read())
                    extracted.append({
                        'filename': member_name,
                        'filepath': str(dest),
                        'size': os.path.getsize(dest),
                        'from_zip': str(zip_path)
                    })
                    logging.info(f"  从ZIP解压: {member_name} ({os.path.getsize(dest)} bytes)")
    except zipfile.BadZipFile:
        logging.warning(f"无效的ZIP文件: {zip_path}")
    return extracted


def _save_attachment(part, output_path, msg_id):
    """保存单个附件，如果是 ZIP 则自动解压 PDF"""
    raw_filename = part.get_filename()
    filename = _decode_mime_filename(raw_filename)
    if not filename:
        return []

    file_ext = os.path.splitext(filename)[1].lower()
    if file_ext not in ['.pdf', '.zip']:
        logging.info(f"跳过非PDF/ZIP附件: {filename} (类型: {file_ext})")
        return []

    filepath = _unique_filepath(output_path, filename)
    with open(filepath, 'wb') as f:
        f.write(part.get_payload(decode=True))

    results = []
    if file_ext == '.zip':
        logging.info(f"已下载附件: {filename} ({os.path.getsize(filepath)} bytes)")
        extracted = _extract_zip(filepath, output_path)
        results.extend(extracted)
        if not extracted:
            results.append({
                'filename': filename, 'filepath': str(filepath),
                'size': os.path.getsize(filepath), 'msg_id': msg_id
            })
    else:
        logging.info(f"已下载附件: {filename} ({os.path.getsize(filepath)} bytes)")
        results.append({
            'filename': filename, 'filepath': str(filepath),
            'size': os.path.getsize(filepath), 'msg_id': msg_id
        })
    return results


def extract_attachments(client, msg_id, output_dir):
    """从邮件中提取附件"""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    raw_msg = client.fetch([msg_id], ['RFC822'])
    email_body = raw_msg[msg_id][b'RFC822']
    msg = message_from_bytes(email_body)

    attachments = []

    if msg.is_multipart():
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition"))
            if "attachment" in content_disposition:
                attachments.extend(_save_attachment(part, output_path, msg_id))
    else:
        content_disposition = str(msg.get("Content-Disposition"))
        if "attachment" in content_disposition:
            attachments.extend(_save_attachment(msg, output_path, msg_id))

    return attachments

def download_attachments_from_list(emails, config, output_dir):
    """从邮件列表下载附件"""
    email_config = config.get('email', {})
    server = email_config.get('server')
    port = email_config.get('port', 993)
    username = email_config.get('username')
    password = email_config.get('password')
    use_ssl = email_config.get('use_ssl', True)

    # 重新连接以下载附件
    client = IMAPClient(server, port=port, use_uid=True, ssl=use_ssl)
    client.login(username, password)
    client.select_folder('INBOX')

    all_attachments = []

    for email_data in emails:
        msg_id = email_data['id']
        logging.info(f"正在处理邮件 {msg_id}: {email_data.get('subject', '')}")

        try:
            attachments = extract_attachments(client, msg_id, output_dir)
            all_attachments.extend(attachments)
        except Exception as e:
            logging.error(f"处理邮件 {msg_id} 时出错: {str(e)}")
            continue

    client.logout()
    return all_attachments

def main():
    parser = argparse.ArgumentParser(description='从邮件下载PDF/ZIP附件')
    parser.add_argument('--emails', '-e', required=True, help='邮件列表文件路径（JSON格式）')
    parser.add_argument('--config', '-c', help='配置文件路径（如果提供邮件列表，则不需要此参数）')
    parser.add_argument('--output', '-o', required=True, help='附件输出目录路径')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细日志')

    args = parser.parse_args()
    setup_logging(args.verbose)

    # 读取邮件列表
    try:
        with open(args.emails, 'r', encoding='utf-8') as f:
            emails = json.load(f)
    except FileNotFoundError:
        logging.error(f"邮件列表文件不存在: {args.emails}")
        sys.exit(1)
    except json.JSONDecodeError:
        logging.error(f"邮件列表文件格式错误: {args.emails}")
        sys.exit(1)

    # 读取配置文件
    if args.config:
        try:
            with open(args.config, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except FileNotFoundError:
            logging.error(f"配置文件不存在: {args.config}")
            sys.exit(1)
        except json.JSONDecodeError:
            logging.error(f"配置文件格式错误: {args.config}")
            sys.exit(1)
    else:
        # 如果未提供配置，尝试从环境变量或其他来源获取
        logging.warning("未提供配置文件，需要邮箱连接信息")
        return

    try:
        logging.info(f"开始下载附件到: {args.output}")
        attachments = download_attachments_from_list(emails, config, args.output)

        logging.info(f"总共下载了 {len(attachments)} 个附件")

        # 输出附件列表
        attachment_list = {
            'output_dir': args.output,
            'attachments': attachments,
            'total_count': len(attachments),
            'total_size': sum(att['size'] for att in attachments)
        }

        output_file = Path(args.output) / 'attachments.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(attachment_list, f, ensure_ascii=False, indent=2)

        logging.info(f"附件列表已保存到: {output_file}")

    except Exception as e:
        logging.error(f"程序执行出错: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()