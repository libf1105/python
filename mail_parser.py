# 邮件解析与登记
# mail_parser.py
"""
应用1：邮件解析与登记
定时任务：每天12:00和21:00执行
"""

# import xbot
# from xbot import print, sleep
import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime
import os

import pandas as pd
import json
import requests

from config import EMAIL_CONFIG, KEYWORDS, FILE_CONFIG
from database import LocalSheetManager
from utils import (
    create_directory_structure,
    analyze_insurance_type,
    analyze_goods_type,
    analyze_import_export_type,
    extract_bill_number,
    parse_excel_attachment,
    save_attachment,
    write_log
)

class MailParser:
    """邮件解析器"""
    
    def __init__(self):
        self.email_config = EMAIL_CONFIG
        self.sheet_manager = LocalSheetManager()
        create_directory_structure()
    
    def connect_email(self):
        """连接邮箱"""
        try:
            mail = imaplib.IMAP4_SSL(
                self.email_config['imap_server'],
                self.email_config['imap_port']
            )
            mail.login(
                self.email_config['username'],
                self.email_config['password']
            )
            print("邮箱连接成功")
            return mail
        except Exception as e:
            write_log('应用1', f"邮箱连接失败: {str(e)}", 'ERROR')
            return None
    
    def decode_subject(self, subject):
        """解码邮件主题"""
        try:
            decoded = decode_header(subject)
            subject_parts = []
            for part, encoding in decoded:
                if isinstance(part, bytes):
                    if encoding:
                        subject_parts.append(part.decode(encoding))
                    else:
                        subject_parts.append(part.decode('utf-8', errors='ignore'))
                else:
                    subject_parts.append(str(part))
            return ''.join(subject_parts)
        except:
            return str(subject)
    
    def parse_email_content(self, msg):
        """解析邮件内容"""
        email_data = {
            'subject': '',
            'from': '',
            'date': '',
            'receive_time': '',
            'body': '',
            'attachments': []
        }
        
        try:
            # 解析发件人
            from_header = msg.get('From', '')
            email_data['from'] = str(from_header)
            
            # 解析主题
            subject = msg.get('Subject', '')
            email_data['subject'] = self.decode_subject(subject)
            
            # 解析日期
            date_header = msg.get('Date', '')
            email_data['date'] = str(date_header)
            
            # 解析邮件接收时间
            from email.utils import parsedate_to_datetime
            try:
                dt = parsedate_to_datetime(date_header)
                email_data['receive_time'] = dt.strftime('%Y-%m-%d %H:%M:%S')
            except:
                email_data['receive_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # 解析邮件ID
            email_id = msg.get('Message-ID', '')
            email_data['email_id'] = str(email_id)
            
            # 解析正文
            body = ''
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    # 获取正文
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        try:
                            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        except:
                            try:
                                body = part.get_payload(decode=True).decode('gbk', errors='ignore')
                            except:
                                body = part.get_payload(decode=True).decode('latin-1', errors='ignore')
                    
                    # 获取附件
                    elif "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename:
                            # 解码附件名
                            if isinstance(filename, str):
                                decoded_filename = filename
                            else:
                                decoded_parts = decode_header(filename)
                                decoded_filename = ''
                                for part_data, encoding in decoded_parts:
                                    if isinstance(part_data, bytes):
                                        decoded_filename += part_data.decode(encoding or 'utf-8', errors='ignore')
                                    else:
                                        decoded_filename += str(part_data)
                            
                            attachment_data = {
                                'filename': decoded_filename,
                                'content': part.get_payload(decode=True)
                            }
                            email_data['attachments'].append(attachment_data)
            else:
                try:
                    body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                except:
                    try:
                        body = msg.get_payload(decode=True).decode('gbk', errors='ignore')
                    except:
                        body = msg.get_payload(decode=True).decode('latin-1', errors='ignore')
            
            email_data['body'] = body
            return email_data
            
        except Exception as e:
            write_log('应用1', f"解析邮件内容失败: {str(e)}", 'ERROR')
            return email_data
    
    def should_process_mail(self, email_data):
        """判断是否应该处理该邮件"""
        # 检查是否包含忽略关键词
        content = f"{email_data['subject']} {email_data['body']}".lower()
        for keyword in KEYWORDS['ignore_keywords']:
            if keyword.lower() in content:
                write_log('应用1', f"忽略测试邮件: {email_data['subject']}", 'INFO')
                return False
        
        # 检查是否与保险相关
        insurance_keywords = ['保险', '投保', '货运', '运输', '保单']
        has_insurance_keyword = any(keyword in content for keyword in insurance_keywords)
        
        if not has_insurance_keyword:
            write_log('应用1', f"忽略非保险邮件: {email_data['subject']}", 'INFO')
            return False
        
        return True
    
    def process_single_mail(self, mail, email_uid):
        """处理单封邮件"""
        try:
            # 获取邮件原始数据
            result, data = mail.fetch(email_uid, '(RFC822)')
            if result != 'OK':
                return False
            
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # 解析邮件内容
            email_data = self.parse_email_content(msg)
            
            # 判断是否处理
            if not self.should_process_mail(email_data):
                return False
            
            # 检查记录是否已存在
            if self.sheet_manager.record_exists(
                email_data['receive_time'],
                email_data['from'],
                email_data['subject']
            ):
                write_log('应用1', 
                    f"邮件已存在，跳过处理: {email_data['subject']} | 接收时间: {email_data['receive_time']} | 发件人: {email_data['from']}",
                    'WARNING'
                )
                # 标记重复邮件为已读
                mail.store(email_uid, '+FLAGS', '\\Seen')
                return False
            
            # 分析投保类型
            insurance_type = analyze_insurance_type(
                email_data['body'],
                email_data['subject']
            )
            
            # 提取提单号
            bill_number = extract_bill_number(email_data['body'])
            
            # 初始化货物信息
            goods_name = ''
            goods_mark = ''
            goods_type = '普货'
            import_export_type = '出口'
            
            # 处理附件
            attachment_paths = []
            attachment_info = {}
            if email_data['attachments']:
                # 获取邮件接收日期
                receive_date = email_data['receive_time'].split(' ')[0]
                attachment_dir = os.path.join(FILE_CONFIG['attachment_dir'], receive_date)
                if not os.path.exists(attachment_dir):
                    os.makedirs(attachment_dir)
                
                for attachment in email_data['attachments']:
                    filename = attachment['filename']
                    # 解码附件名
                    if '=?' in filename:
                        decoded_parts = decode_header(filename)
                        decoded_filename = ''
                        for part_data, encoding in decoded_parts:
                            if isinstance(part_data, bytes):
                                decoded_filename += part_data.decode(encoding or 'utf-8', errors='ignore')
                            else:
                                decoded_filename += str(part_data)
                        filename = decoded_filename
                    
                    filepath = os.path.join(attachment_dir, filename)
                    
                    # 保存附件（指定UTF-8编码）
                    with open(filepath, 'wb') as f:
                        f.write(attachment['content'])
                    
                    # 使用Windows路径格式
                    windows_path = filepath.replace('/', '\\')
                    attachment_paths.append(windows_path)
                    print(f"附件保存成功: {windows_path}")
                    
                    # 如果是Excel文件，解析内容
                    if filename.lower().endswith(('.xlsx', '.xls')):
                        excel_data = parse_excel_attachment(filepath)
                        if excel_data:
                            attachment_info = excel_data
                            # 提取货物信息
                            main_info = excel_data.get('main_info', {})
                            goods_info = main_info.get('goods', {})
                            transport_info = main_info.get('transport', {})
                            
                            # 货物名称
                            goods_name = goods_info.get('货物描述', '')
                            # 唔头
                            goods_mark = goods_info.get('唔头', '')
                            # 起运国家地区
                            origin_country = transport_info.get('起运国家地区', '')
                            
                            # 分析货物类型
                            goods_type = analyze_goods_type(goods_name)
                            # 分析进出口类型
                            import_export_type = analyze_import_export_type(origin_country)
            
            # 准备记录数据（所有字符串字段去除首尾空格）
            def s(v):
                return str(v).strip() if v is not None else ''

            record_data = {
                '邮件接收时间': s(email_data['receive_time']),
                '发件人': s(email_data['from']),
                '邮件标题': s(email_data['subject']),
                '邮件内容': ' '.join(email_data['body'][:500].split()),
                '邮件附件': '; '.join(attachment_paths) if attachment_paths else '',
                '投保类型': s(insurance_type),
                '货物名称': s(goods_name),
                '货物类型': s(goods_type),
                '进出口类型': s(import_export_type),
                '邮件ID': s(email_data.get('email_id', '')),
                '处理备注': s(f"提单号: {bill_number} | 唛头: {goods_mark}")
            }
            
            # 添加Excel附件解析字段
            if attachment_info:
                main_info = attachment_info.get('main_info', {})
                insured_info = main_info.get('insured', {})
                transport_info = main_info.get('transport', {})
                goods_info = main_info.get('goods', {})
                invoice_info = main_info.get('invoice', {})
                
                record_data.update({
                    '被保险人': s(insured_info.get('被保险人', '')),
                    '联系电话': s(insured_info.get('联系电话', '')),
                    '通讯地址': s(insured_info.get('通讯地址', '')),
                    '起运地': s(transport_info.get('起运地', '')),
                    '起运国家地区': s(transport_info.get('起运国家地区', '')),
                    '目的地': s(transport_info.get('目的地', '')),
                    '目的国家地区': s(transport_info.get('国家地区', '')),
                    '转运地': s(transport_info.get('转运地', '')),
                    '起运日期': s(transport_info.get('起运日期', '')),
                    '起运日期打印格式': s(transport_info.get('起运日期打印格式', '')),
                    '运输工具号': s(transport_info.get('运输工具号', '')),
                    '运输方式': s(transport_info.get('运输方式', '')),
                    '装载方式': s(transport_info.get('装载方式', '')),
                    '赔款偿付地点': s(transport_info.get('赔款偿付地点', '')),
                    '贸易类型': s(transport_info.get('贸易类型', '')),
                    '唔头': s(goods_info.get('唔头', '')),
                    '包装数量': s(goods_info.get('包装数量', '')),
                    '条款': s(goods_info.get('条款', '')),
                    '发票号': s(invoice_info.get('发票号', '')),
                    '提单号': s(invoice_info.get('提单号', '')),
                    '发票金额': invoice_info.get('发票金额', ''),
                    '发票币种': s(invoice_info.get('发票币种', '')),
                    '工作编号': s(invoice_info.get('工作编号', '')),
                    '备注': s(invoice_info.get('备注', ''))
                })
            
            # 添加到在线表格
            success = self.sheet_manager.add_record(record_data)
            
            if success:
                write_log('应用1', 
                    f"邮件登记成功: {email_data['subject']} | 类型: {insurance_type}",
                    'INFO'
                )
                
                # 标记邮件为已读
                mail.store(email_uid, '+FLAGS', '\\Seen')
                return True
            else:
                write_log('应用1', 
                    f"邮件登记失败: {email_data['subject']}",
                    'ERROR'
                )
                return False
                
        except Exception as e:
            write_log('应用1', f"处理邮件异常: {str(e)}", 'ERROR')
            return False
    
    def run(self):
        """主执行函数"""
        write_log('应用1', "开始执行邮件解析任务", 'INFO')
        
        try:
            # 连接邮箱
            mail = self.connect_email()
            if not mail:
                return False
            
            # 选择文件夹
            result = mail.select(self.email_config['folder'])
            if result[0] != 'OK':
                write_log('应用1', f"选择文件夹失败: {self.email_config['folder']}", 'ERROR')
                mail.logout()
                return False
            
            # 搜索未读邮件
            result, data = mail.search(None, 'UNSEEN')
            if result != 'OK':
                write_log('邮件解析', "没有找到未读邮件", 'INFO')
                mail.logout()
                return True
            
            email_uids = data[0].split()
            total_count = len(email_uids)
            
            if total_count == 0:
                write_log('应用1', "没有未读邮件需要处理", 'INFO')
                mail.logout()
                return True
            
            write_log('应用1', f"找到 {total_count} 封未读邮件", 'INFO')
            
            # 处理每封邮件
            success_count = 0
            for i, email_uid in enumerate(email_uids):
                write_log('应用1', f"处理第 {i+1}/{total_count} 封邮件", 'INFO')
                
                if self.process_single_mail(mail, email_uid):
                    success_count += 1
                
                # 避免处理过快
                import time
                time.sleep(1)
            
            # 登出
            mail.logout()
            
            write_log('应用1', 
                f"邮件解析任务完成: 成功 {success_count}/{total_count}",
                'INFO'
            )
            
            return True
            
        except Exception as e:
            write_log('应用1', f"邮件解析任务异常: {str(e)}", 'ERROR')
            return False

def main():
    """主函数 - 用于影刀调用"""
    parser = MailParser()
    success = parser.run()
    
    if success:
        return {"status": "success", "message": "邮件解析任务完成"}
    else:
        return {"status": "error", "message": "邮件解析任务失败"}

if __name__ == "__main__":
    import sys
    import io
    # 设置标准输出编码为UTF-8
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # 本地测试
    print(pd.__version__)
    print(requests.__version__)
    print(json.__version__)

    result = main()
    print(result)
