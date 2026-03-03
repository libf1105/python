# 邮件解析与登记
# mail_parser.py
"""
应用1：邮件解析与登记
定时任务：每天12:00和21:00执行
"""

import xbot
from xbot import print, sleep
import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime
import os

from config import EMAIL_CONFIG, KEYWORDS
from database import OnlineSheetManager
from utils import (
    create_directory_structure,
    analyze_insurance_type,
    extract_bill_number,
    parse_excel_attachment,
    save_attachment,
    write_log
)

class MailParser:
    """邮件解析器"""
    
    def __init__(self):
        self.email_config = EMAIL_CONFIG
        self.sheet_manager = OnlineSheetManager()
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
                            body = part.get_payload(decode=True).decode()
                        except:
                            body = part.get_payload(decode=True).decode('gbk', errors='ignore')
                    
                    # 获取附件
                    elif "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename:
                            attachment_data = {
                                'filename': filename,
                                'content': part.get_payload(decode=True)
                            }
                            email_data['attachments'].append(attachment_data)
            else:
                try:
                    body = msg.get_payload(decode=True).decode()
                except:
                    body = msg.get_payload(decode=True).decode('gbk', errors='ignore')
            
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
            
            raw_email = data
            msg = email.message_from_bytes(raw_email)
            
            # 解析邮件内容
            email_data = self.parse_email_content(msg)
            
            # 判断是否处理
            if not self.should_process_mail(email_data):
                return False
            
            # 分析投保类型
            insurance_type = analyze_insurance_type(
                email_data['body'],
                email_data['subject']
            )
            
            # 提取提单号
            bill_number = extract_bill_number(email_data['body'])
            
            # 处理附件
            attachment_info = {}
            if email_data['attachments']:
                for attachment in email_data['attachments']:
                    filename = attachment['filename']
                    if filename.lower().endswith(('.xlsx', '.xls')):
                        # 保存附件
                        filepath = save_attachment(attachment['content'], filename)
                        if filepath:
                            # 解析Excel
                            excel_data = parse_excel_attachment(filepath)
                            attachment_info = excel_data
            
            # 准备记录数据
            record_data = {
                '发件人': email_data['from'],
                '邮件标题': email_data['subject'],
                '邮件内容': email_data['body'][:500],  # 只保存前500字符
                '投保类型': insurance_type,
                '邮件ID': email_data.get('email_id', ''),
                '处理备注': f"提单号: {bill_number}"
            }
            
            # 合并附件信息
            if attachment_info:
                record_data['处理备注'] += f" | 附件信息: {str(attachment_info)}"
            
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
            mail.select(self.email_config['folder'])
            
            # 搜索未读邮件
            result, data = mail.search(None, 'UNSEEN')
            if result != 'OK':
                write_log('应用1', "没有找到未读邮件", 'INFO')
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
                sleep(1)
            
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
    # 本地测试
    result = main()
    print(result)
