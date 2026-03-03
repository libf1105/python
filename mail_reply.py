# 邮件回复
# mail_replier.py
"""
应用4：邮件回复
定时任务：每天14:00和23:00执行
"""

import xbot
from xbot import print, sleep
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime
import os

from config import EMAIL_CONFIG
from database import OnlineSheetManager
from utils import write_log, clean_temp_files

class MailReplier:
    """邮件回复器"""
    
    def __init__(self):
        self.email_config = EMAIL_CONFIG
        self.sheet_manager = OnlineSheetManager()
    
    def connect_smtp(self):
        """连接SMTP服务器"""
        try:
            smtp = smtplib.SMTP_SSL(
                self.email_config['smtp_server'],
                self.email_config['smtp_port']
            )
            smtp.login(
                self.email_config['username'],
                self.email_config['password']
            )
            print("SMTP连接成功")
            return smtp
        except Exception as e:
            write_log('应用4', f"SMTP连接失败: {str(e)}", 'ERROR')
            return None
    
    def connect_imap(self):
        """连接IMAP服务器"""
        try:
            mail = imaplib.IMAP4_SSL(
                self.email_config['imap_server'],
                self.email_config['imap_port']
            )
            mail.login(
                self.email_config['username'],
                self.email_config['password']
            )
            print("IMAP连接成功")
            return mail
        except Exception as e:
            write_log('应用4', f"IMAP连接失败: {str(e)}", 'ERROR')
            return None
    
    def find_original_mail(self, mail, sender, subject):
        """查找原始邮件"""
        try:
            # 选择收件箱
            mail.select('inbox')
            
            # 构建搜索条件
            search_criteria = f'(FROM "{sender}" SUBJECT "{subject}")'
            result, data = mail.search(None, search_criteria)
            
            if result != 'OK' or not data[0]:
                write_log('应用4', f"未找到原始邮件: {sender} - {subject}", 'WARN')
                return None
            
            # 获取最新的匹配邮件
            email_uids = data[0].split()
            latest_uid = email_uids[-1]  # 最新的邮件
            
            # 获取邮件信息
            result, data = mail.fetch(latest_uid, '(RFC822)')
            if result != 'OK':
                return None
            
            raw_email = data
            msg = email.message_from_bytes(raw_email)
            
            # 获取邮件ID
            email_id = msg.get('Message-ID', '')
            if email_id:
                email_id = email_id.strip('<>')
            
            write_log('应用4', f"找到原始邮件: {email_id}", 'INFO')
            return email_id
            
        except Exception as e:
            write_log('应用4', f"查找原始邮件异常: {str(e)}", 'ERROR')
            return None
    
    def create_reply_email(self, original_sender, original_subject, policy_file_path):
        """创建回复邮件"""
        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            
            # 邮件头
            msg['From'] = self.email_config['username']
            msg['To'] = original_sender
            msg['Subject'] = Header(f'Re: {original_subject}', 'utf-8')
            
            # 邮件正文
            body = """尊敬的客户：

您的货运险保单已生效，保单详情请见附件。

保单号：见附件文件名
生效时间：见系统记录

如有任何问题，请随时联系我们。

此致
敬礼

保险服务团队
"""
            
            text_part = MIMEText(body, 'plain', 'utf-8')
            msg.attach(text_part)
            
            # 添加附件
            if policy_file_path and os.path.exists(policy_file_path):
                with open(policy_file_path, 'rb') as f:
                    file_data = f.read()
                
                attachment = MIMEText(file_data, 'base64', 'utf-8')
                attachment["Content-Type"] = 'application/pdf'
                attachment["Content-Disposition"] = f'attachment; filename="{os.path.basename(policy_file_path)}"'
                msg.attach(attachment)
                
                write_log('应用4', f"添加附件: {policy_file_path}", 'INFO')
            else:
                write_log('应用4', "保单文件不存在，邮件无附件", 'WARN')
            
            return msg
            
        except Exception as e:
            write_log('应用4', f"创建回复邮件失败: {str(e)}", 'ERROR')
            return None
    
    def send_email(self, smtp, msg):
        """发送邮件"""
        try:
            smtp.send_message(msg)
            write_log('应用4', "邮件发送成功", 'INFO')
            return True
        except Exception as e:
            write_log('应用4', f"邮件发送失败: {str(e)}", 'ERROR')
            return False
    
    def process_single_record(self, record):
        """处理单条记录"""
        try:
            sender = record.get('发件人', '')
            subject = record.get('邮件标题', '')
            policy_number = record.get('保单号', '')
            policy_file_path = record.get('保单文件路径', '')
            
            if not sender or not subject:
                write_log('应用4', "记录信息不全，跳过", 'WARN')
                return False
            
            if not policy_file_path or not os.path.exists(policy_file_path):
                write_log('应用4', f"保单文件不存在: {policy_file_path}", 'WARN')
                # 仍然发送邮件，但没有附件
                policy_file_path = None
            
            write_log('应用4', 
                f"开始回复邮件: {sender} - {subject}",
                'INFO'
            )
            
            # 连接IMAP查找原始邮件
            imap = self.connect_imap()
            if not imap:
                return False
            
            original_email_id = self.find_original_mail(imap, sender, subject)
            imap.logout()
            
            # 创建回复邮件
            msg = self.create_reply_email(sender, subject, policy_file_path)
            if not msg:
                return False
            
            # 连接SMTP发送邮件
            smtp = self.connect_smtp()
            if not smtp:
                return False
            
            send_success = self.send_email(smtp, msg)
            smtp.quit()
            
            if send_success:
                # 更新表格状态
                update_fields = {
                    '邮件回复状态': '已回复',
                    '处理备注': f"{record.get('处理备注', '')} | 回复时间: {datetime.now().strftime('%H:%M:%S')}"
                }
                
                success = self.sheet_manager.update_record(
                    '保单号',  # 搜索字段
                    policy_number,  # 搜索值
                    update_fields
                )
                
                if success:
                    write_log('应用4', 
                        f"邮件回复成功: {sender} - {policy_number}",
                        'INFO'
                    )
                    return True
                else:
                    write_log('应用4', f"更新表格失败: {policy_number}", 'ERROR')
                    return False
            else:
                write_log('应用4', f"邮件发送失败: {sender}", 'ERROR')
                return False
                
        except Exception as e:
            write_log('应用4', f"处理记录异常: {str(e)}", 'ERROR')
            return False
    
    def run(self):
        """主执行函数"""
        write_log('应用4', "开始执行邮件回复任务", 'INFO')
        
        try:
            # 获取需要处理的记录（邮件回复状态=未回复 且 投保状态=已生效）
            records = self.sheet_manager.get_all_records()
            
            if not records:
                write_log('应用4', "没有需要回复的邮件", 'INFO')
                return True
            
            # 筛选符合条件的记录
            filtered_records = []
            for record in records:
                if (record.get('邮件回复状态') == '未回复' and 
                    record.get('投保状态') == '已生效'):
                    filtered_records.append(record)
            
            if not filtered_records:
                write_log('应用4', "没有符合条件的待回复邮件", 'INFO')
                return True
            
            total_count = len(filtered_records)
            write_log('应用4', f"找到 {total_count} 封待回复邮件", 'INFO')
            
            # 处理每条记录
            success_count = 0
            for i, record in enumerate(filtered_records):
                write_log('应用4', f"处理第 {i+1}/{total_count} 封邮件", 'INFO')
                
                if self.process_single_record(record):
                    success_count += 1
                
                # 避免发送过快
                sleep(3)
            
            # 清理临时文件
            clean_temp_files()
            
            write_log('应用4', 
                f"邮件回复任务完成: 成功 {success_count}/{total_count}",
                'INFO'
            )
            
            return True
            
        except Exception as e:
            write_log('应用4', f"邮件回复任务异常: {str(e)}", 'ERROR')
            return False

def main():
    """主函数 - 用于影刀调用"""
    replier = MailReplier()
    success = replier.run()
    
    if success:
        return {"status": "success", "message": "邮件回复任务完成"}
    else:
        return {"status": "error", "message": "邮件回复任务失败"}

if __name__ == "__main__":
    # 本地测试
    result = main()
    print(result)
