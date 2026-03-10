#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
回复邮件模块
reply_emails.py
"""
# import xbot
# from xbot import print, sleep
import pandas as pd
import os
import sys
import io
import imaplib
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from config import FILE_CONFIG, EMAIL_CONFIG


class EmailReplyHandler:
    """邮件回复处理器"""
    
    def __init__(self):
        self.imap_conn = None
        self.smtp_conn = None
    
    def connect_email(self):
        """连接邮箱服务器"""
        try:
            # 连接IMAP
            self.imap_conn = imaplib.IMAP4_SSL(EMAIL_CONFIG['imap_server'], EMAIL_CONFIG['imap_port'])
            self.imap_conn.login(EMAIL_CONFIG['username'], EMAIL_CONFIG['password'])
            
            # 连接SMTP
            self.smtp_conn = smtplib.SMTP_SSL(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
            self.smtp_conn.login(EMAIL_CONFIG['username'], EMAIL_CONFIG['password'])
            
            print("邮箱连接成功")
            return True
        except Exception as e:
            print(f"邮箱连接失败: {str(e)}")
            return False
    
    def find_original_email_by_id(self, email_id):
        """根据邮件ID查找原始邮件"""
        try:
            self.imap_conn.select(EMAIL_CONFIG['folder'])
            
            # 使用邮件ID搜索
            search_criteria = f'HEADER Message-ID "{email_id}"'
            result, data = self.imap_conn.search(None, search_criteria)
            
            if result != 'OK' or not data[0]:
                print(f"未找到邮件ID对应的邮件: {email_id}")
                return None
            
            email_uids = data[0].split()
            if email_uids:
                email_uid = email_uids[0]  # 取第一个匹配的邮件
                result, data = self.imap_conn.fetch(email_uid, '(RFC822)')
                if result == 'OK':
                    msg = email.message_from_bytes(data[0][1])
                    print(f"找到邮件ID对应的邮件: {msg.get('Subject', '')[:50]}...")
                    return msg, email_uid
            
            return None
        except Exception as e:
            print(f"按邮件ID查找失败: {str(e)}")
            return None
    
    def send_reply_email(self, original_msg, policy_file_path, record):
        """发送回复邮件"""
        try:
            # 创建回复邮件
            reply_msg = MIMEMultipart()
            
            # 获取收件人信息
            to_address = original_msg.get('From')
            subject = f"Re: {original_msg.get('Subject', '')}"
            
            print(f"准备发送回复邮件:")
            print(f"  收件人: {to_address}")
            print(f"  主题: {subject}")
            print(f"  保单文件: {policy_file_path}")
            
            # 设置邮件头
            reply_msg['From'] = EMAIL_CONFIG['username']
            reply_msg['To'] = to_address
            reply_msg['Subject'] = subject
            reply_msg['In-Reply-To'] = original_msg.get('Message-ID', '')
            reply_msg['References'] = original_msg.get('Message-ID', '')
            
            # 添加邮件正文
            try:
                body = EMAIL_CONFIG['reply_body']
                # 确保正文是UTF-8编码的字符串
                if isinstance(body, bytes):
                    body = body.decode('utf-8', errors='ignore')
                elif not isinstance(body, str):
                    body = str(body)
                
                # 创建正文部分
                text_part = MIMEText(body, 'plain', 'utf-8')
                reply_msg.attach(text_part)
                print(f"  正文添加成功，长度: {len(body)} 字符")
            except Exception as body_error:
                print(f"  正文处理失败: {str(body_error)}")
                # 使用默认正文
                default_body = "尊敬的客户，\n\n您好！\n\n您的货物保险已成功投保，保单文件请见附件。\n\n谢谢！"
                reply_msg.attach(MIMEText(default_body, 'plain', 'utf-8'))
                print(f"  使用默认正文")
            
            # 添加保单附件
            if os.path.exists(policy_file_path):
                try:
                    with open(policy_file_path, 'rb') as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename= {os.path.basename(policy_file_path)}'
                        )
                        reply_msg.attach(part)
                    print(f"  附件添加成功: {os.path.basename(policy_file_path)}")
                except Exception as attach_error:
                    print(f"  附件添加失败: {str(attach_error)}")
                    return False
            else:
                print(f"  警告: 保单文件不存在 {policy_file_path}")
                return False
            
            # 发送邮件
            try:
                self.smtp_conn.send_message(reply_msg)
                print(f"回复邮件发送成功: {record.get('序号', '')} - {record.get('邮件标题', '')}")
                return True
            except smtplib.SMTPException as smtp_error:
                print(f"SMTP发送失败: {str(smtp_error)}")
                print(f"  SMTP服务器: {EMAIL_CONFIG['smtp_server']}:{EMAIL_CONFIG['smtp_port']}")
                print(f"  发件人: {EMAIL_CONFIG['username']}")
                print(f"  收件人: {to_address}")
                return False
            
        except Exception as e:
            print(f"发送回复邮件失败 - 详细错误信息:")
            print(f"  错误类型: {type(e).__name__}")
            print(f"  错误描述: {str(e)}")
            print(f"  记录序号: {record.get('序号', '')}")
            print(f"  邮件标题: {record.get('邮件标题', '')}")
            print(f"  保单路径: {policy_file_path}")
            import traceback
            print(f"  详细堆栈: {traceback.format_exc()}")
            return False
    
    def close_connections(self):
        """关闭连接"""
        try:
            if self.imap_conn:
                self.imap_conn.close()
                self.imap_conn.logout()
            if self.smtp_conn:
                self.smtp_conn.quit()
        except:
            pass
    
    def process_downloaded_policies(self):
        """处理已下载保单的记录"""
        filename = FILE_CONFIG.get('record_filename')
        if filename:
            excel_path = os.path.join(os.path.dirname(FILE_CONFIG['excel_path']), filename)
        else:
            excel_path = FILE_CONFIG['excel_path']
        
        try:
            if not os.path.exists(excel_path):
                print(f"文件不存在: {excel_path}")
                return False
            
            # 读取Excel文件
            df = pd.read_excel(excel_path, engine='openpyxl', dtype=str)
            df = df.fillna('')
            
            # 筛选已下载保单的记录
            downloaded_records = df[df['投保状态'] == '已下载保单']
            
            if len(downloaded_records) == 0:
                print("没有已下载保单的记录")
                return True
            
            print(f"找到 {len(downloaded_records)} 条已下载保单记录")
            
            # 连接邮箱
            if not self.connect_email():
                return False
            
            success_count = 0
            
            # 处理每条记录
            for index, record in downloaded_records.iterrows():
                try:
                    email_id = record.get('邮件ID', '')
                    policy_file_path = record.get('保单文件路径', '')
                    
                    if not email_id or not policy_file_path:
                        print(f"记录信息不完整，跳过: 序号 {record.get('序号', '')}")
                        continue
                    
                    print(f"处理记录: 序号 {record.get('序号', '')} - 邮件ID: {email_id[:20]}...")
                    
                    # 根据邮件ID查找原始邮件
                    result = self.find_original_email_by_id(email_id)
                    if not result:
                        print(f"未找到原始邮件，跳过: 序号 {record.get('序号', '')}")
                        continue
                    
                    original_msg, email_uid = result
                    
                    # 发送回复邮件
                    if self.send_reply_email(original_msg, policy_file_path, record):
                        # 更新状态为已发送邮件
                        df.loc[index, '投保状态'] = '已发送邮件'
                        success_count += 1
                    
                except Exception as e:
                    print(f"处理记录失败: 序号 {record.get('序号', '')} - {str(e)}")
                    continue
            
            # 保存更新后的文件
            if success_count > 0:
                try:
                    df.to_excel(excel_path, index=False, engine='openpyxl')
                    print(f"成功发送 {success_count} 封回复邮件，状态已更新")
                except PermissionError:
                    print(f"成功发送 {success_count} 封回复邮件，但无法更新Excel文件（文件可能被占用）")
                except Exception as e:
                    print(f"成功发送 {success_count} 封回复邮件，但保存文件失败: {str(e)}")
            
            return True
            
        except Exception as e:
            print(f"处理失败: {str(e)}")
            return False
        finally:
            self.close_connections()


def main():
    """主函数"""
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    print("=" * 60)
    print("邮件回复处理")
    print("=" * 60)
    
    handler = EmailReplyHandler()
    success = handler.process_downloaded_policies()
    
    if success:
        print("邮件回复处理完成")
    else:
        print("邮件回复处理失败")


if __name__ == "__main__":
    main()