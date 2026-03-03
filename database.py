# 在线表格操作
# database.py
"""
在线表格操作模块 - 基于飞书多维表格API
"""

import xbot
from xbot import print, sleep
import pandas as pd
import json
import requests
from datetime import datetime
from config import ONLINE_SHEET_CONFIG

class OnlineSheetManager:
    """在线表格管理器"""
    
    def __init__(self):
        self.config = ONLINE_SHEET_CONFIG
        self.access_token = None
        self.base_url = " https://open.feishu.cn/open-apis/sheets/v2 "
        self._get_access_token()
    
    def _get_access_token(self):
        """获取飞书访问令牌"""
        try:
            url = " https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal "
            headers = {"Content-Type": "application/json"}
            data = {
                "app_id": self.config['app_id'],
                "app_secret": self.config['app_secret']
            }
            
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 200:
                result = response.json()
                self.access_token = result.get('tenant_access_token')
                print(f"获取访问令牌成功: {self.access_token[:20]}...")
            else:
                print(f"获取访问令牌失败: {response.text}")
                self.access_token = None
                
        except Exception as e:
            print(f"获取访问令牌异常: {str(e)}")
            self.access_token = None
    
    def _make_request(self, method, endpoint, data=None):
        """发送API请求"""
        if not self.access_token:
            self._get_access_token()
            if not self.access_token:
                return None
        
        url = f"{self.base_url}/{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            if method == "GET":
                response = requests.get(url, headers=headers, params=data)
            elif method == "POST":
                response = requests.post(url, headers=headers, json=data)
            elif method == "PUT":
                response = requests.put(url, headers=headers, json=data)
            
            if response.status_code == 200:
                return response.json()
            else:
                print(f"API请求失败: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            print(f"API请求异常: {str(e)}")
            return None
    
    def get_all_records(self):
        """获取所有记录"""
        try:
            endpoint = f"spreadsheets/{self.config['sheet_token']}/values/{self.config['sheet_id']}"
            result = self._make_request("GET", endpoint)
            
            if result and 'data' in result:
                values = result['data'].get('valueRange', {}).get('values', [])
                if len(values) > 1:
                    # 第一行是表头
                    headers = values
                    records = []
                    for row in values[1:]:
                        record = {}
                        for i, header in enumerate(headers):
                            if i < len(row):
                                record[header] = row[i]
                            else:
                                record[header] = ''
                        records.append(record)
                    return records
            return []
            
        except Exception as e:
            print(f"获取记录失败: {str(e)}")
            return []
    
    def add_record(self, record_data):
        """添加新记录"""
        try:
            # 先获取现有数据以确定行号
            endpoint = f"spreadsheets/{self.config['sheet_token']}/values/{self.config['sheet_id']}"
            result = self._make_request("GET", endpoint)
            
            if not result:
                return False
            
            values = result['data'].get('valueRange', {}).get('values', [])
            next_row = len(values) + 1 if values else 2  # 从第2行开始
            
            # 准备数据行
            headers = [
                '序号', '登记时间', '发件人', '邮件标题', '邮件内容',
                '投保类型', '投保状态', '邮件回复状态', '保单号',
                '保单生效时间', '保单文件路径', '邮件ID', '处理备注'
            ]
            
            row_data = []
            for header in headers:
                if header == '序号':
                    row_data.append(str(next_row - 1))
                elif header == '登记时间':
                    row_data.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                elif header == '投保状态':
                    row_data.append('未投保')
                elif header == '邮件回复状态':
                    row_data.append('未回复')
                else:
                    row_data.append(record_data.get(header, ''))
            
            # 更新表格
            update_data = {
                "valueRange": {
                    "range": f"{self.config['sheet_id']}!A{next_row}:M{next_row}",
                    "values": [row_data]
                }
            }
            
            endpoint = f"spreadsheets/{self.config['sheet_token']}/values"
            result = self._make_request("PUT", endpoint, update_data)
            
            if result and result.get('code') == 0:
                print(f"添加记录成功: 行{next_row}")
                return True
            else:
                print(f"添加记录失败: {result}")
                return False
                
        except Exception as e:
            print(f"添加记录异常: {str(e)}")
            return False
    
    def update_record(self, search_field, search_value, update_fields):
        """更新记录"""
        try:
            # 先找到要更新的行
            records = self.get_all_records()
            if not records:
                return False
            
            # 查找匹配的记录
            for i, record in enumerate(records):
                if record.get(search_field) == search_value:
                    row_num = i + 2  # 行号（表头在第1行）
                    
                    # 获取当前行的所有值
                    endpoint = f"spreadsheets/{self.config['sheet_token']}/values/{self.config['sheet_id']}!A{row_num}:M{row_num}"
                    result = self._make_request("GET", endpoint)
                    
                    if not result:
                        return False
                    
                    current_values = result['data'].get('valueRange', {}).get('values', [[]])
                    
                    # 更新字段
                    headers = [
                        '序号', '登记时间', '发件人', '邮件标题', '邮件内容',
                        '投保类型', '投保状态', '邮件回复状态', '保单号',
                        '保单生效时间', '保单文件路径', '邮件ID', '处理备注'
                    ]
                    
                    for field, value in update_fields.items():
                        if field in headers:
                            index = headers.index(field)
                            if index < len(current_values):
                                current_values[index] = value
                    
                    # 写回更新
                    update_data = {
                        "valueRange": {
                            "range": f"{self.config['sheet_id']}!A{row_num}:M{row_num}",
                            "values": [current_values]
                        }
                    }
                    
                    endpoint = f"spreadsheets/{self.config['sheet_token']}/values"
                    result = self._make_request("PUT", endpoint, update_data)
                    
                    if result and result.get('code') == 0:
                        print(f"更新记录成功: {search_field}={search_value}")
                        return True
                    else:
                        return False
            
            print(f"未找到匹配记录: {search_field}={search_value}")
            return False
            
        except Exception as e:
            print(f"更新记录异常: {str(e)}")
            return False
    
    def get_records_by_status(self, status_field, status_value):
        """根据状态获取记录"""
        try:
            records = self.get_all_records()
            filtered = []
            
            for record in records:
                if record.get(status_field) == status_value:
                    filtered.append(record)
            
            return filtered
            
        except Exception as e:
            print(f"按状态查询失败: {str(e)}")
            return []

# 本地备份功能（可选）
class LocalBackupManager:
    """本地Excel备份管理器"""
    
    def __init__(self, backup_path='D:/保险保单/备份/'):
        self.backup_path = backup_path
    
    def backup_to_excel(self, records, filename=None):
        """备份到本地Excel"""
        try:
            if not filename:
                filename = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            df = pd.DataFrame(records)
            filepath = f"{self.backup_path}/{filename}"
            df.to_excel(filepath, index=False)
            print(f"本地备份完成: {filepath}")
            return True
            
        except Exception as e:
            print(f"本地备份失败: {str(e)}")
            return False
