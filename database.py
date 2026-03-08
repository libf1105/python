# 在线表格操作
# database.py
"""
在线表格操作模块 - 基于本地Excel
"""

# import xbot
# from xbot import print, sleep
import pandas as pd
import os
from datetime import datetime
from config import FILE_CONFIG

class LocalSheetManager:
    """本地Excel管理器"""
    
    def __init__(self):
        self.headers = [
            '序号', '登记时间', '邮件接收时间', '发件人', '邮件标题', '邮件内容', '邮件附件',
            '投保类型','货物名称','货物类型','进出口类型', '投保状态', '邮件回复状态', '保单号',
            '保单生效时间', '保单文件路径', '邮件ID', '处理备注',
            # Excel附件解析字段
            '被保险人', '联系电话', '通讯地址',
            '起运地', '起运国家地区', '目的地', '目的国家地区', '转运地', '起运日期', '起运日期打印格式',
            '运输工具号', '运输方式', '装载方式', '赔款偿付地点', '贸易类型',
            '唔头', '包装数量', '条款',
            '发票号', '提单号', '发票金额', '发票币种', '工作编号', '备注'
        ]
        self.excel_path = self._get_daily_excel_path()
        self._init_excel()
    
    def _get_daily_excel_path(self):
        """获取当天的Excel文件路径"""
        base_path = FILE_CONFIG['excel_path']
        dir_name = os.path.dirname(base_path)
        file_name, file_ext = os.path.splitext(os.path.basename(base_path))
        date_suffix = datetime.now().strftime('%Y%m%d')
        return os.path.join(dir_name, f"{file_name}_{date_suffix}{file_ext}")
    
    def _init_excel(self):
        """初始化Excel文件"""
        try:
            if not os.path.exists(self.excel_path):
                df = pd.DataFrame(columns=self.headers)
                df.to_excel(self.excel_path, index=False)
                print(f"创建Excel文件: {self.excel_path}")
        except Exception as e:
            print(f"初始化Excel失败: {str(e)}")
    
    def get_all_records(self):
        """获取所有记录"""
        try:
            if os.path.exists(self.excel_path):
                df = pd.read_excel(self.excel_path)
                return df.to_dict('records')
            return []
        except Exception as e:
            print(f"获取记录失败: {str(e)}")
            return []
    
    def record_exists(self, receive_time, sender, subject):
        """检查记录是否已存在"""
        try:
            if not os.path.exists(self.excel_path):
                return False
            
            df = pd.read_excel(self.excel_path)
            if len(df) == 0:
                return False
            
            mask = (
                (df['邮件接收时间'] == receive_time) &
                (df['发件人'] == sender) &
                (df['邮件标题'] == subject)
            )
            return mask.any()
        except Exception as e:
            print(f"检查记录异常: {str(e)}")
            return False
    
    def add_record(self, record_data):
        """添加新记录"""
        try:
            df = pd.read_excel(self.excel_path) if os.path.exists(self.excel_path) else pd.DataFrame(columns=self.headers)
            
            next_row = len(df) + 1
            row_data = {
                '序号': next_row,
                '登记时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                '投保状态': '未投保',
                '邮件回复状态': '未回复'
            }
            row_data.update(record_data)
            
            df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
            df.to_excel(self.excel_path, index=False)
            print(f"添加记录成功: 行{next_row}")
            return True
        except Exception as e:
            print(f"添加记录异常: {str(e)}")
            return False
    
    def update_record(self, search_field, search_value, update_fields):
        """更新记录"""
        try:
            df = pd.read_excel(self.excel_path)
            mask = df[search_field] == search_value
            
            if mask.any():
                for field, value in update_fields.items():
                    df.loc[mask, field] = value
                df.to_excel(self.excel_path, index=False)
                print(f"更新记录成功: {search_field}={search_value}")
                return True
            else:
                print(f"未找到匹配记录: {search_field}={search_value}")
                return False
        except Exception as e:
            print(f"更新记录异常: {str(e)}")
            return False
    
    def get_uninsured_records(self):
        """获取未投保记录"""
        try:
            if not os.path.exists(self.excel_path):
                print("投保记录文件不存在")
                return []
            
            df = pd.read_excel(self.excel_path, engine='openpyxl')
            
            if len(df) == 0:
                print("没有记录")
                return []
            
            # 筛L列（投保状态）为"未投保"的记录
            uninsured_df = df[df['投保状态'] == '未投保']
            
            # 转换为字典列表，每个字典可以通过表头名获取值
            records = uninsured_df.to_dict('records')
            
            print(f"找到 {len(records)} 条未投保记录")
            return records
            
        except Exception as e:
            print(f"获取未投保记录失败: {str(e)}")
            return []
    
    def get_records_by_status(self, status_field, status_value):
        """根据状态获取记录"""
        try:
            df = pd.read_excel(self.excel_path)
            filtered = df[df[status_field] == status_value]
            return filtered.to_dict('records')
        except Exception as e:
            print(f"按状态查询失败: {str(e)}")
            return []
