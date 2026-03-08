# 工具函数
# utils.py
"""
工具函数模块
"""

# import xbot
# from xbot import print, sleep
import re
import os
import shutil
from datetime import datetime, timedelta
from config import KEYWORDS, FILE_CONFIG

def create_directory_structure():
    """创建目录结构"""
    directories = [
        FILE_CONFIG['base_dir'],
        FILE_CONFIG['temp_dir'],
        FILE_CONFIG['log_dir'],
        f"{FILE_CONFIG['base_dir']}保单文件/",
        f"{FILE_CONFIG['base_dir']}备份/",
        f"{FILE_CONFIG['base_dir']}日志/应用1/",
        f"{FILE_CONFIG['base_dir']}日志/应用2/",
        f"{FILE_CONFIG['base_dir']}日志/应用3/",
        f"{FILE_CONFIG['base_dir']}日志/应用4/",
    ]
    
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"创建目录: {directory}")

def analyze_goods_type(goods_description):
    """分析货物类型"""
    if not goods_description:
        return '普货'
    
    content = str(goods_description).lower()
    
    # 检查危险品
    for keyword in KEYWORDS['dangerous_keywords']:
        if keyword.lower() in content:
            return '危险品'
    
    # 检查易碎品
    for keyword in KEYWORDS['fragile_keywords']:
        if keyword.lower() in content:
            return '易碎品'
    
    return '普货'

def analyze_import_export_type(origin_country):
    """分析进出口类型"""
    if not origin_country:
        return '出口'
    
    content = str(origin_country).lower()
    # 判断是否包含中国
    if '中国' in content or 'china' in content or 'cn' in content:
        return '出口'
    else:
        return '进口'

def analyze_insurance_type(mail_content, mail_subject):
    """分析投保类型"""
    content = f"{mail_content} {mail_subject}".lower()
    
    # 检查是否包含快递关键词
    for keyword in KEYWORDS['express_keywords']:
        if keyword.lower() in content:
            return '快递'
    
    return '非快递'

def extract_bill_number(mail_content):
    """从邮件内容提取提单号"""
    # 常见的提单号格式：8-12位数字或字母数字组合
    patterns = [
        r'提单[号碼码][:：]?\s*([A-Za-z0-9]{8,12})',
        r'B/L[：:]?\s*([A-Za-z0-9]{8,12})',
        r'提单[号碼码][:：]?\s*(\d{8,12})',
        r'([A-Z]{3}\d{8})',  # 如：ABC12345678
        r'(\d{8,12})',       # 纯数字
    ]
    
    for pattern in patterns:
        match = re.search(pattern, mail_content)
        if match:
            return match.group(1)
    
    return ''

def _to_str(val):
    """将单元格值转为去除首尾空格的字符串"""
    if val is None or val == '':
        return ''
    return str(val).strip()


def _to_amount(val):
    """将单元格值转为数字（发票金额专用），无法转换则返回空字符串"""
    if val is None or val == '':
        return ''
    try:
        return float(val)
    except (ValueError, TypeError):
        return ''


def parse_excel_attachment(file_path):
    """解析Excel附件（UTF-8编码）"""
    try:
        import pandas as pd
        
        # 读取所有工作表
        excel_data = {}
        
        # 1. 读取主投保单信息（第一个工作表）
        df_main = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl').fillna('')
        
        # 基于模板结构解析关键字段
        insured_info = {}
        transport_info = {}
        goods_info = {}
        invoice_info = {}
        
        # 解析被保险人信息（行5-10）
        insured_info['被保险人'] = _to_str(df_main.iloc[8, 1] if len(df_main) > 8 else '')
        insured_info['联系电话'] = _to_str(df_main.iloc[8, 5] if len(df_main) > 8 else '')
        insured_info['通讯地址'] = _to_str(df_main.iloc[9, 1] if len(df_main) > 9 else '')
        
        # 解析货物运输信息（行12-20）
        transport_info['起运地'] = _to_str(df_main.iloc[12, 1] if len(df_main) > 12 else '')
        transport_info['起运国家地区'] = _to_str(df_main.iloc[12, 3] if len(df_main) > 12 else '')
        transport_info['目的地'] = _to_str(df_main.iloc[12, 5] if len(df_main) > 12 else '')
        transport_info['国家地区'] = _to_str(df_main.iloc[12, 7] if len(df_main) > 12 else '')
        transport_info['转运地'] = _to_str(df_main.iloc[13, 1] if len(df_main) > 13 else '')
        transport_info['起运日期'] = _to_str(df_main.iloc[13, 5] if len(df_main) > 13 else '')[:10]
        transport_info['起运日期打印格式'] = _to_str(df_main.iloc[13, 7] if len(df_main) > 13 else '')
        transport_info['运输工具号'] = _to_str(df_main.iloc[14, 1] if len(df_main) > 14 else '')
        transport_info['运输方式'] = _to_str(df_main.iloc[14, 5] if len(df_main) > 14 else '')
        transport_info['装载方式'] = _to_str(df_main.iloc[14, 7] if len(df_main) > 14 else '')
        transport_info['赔款偿付地点'] = _to_str(df_main.iloc[15, 1] if len(df_main) > 15 else '')
        transport_info['贸易类型'] = _to_str(df_main.iloc[15, 5] if len(df_main) > 15 else '')
        
        # 解析货物信息（行16-17）
        goods_info['唛头'] = _to_str(df_main.iloc[18, 0] if len(df_main) > 18 else '')
        goods_info['货物描述'] = _to_str(df_main.iloc[18, 1] if len(df_main) > 18 else '')
        goods_info['包装数量'] = _to_str(df_main.iloc[18, 4] if len(df_main) > 18 else '')
        goods_info['条款'] = _to_str(df_main.iloc[19, 1] if len(df_main) > 19 else '')
        
        # 解析发票信息（行18-19）
        invoice_info['发票号'] = _to_str(df_main.iloc[20, 1] if len(df_main) > 20 else '')
        invoice_info['提单号'] = _to_str(df_main.iloc[20, 5] if len(df_main) > 20 else '')
        invoice_info['发票金额'] = _to_amount(df_main.iloc[21, 1] if len(df_main) > 21 else '')
        invoice_info['发票币种'] = _to_str(df_main.iloc[21, 5] if len(df_main) > 21 else '')
        invoice_info['工作编号'] = _to_str(df_main.iloc[22, 1] if len(df_main) > 22 else '')
        invoice_info['备注'] = _to_str(df_main.iloc[22, 5] if len(df_main) > 22 else '')
        
        excel_data['main_info'] = {
            'insured': insured_info,
            'transport': transport_info,
            'goods': goods_info,
            'invoice': invoice_info
        }
        
        # # 2. 读取投保人信息（第二个工作表）
        # try:
        #     df_applicant = pd.read_excel(file_path, sheet_name=1)
        #     if not df_applicant.empty:
        #         applicant_info = df_applicant.iloc[0].to_dict()
        #         excel_data['applicant_info'] = applicant_info
        # except:
        #     print("未找到投保人信息工作表")
        
        # # 3. 读取国家地区信息（第三个工作表，用于验证）
        # try:
        #     df_country = pd.read_excel(file_path, sheet_name=2)
        #     country_list = df_country.iloc[:, 0].dropna().tolist()
        #     excel_data['country_list'] = country_list
        # except:
        #     print("未找到国家地区信息工作表")
        
        print(f"Excel解析完成: 被保险人={insured_info['被保险人']}, 提单号={invoice_info['提单号']}")
        print(excel_data)
        return excel_data
        
    except Exception as e:
        print(f"解析Excel失败: {str(e)}")
        return None

def save_attachment(attachment_data, filename):
    """保存附件到临时目录"""
    try:
        temp_dir = FILE_CONFIG['temp_dir']
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        filepath = os.path.join(temp_dir, filename)
        
        # 如果是二进制数据
        if isinstance(attachment_data, bytes):
            with open(filepath, 'wb') as f:
                f.write(attachment_data)
        # 如果是文件路径
        elif isinstance(attachment_data, str) and os.path.exists(attachment_data):
            shutil.copy(attachment_data, filepath)
        else:
            print(f"附件数据格式不支持: {type(attachment_data)}")
            return None
        
        print(f"附件保存成功: {filepath}")
        return filepath
        
    except Exception as e:
        print(f"保存附件失败: {str(e)}")
        return None

def generate_policy_filename(policy_number, file_type='pdf'):
    """生成保单文件名"""
    date_str = datetime.now().strftime('%Y%m%d')
    timestamp = datetime.now().strftime('%H%M%S')
    return f"{policy_number}_{date_str}_{timestamp}.{file_type}"

def get_date_folder_path():
    """获取按日期组织的文件夹路径"""
    date_str = datetime.now().strftime('%Y/%m/%d')  # 年/月/日
    base_dir = f"{FILE_CONFIG['base_dir']}保单文件/"
    folder_path = os.path.join(base_dir, date_str)
    
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    return folder_path

def clean_temp_files(days=7):
    """清理临时文件（保留指定天数）"""
    try:
        temp_dir = FILE_CONFIG['temp_dir']
        cutoff_date = datetime.now() - timedelta(days=days)
        
        if not os.path.exists(temp_dir):
            return
        
        for filename in os.listdir(temp_dir):
            filepath = os.path.join(temp_dir, filename)
            if os.path.isfile(filepath):
                file_mtime = datetime.fromtimestamp(os.path.getmtime(filepath))
                if file_mtime < cutoff_date:
                    os.remove(filepath)
                    print(f"清理临时文件: {filename}")
        
        print("临时文件清理完成")
        
    except Exception as e:
        print(f"清理临时文件失败: {str(e)}")

def write_log(app_name, message, level='INFO'):
    """写入日志文件"""
    try:
        log_dir = f"{FILE_CONFIG['log_dir']}{app_name}/"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        date_str = datetime.now().strftime('%Y%m%d')
        log_file = f"{log_dir}{date_str}.log"
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(log_entry)
        
        # 同时在控制台输出
        print(f"[{app_name}] {log_entry.strip()}")
        
    except Exception as e:
        print(f"写入日志失败: {str(e)}")
