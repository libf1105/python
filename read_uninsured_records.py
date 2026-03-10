#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
读取未投保记录
read_uninsured_records.py
"""
# import xbot
# from xbot import print, sleep
import pandas as pd
import os
import sys
import io
from datetime import datetime


from config import FILE_CONFIG


class UninsuredRecordReader:
    """未投保记录读取器"""
    
    def __init__(self, filename=None):
        """
        初始化读取器
        参数:
            filename: 指定要读取的文件名（不含路径），默认从 FILE_CONFIG['record_filename'] 读取
        """
        filename = filename or FILE_CONFIG.get('record_filename')
        if filename:
            self.excel_path = os.path.join(os.path.dirname(FILE_CONFIG['excel_path']), filename)
        else:
            self.excel_path = FILE_CONFIG['excel_path']
    
    def get_uninsured_records(self):
        """
        获取未投保记录
        返回: List[Dict] - 每个字典代表一行记录，可通过表头名访问单元格内容
        """
        try:
            if not os.path.exists(self.excel_path):
                print(f"文件不存在: {self.excel_path}")
                return []
            
            # 读取Excel文件，将NaN转换为空字符串
            df = pd.read_excel(self.excel_path, engine='openpyxl')
            df = df.fillna('')

            if len(df) == 0:
                print("Excel文件中没有记录")
                return []
            
            # 筛选投保状态为"未投保"的记录
            uninsured_df = df[df['投保状态'] == '未投保']
            
            # 转换为字典列表
            records = uninsured_df.to_dict('records')
            
            print(f"找到 {len(records)} 条未投保记录")
            return records
            
        except Exception as e:
            print(f"读取未投保记录失败: {str(e)}")
            return []


def main():
    """主函数"""
    # 设置UTF-8编码输出
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    print("=" * 60)
    print("读取未投保记录")
    print("=" * 60)
    
    # 创建读取器实例
    # 可以指定文件名：reader = UninsuredRecordReader(filename="投保记录_20060308.xlsx")
    # 或使用当天日期：reader = UninsuredRecordReader()
    filename = sys.argv[1] if len(sys.argv) > 1 else None
    reader = UninsuredRecordReader(filename=filename)
    print(f"读取文件: {reader.excel_path}\n")
    
    # 获取未投保记录
    records = reader.get_uninsured_records()
    
    if not records:
        print("\n没有未投保记录")
        return
    
    print(f"\n=== 未投保记录列表 ===")
    print(f"总数: {len(records)}\n")
    
    # 显示每条记录的关键信息
    for i, record in enumerate(records, 1):
        print(f"记录 {i}:")
        print(f"  序号: {record.get('序号', '')}")
        print(f"  邮件标题: {record.get('邮件标题', '')}")
        print(f"  投保类型: {record.get('投保类型', '')}")
        print(f"  投保状态: {record.get('投保状态', '')}")
        print(f"  被保险人: {record.get('被保险人', '')}")
        print(f"  货物名称: {record.get('货物名称', '')}")
        print(f"  起运地: {record.get('起运地', '')}")
        print(f"  目的地: {record.get('目的地', '')}")
        print(f"  发票金额: {record.get('发票金额', '')}")
        print(f"  发票币种: {record.get('发票币种', '')}")
        print()
    
    # 演示如何访问单个记录的字段
    print("=" * 60)
    print("=== 字段访问示例 ===")
    first_record = records[0]
    print(f"第一条记录的被保险人: {first_record['被保险人']}")
    print(f"第一条记录的发票金额: {first_record['发票金额']}")
    print(f"第一条记录的运输方式: {first_record['运输方式']}")
    print(f"第一条记录的提单号: {first_record['提单号']}")
    print("=" * 60)
    
    return records


if __name__ == "__main__":
    main()
