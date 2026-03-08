#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
读取已投保记录
read_insured_records.py
"""

# import xbot
# from xbot import print, sleep
import pandas as pd
import os
import sys
import io
from config import FILE_CONFIG


class InsuredRecordReader:
    """已投保记录读取器"""

    def __init__(self, filename=None):
        filename = filename or FILE_CONFIG.get('record_filename')
        if filename:
            self.excel_path = os.path.join(os.path.dirname(FILE_CONFIG['excel_path']), filename)
        else:
            self.excel_path = FILE_CONFIG['excel_path']

    def get_insured_records(self):
        """
        获取已投保记录
        返回: List[Dict]
        """
        try:
            if not os.path.exists(self.excel_path):
                print(f"文件不存在: {self.excel_path}")
                return []

            df = pd.read_excel(self.excel_path, engine='openpyxl')
            df = df.fillna('')
            if '提单号' in df.columns:
                df['提单号'] = df['提单号'].apply(lambda x: str(int(float(x))) if x != '' else '')

            if len(df) == 0:
                print("Excel文件中没有记录")
                return []

            insured_df = df[df['投保状态'] == '已投保']
            records = insured_df.to_dict('records')
            print(f"找到 {len(records)} 条已投保记录")
            return records

        except Exception as e:
            print(f"读取已投保记录失败: {str(e)}")
            return []


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    filename = sys.argv[1] if len(sys.argv) > 1 else None
    reader = InsuredRecordReader(filename=filename)
    print(f"读取文件: {reader.excel_path}\n")

    records = reader.get_insured_records()
    if not records:
        print("没有已投保记录")
        return

    for i, record in enumerate(records, 1):
        print(f"记录 {i}:")
        print(f"  序号: {record.get('序号', '')}")
        print(f"  邮件标题: {record.get('邮件标题', '')}")
        print(f"  投保类型: {record.get('投保类型', '')}")
        print(f"  投保状态: {record.get('投保状态', '')}")
        print(f"  保单号: {record.get('保单号', '')}")
        print(f"  被保险人: {record.get('被保险人', '')}")
        print(f"  货物名称: {record.get('货物名称', '')}")
        print(f"  起运地: {record.get('起运地', '')}")
        print(f"  目的地: {record.get('目的地', '')}")
        print(f"  发票金额: {record.get('发票金额', '')}")
        print(f"  发票币种: {record.get('发票币种', '')}")
        print()

    return records


if __name__ == "__main__":
    main()
