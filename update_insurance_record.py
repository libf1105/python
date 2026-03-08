#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
更新投保记录
update_insurance_record.py
"""

# import xbot
# from xbot import print, sleep
import pandas as pd
import os
import sys
import io
from config import FILE_CONFIG


def update_insurance_status(row_number, policy_number):
    """
    更新投保记录的保单号和投保状态
    
    参数:
        row_number: int - 行号（对应Excel第一列的序号值）
        policy_number: str - 保单号（流水号）
    
    返回:
        bool - 更新成功返回True，失败返回False
    """
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
        df = pd.read_excel(excel_path, engine='openpyxl')
        
        if len(df) == 0:
            print("Excel文件中没有记录")
            return False
        
        # 根据序号找到对应的行
        mask = df['序号'] == row_number
        
        if not mask.any():
            print(f"未找到序号为 {row_number} 的记录")
            return False
        
        # 确保保单号列为字符串类型
        df['保单号'] = df['保单号'].astype(str)
        
        # 更新保单号和投保状态
        df.loc[mask, '保单号'] = policy_number
        df.loc[mask, '投保状态'] = '已投保'
        
        # 保存回Excel文件
        df.to_excel(excel_path, index=False, engine='openpyxl')
        
        print(f"更新成功: 序号 {row_number} -> 保单号 {policy_number}, 状态: 已投保")
        return True
        
    except Exception as e:
        print(f"更新失败: {str(e)}")
        return False

