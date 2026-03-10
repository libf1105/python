#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
更新保单下载记录
update_policy_download.py
"""
# import xbot
# from xbot import print, sleep
import pandas as pd
import os
import sys
import io
from config import FILE_CONFIG


def update_policy_download_status(row_number, policy_file_path):
    """
    更新投保记录的保单文件路径和投保状态为已下载保单
    
    参数:
        row_number: int - 行号（对应Excel第一列的序号值）
        policy_file_path: str - 保单文件路径
    
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
        
        # 读取Excel文件，所有数据按字符串处理
        df = pd.read_excel(excel_path, engine='openpyxl', dtype=str)
        df = df.fillna('')
        
        if len(df) == 0:
            print("Excel文件中没有记录")
            return False
        
        # 根据序号找到对应的行
        mask = df['序号'] == str(row_number)
        
        if not mask.any():
            print(f"未找到序号为 {row_number} 的记录")
            return False
        
        # 更新保单文件路径和投保状态
        df.loc[mask, '保单文件路径'] = policy_file_path
        df.loc[mask, '投保状态'] = '已下载保单'
        
        # 保存回Excel文件
        df.to_excel(excel_path, index=False, engine='openpyxl')
        
        print(f"更新成功: 序号 {row_number} -> 保单文件路径 {policy_file_path}, 状态: 已下载保单")
        return True
        
    except Exception as e:
        print(f"更新失败: {str(e)}")
        return False


def main():
    """主函数"""
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    if len(sys.argv) != 3:
        print("用法: python update_policy_download.py <行号> <保单文件路径>")
        return
    
    try:
        row_number = int(sys.argv[1])
        policy_file_path = sys.argv[2]
        
        success = update_policy_download_status(row_number, policy_file_path)
        if success:
            print("保单下载状态更新成功")
        else:
            print("保单下载状态更新失败")
            
    except ValueError:
        print("错误: 行号必须是数字")
    except Exception as e:
        print(f"执行失败: {str(e)}")


if __name__ == "__main__":
    main()