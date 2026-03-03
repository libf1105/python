# 保单下载
# policy_downloader.py
"""
应用3：保单下载
定时任务：每天14:00和23:00执行
"""

import xbot
from xbot import print, sleep
import os
import time
from datetime import datetime

from config import INSURANCE_SYSTEMS, FILE_CONFIG
from database import OnlineSheetManager
from utils import (
    write_log,
    get_date_folder_path,
    generate_policy_filename,
    clean_temp_files
)

class PolicyDownloader:
    """保单下载器"""
    
    def __init__(self):
        self.sheet_manager = OnlineSheetManager()
        self.systems = INSURANCE_SYSTEMS
    
    def login_to_query_system(self, system_type):
        """登录查询系统"""
        try:
            system_config = self.systems[system_type]
            
            write_log('应用3', 
                f"登录{system_type}查询系统: {system_config['query_url']}",
                'INFO'
            )
            
            # 使用影刀打开浏览器
            # browser = xbot.open_browser('chrome')
            # xbot.navigate_to(system_config['query_url'])
            
            # 输入用户名密码（如果和投保系统不同）
            # ...
            
            write_log('应用3', f"{system_type}查询系统登录成功", 'INFO')
            return True
            
        except Exception as e:
            write_log('应用3', f"登录查询系统失败: {str(e)}", 'ERROR')
            return False
    
    def query_policy_status(self, policy_number, system_type):
        """查询保单状态"""
        try:
            write_log('应用3', f"查询保单状态: {policy_number}", 'INFO')
            
            # 在查询页面输入保单号
            # policy_input = xbot.find_element('id', 'policyNumber')
            # xbot.type_text(policy_input, policy_number)
            
            # 点击查询按钮
            # query_button = xbot.find_element('xpath', '//button[contains(text(), "查询")]')
            # xbot.click(query_button)
            
            # 等待查询结果
            sleep(3)
            
            # 获取状态（需要根据实际页面调整）
            status = self._extract_policy_status()
            
            if status:
                write_log('应用3', f"保单状态: {policy_number} -> {status}", 'INFO')
                return status
            else:
                write_log('应用3', f"获取保单状态失败: {policy_number}", 'WARN')
                return None
                
        except Exception as e:
            write_log('应用3', f"查询保单状态异常: {str(e)}", 'ERROR')
            return None
    
    def _extract_policy_status(self):
        """从页面提取保单状态"""
        try:
            # 这里需要根据实际的页面结构调整
            # status_element = xbot.find_element('class', 'policy-status')
            # status_text = xbot.get_text(status_element)
            
            # 模拟返回状态
            status_options = ['已生效', '处理中', '已拒绝', '已取消']
            import random
            status = random.choice(status_options)
            
            return status
            
        except:
            return None
    
    def download_policy_file(self, policy_number, system_type):
        """下载保单文件"""
        try:
            write_log('应用3', f"下载保单文件: {policy_number}", 'INFO')
            
            # 点击下载按钮
            # download_button = xbot.find_element('xpath', '//button[contains(text(), "下载")]')
            # xbot.click(download_button)
            
            # 等待下载完成
            sleep(5)
            
            # 获取下载的文件（需要根据浏览器下载设置调整）
            downloaded_file = self._get_downloaded_file(policy_number)
            
            if downloaded_file and os.path.exists(downloaded_file):
                # 移动到指定目录
                target_dir = get_date_folder_path()
                filename = generate_policy_filename(policy_number)
                target_path = os.path.join(target_dir, filename)
                
                # 移动文件
                import shutil
                shutil.move(downloaded_file, target_path)
                
                write_log('应用3', f"保单文件保存到: {target_path}", 'INFO')
                return target_path
            else:
                write_log('应用3', f"下载文件失败: {policy_number}", 'ERROR')
                return None
                
        except Exception as e:
            write_log('应用3', f"下载保单文件异常: {str(e)}", 'ERROR')
            return None
    
    def _get_downloaded_file(self, policy_number):
        """获取下载的文件路径"""
        # 这里需要根据实际的下载目录和文件名模式调整
        download_dir = "C:/Users/用户名/Downloads/"  # 默认下载目录
        
        # 查找最近下载的文件
        import glob
        files = glob.glob(f"{download_dir}/*.pdf")
        files.sort(key=os.path.getmtime, reverse=True)
        
        if files:
            return files
        return None
    
    def process_single_record(self, record):
        """处理单条记录"""
        try:
            policy_number = record.get('保单号', '')
            if not policy_number:
                write_log('应用3', "记录无保单号，跳过", 'WARN')
                return False
            
            insurance_type = record.get('投保类型', '非快递')
            
            write_log('应用3', 
                f"开始处理保单: {policy_number} | 类型: {insurance_type}",
                'INFO'
            )
            
            # 登录查询系统
            if not self.login_to_query_system(insurance_type):
                return False
            
            # 查询保单状态
            status = self.query_policy_status(policy_number, insurance_type)
            
            if status == '已生效':
                # 下载保单文件
                file_path = self.download_policy_file(policy_number, insurance_type)
                
                if file_path:
                    # 更新表格状态
                    update_fields = {
                        '投保状态': '已生效',
                        '保单生效时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        '保单文件路径': file_path,
                        '处理备注': f"{record.get('处理备注', '')} | 生效时间: {datetime.now().strftime('%H:%M:%S')}"
                    }
                    
                    success = self.sheet_manager.update_record(
                        '保单号',  # 搜索字段
                        policy_number,  # 搜索值
                        update_fields
                    )
                    
                    if success:
                        write_log('应用3', 
                            f"保单处理成功: {policy_number} -> 已生效",
                            'INFO'
                        )
                        return True
                    else:
                        write_log('应用3', f"更新表格失败: {policy_number}", 'ERROR')
                        return False
                else:
                    write_log('应用3', f"下载保单失败: {policy_number}", 'ERROR')
                    return False
            else:
                write_log('应用3', 
                    f"保单未生效: {policy_number} -> {status}",
                    'INFO'
                )
                return False  # 未生效，下次再处理
                
        except Exception as e:
            write_log('应用3', f"处理保单异常: {str(e)}", 'ERROR')
            return False
    
    def run(self):
        """主执行函数"""
        write_log('应用3', "开始执行保单下载任务", 'INFO')
        
        try:
            # 获取需要处理的记录（投保状态=已投保）
            records = self.sheet_manager.get_records_by_status('投保状态', '已投保')
            
            if not records:
                write_log('应用3', "没有需要下载的保单", 'INFO')
                return True
            
            total_count = len(records)
            write_log('应用3', f"找到 {total_count} 条待下载保单", 'INFO')
            
            # 处理每条记录
            success_count = 0
            for i, record in enumerate(records):
                write_log('应用3', f"处理第 {i+1}/{total_count} 条保单", 'INFO')
                
                if self.process_single_record(record):
                    success_count += 1
                
                # 避免处理过快
                sleep(2)
            
            # 清理临时文件
            clean_temp_files()
            
            write_log('应用3', 
                f"保单下载任务完成: 成功 {success_count}/{total_count}",
                'INFO'
            )
            
            return True
            
        except Exception as e:
            write_log('应用3', f"保单下载任务异常: {str(e)}", 'ERROR')
            return False

def main():
    """主函数 - 用于影刀调用"""
    downloader = PolicyDownloader()
    success = downloader.run()
    
    if success:
        return {"status": "success", "message": "保单下载任务完成"}
    else:
        return {"status": "error", "message": "保单下载任务失败"}

if __name__ == "__main__":
    # 本地测试
    result = main()
    print(result)

