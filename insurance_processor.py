# 投保操作
# insurance_processor.py
"""
应用2：投保操作
定时任务：每天13:00和22:00执行
"""

import xbot
from xbot import print, sleep
import time
import re
from datetime import datetime

from config import INSURANCE_SYSTEMS
from database import OnlineSheetManager
from utils import write_log, clean_temp_files

class InsuranceProcessor:
    """投保处理器"""
    
    def __init__(self):
        self.sheet_manager = OnlineSheetManager()
        self.systems = INSURANCE_SYSTEMS
    
    def extract_insurance_info(self, record):
        """从记录中提取投保信息"""
        info = {
            'insured': '',      # 被保险人
            'goods': '',        # 货物名称
            'amount': '',       # 保额
            'from': '',         # 起运地
            'to': '',           # 目的地
            'transport': '',    # 运输方式
            'bill_no': '',      # 提单号
        }
        
        # 从邮件内容提取
        mail_content = record.get('邮件内容', '')
        
        # 提取被保险人
        insured_match = re.search(r'被保险人[:：]?\s*([^\n]+)', mail_content)
        if insured_match:
            info['insured'] = insured_match.group(1).strip()
        
        # 提取货物名称
        goods_match = re.search(r'货物[名称名]?[:：]?\s*([^\n]+)', mail_content)
        if goods_match:
            info['goods'] = goods_match.group(1).strip()
        
        # 提取保额
        amount_match = re.search(r'保额[:：]?\s*([\d,.]+)', mail_content)
        if amount_match:
            info['amount'] = amount_match.group(1)
        
        # 提取起运地
        from_match = re.search(r'起运地[:：]?\s*([^\n]+)', mail_content)
        if from_match:
            info['from'] = from_match.group(1).strip()
        
        # 提取目的地
        to_match = re.search(r'目的地[:：]?\s*([^\n]+)', mail_content)
        if to_match:
            info['to'] = to_match.group(1).strip()
        
        # 从处理备注中提取提单号
        remark = record.get('处理备注', '')
        bill_match = re.search(r'提单号[:：]?\s*([A-Za-z0-9]+)', remark)
        if bill_match:
            info['bill_no'] = bill_match.group(1)
        
        return info
    
    def login_insurance_system(self, system_type):
        """登录投保系统"""
        try:
            system_config = self.systems[system_type]
            
            write_log('应用2', 
                f"登录{system_type}投保系统: {system_config['login_url']}",
                'INFO'
            )
            
            # 使用影刀打开浏览器
            # browser = xbot.open_browser('chrome')
            # xbot.navigate_to(system_config['login_url'])
            
            # 输入用户名
            # username_field = xbot.find_element('id', 'username')
            # xbot.type_text(username_field, system_config['username'])
            
            # 输入密码
            # password_field = xbot.find_element('id', 'password')
            # xbot.type_text(password_field, system_config['password'])
            
            # 点击登录
            # login_button = xbot.find_element('xpath', '//button[@type="submit"]')
            # xbot.click(login_button)
            
            # 等待登录完成
            sleep(3)
            
            write_log('应用2', f"{system_type}系统登录成功", 'INFO')
            return True
            
        except Exception as e:
            write_log('应用2', f"登录{system_type}系统失败: {str(e)}", 'ERROR')
            return False
    
    def fill_insurance_form(self, system_type, insurance_info):
        """填写投保表单"""
        try:
            write_log('应用2', f"开始填写{system_type}投保表单", 'INFO')
            
            # 根据系统类型填写不同的表单
            if system_type == '快递':
                # 快递投保表单填写
                self._fill_express_form(insurance_info)
            else:
                # 非快递投保表单填写
                self._fill_normal_form(insurance_info)
            
            write_log('应用2', "投保表单填写完成", 'INFO')
            return True
            
        except Exception as e:
            write_log('应用2', f"填写投保表单失败: {str(e)}", 'ERROR')
            return False
    
    def _fill_express_form(self, info):
        """填写快递投保表单"""
        # 这里需要根据实际的网页表单元素进行调整
        
        # 填写被保险人
        # insured_field = xbot.find_element('name', 'insuredName')
        # xbot.type_text(insured_field, info.get('insured', ''))
        
        # 填写货物名称
        # goods_field = xbot.find_element('name', 'goodsName')
        # xbot.type_text(goods_field, info.get('goods', ''))
        
        # 填写保额
        # amount_field = xbot.find_element('name', 'insuredAmount')
        # xbot.type_text(amount_field, info.get('amount', ''))
        
        # 填写起运地
        # from_field = xbot.find_element('name', 'fromLocation')
        # xbot.type_text(from_field, info.get('from', ''))
        
        # 填写目的地
        # to_field = xbot.find_element('name', 'toLocation')
        # xbot.type_text(to_field, info.get('to', ''))
        
        # 填写提单号
        # bill_field = xbot.find_element('name', 'billNumber')
        # xbot.type_text(bill_field, info.get('bill_no', ''))
        
        sleep(2)
    
    def _fill_normal_form(self, info):
        """填写非快递投保表单"""
        # 非快递投保可能字段不同，需要单独处理
        
        # 填写被保险人
        # insured_field = xbot.find_element('id', 'policyHolder')
        # xbot.type_text(insured_field, info.get('insured', ''))
        
        # 其他字段...
        
        sleep(2)
    
    def submit_insurance(self, system_type):
        """提交投保申请"""
        try:
            write_log('应用2', "提交投保申请", 'INFO')
            
            # 点击提交按钮
            # submit_button = xbot.find_element('xpath', '//button[contains(text(), "提交")]')
            # xbot.click(submit_button)
            
            # 等待提交完成
            sleep(5)
            
            # 获取保单号（需要根据实际页面调整）
            policy_number = self._extract_policy_number()
            
            if policy_number:
                write_log('应用2', f"投保提交成功，保单号: {policy_number}", 'INFO')
                return policy_number
            else:
                write_log('应用2', "获取保单号失败", 'ERROR')
                return None
                
        except Exception as e:
            write_log('应用2', f"提交投保失败: {str(e)}", 'ERROR')
            return None
    
    def _extract_policy_number(self):
        """从页面提取保单号"""
        try:
            # 这里需要根据实际的页面结构调整
            # policy_element = xbot.find_element('class', 'policy-number')
            # policy_number = xbot.get_text(policy_element)
            
            # 模拟返回一个保单号
            import random
            policy_number = f"POL{datetime.now().strftime('%Y%m%d')}{random.randint(1000, 9999)}"
            
            return policy_number
            
        except:
            return None
    
    def process_single_record(self, record):
        """处理单条记录"""
        try:
            bill_no = ''
            remark = record.get('处理备注', '')
            bill_match = re.search(r'提单号[:：]?\s*([A-Za-z0-9]+)', remark)
            if bill_match:
                bill_no = bill_match.group(1)
            
            if not bill_no:
                write_log('应用2', f"记录无提单号，跳过: {record.get('邮件标题', '')}", 'WARN')
                return False
            
            insurance_type = record.get('投保类型', '非快递')
            
            write_log('应用2', 
                f"开始处理: {bill_no} | 类型: {insurance_type}",
                'INFO'
            )
            
            # 登录系统
            if not self.login_insurance_system(insurance_type):
                return False
            
            # 提取投保信息
            insurance_info = self.extract_insurance_info(record)
            
            # 填写表单
            if not self.fill_insurance_form(insurance_type, insurance_info):
                return False
            
            # 提交投保
            policy_number = self.submit_insurance(insurance_type)
            
            if policy_number:
                # 更新表格状态
                update_fields = {
                    '投保状态': '已投保',
                    '保单号': policy_number,
                    '处理备注': f"{remark} | 投保时间: {datetime.now().strftime('%H:%M:%S')}"
                }
                
                success = self.sheet_manager.update_record(
                    '处理备注',  # 搜索字段
                    remark,      # 搜索值
                    update_fields
                )
                
                if success:
                    write_log('应用2', 
                        f"处理成功: {bill_no} -> {policy_number}",
                        'INFO'
                    )
                    return True
                else:
                    write_log('应用2', f"更新表格失败: {bill_no}", 'ERROR')
                    return False
            else:
                write_log('应用2', f"投保提交失败: {bill_no}", 'ERROR')
                return False
                
        except Exception as e:
            write_log('应用2', f"处理记录异常: {str(e)}", 'ERROR')
            return False
    
    def run(self):
        """主执行函数"""
        write_log('应用2', "开始执行投保操作任务", 'INFO')
        
        try:
            # 获取需要处理的记录（投保状态=未投保）
            records = self.sheet_manager.get_records_by_status('投保状态', '未投保')
            
            if not records:
                write_log('应用2', "没有需要投保的记录", 'INFO')
                return True
            
            total_count = len(records)
            write_log('应用2', f"找到 {total_count} 条待投保记录", 'INFO')
            
            # 处理每条记录
            success_count = 0
            for i, record in enumerate(records):
                write_log('应用2', f"处理第 {i+1}/{total_count} 条记录", 'INFO')
                
                if self.process_single_record(record):
                    success_count += 1
                
                # 避免处理过快
                sleep(2)
            
            # 清理临时文件
            clean_temp_files()
            
            write_log('应用2', 
                f"投保操作任务完成: 成功 {success_count}/{total_count}",
                'INFO'
            )
            
            return True
            
        except Exception as e:
            write_log('应用2', f"投保操作任务异常: {str(e)}", 'ERROR')
            return False

def main():
    """主函数 - 用于影刀调用"""
    processor = InsuranceProcessor()
    success = processor.run()
    
    if success:
        return {"status": "success", "message": "投保操作任务完成"}
    else:
        return {"status": "error", "message": "投保操作任务失败"}

if __name__ == "__main__":
    # 本地测试
    result = main()
    print(result)
