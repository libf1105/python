# 主函数
# main.py
"""
主程序入口 - 用于调度四个应用
"""

import xbot
from xbot import print, sleep
import sys
import argparse
from datetime import datetime

from mail_parser import main as mail_parser_main
from insurance_processor import main as insurance_processor_main
from policy_downloader import main as policy_downloader_main
from mail_replier import main as mail_replier_main

def run_application(app_name):
    """运行指定应用"""
    print(f"开始执行应用: {app_name}")
    
    try:
        if app_name == 'mail_parser':
            result = mail_parser_main()
        elif app_name == 'insurance_processor':
            result = insurance_processor_main()
        elif app_name == 'policy_downloader':
            result = policy_downloader_main()
        elif app_name == 'mail_replier':
            result = mail_replier_main()
        else:
            return {"status": "error", "message": f"未知应用: {app_name}"}
        
        print(f"应用执行结果: {result}")
        return result
        
    except Exception as e:
        error_msg = f"应用执行异常: {str(e)}"
        print(error_msg)
        return {"status": "error", "message": error_msg}

def run_all_applications():
    """按顺序运行所有应用"""
    print("=" * 50)
    print(f"开始执行全流程 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)
    
    results = {}
    
    # 1. 邮件解析
    print("\n[阶段1] 邮件解析与登记")
    results['mail_parser'] = run_application('mail_parser')
    sleep(2)
    
    # 2. 投保操作
    print("\n[阶段2] 投保操作")
    results['insurance_processor'] = run_application('insurance_processor')
    sleep(2)
    
    # 3. 保单下载
    print("\n[阶段3] 保单下载")
    results['policy_downloader'] = run_application('policy_downloader')
    sleep(2)
    
    # 4. 邮件回复
    print("\n[阶段4] 邮件回复")
    results['mail_replier'] = run_application('mail_replier')
    
    # 汇总结果
    print("\n" + "=" * 50)
    print("执行结果汇总:")
    print("=" * 50)
    
    success_count = 0
    for app_name, result in results.items():
        status = result.get('status', 'unknown')
        if status == 'success':
            success_count += 1
        print(f"{app_name}: {status}")
    
    print(f"\n成功: {success_count}/4")
    
    if success_count == 4:
        return {"status": "success", "message": "全流程执行完成"}
    else:
        return {"status": "partial_success", "message": f"部分成功 ({success_count}/4)"}

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='保险自动化处理系统')
    parser.add_argument('--app', choices=['all', 'mail_parser', 'insurance_processor', 
                                         'policy_downloader', 'mail_replier'],
                       default='all', help='选择要运行的应用')
    
    args = parser.parse_args()
    
    if args.app == 'all':
        result = run_all_applications()
    else:
        result = run_application(args.app)
    
    print(f"\n最终结果: {result}")
    return result

if __name__ == "__main__":
    # 在影刀中，可以通过命令行参数指定运行哪个应用
    # 例如: python main.py --app mail_parser
    result = main()
    
    # 返回结果给影刀
    if result.get('status') == 'success':
        sys.exit(0)
    else:
        sys.exit(1)
