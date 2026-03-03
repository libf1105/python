# 配置文件
# config.py
"""
配置文件 - 所有配置参数集中管理
"""

# 邮箱配置
EMAIL_CONFIG = {
    'imap_server': 'imap.qiye.163.com',  # IMAP服务器
    'imap_port': 993,                     # IMAP端口
    'smtp_server': 'smtp.qiye.163.com',  # SMTP服务器
    'smtp_port': 465,                     # SMTP端口
    'username': 'libingfeng@jctrans.net', # 邮箱账号
    'password': '2LnX!cZ@BFe8.*n',          # 邮箱密码/授权码
    'folder': 'insurance',               # 监控的文件夹' 
}

# 投保系统配置
INSURANCE_SYSTEMS = {
    '快递': {
        'login_url': ' https://express-insurance-system.com/login',
        'submit_url': ' https://express-insurance-system.com/submit',
        'query_url': ' https://express-insurance-system.com/query',
        'username': 'express_user',
        'password': 'express_pass',
    },
    '非快递': {
        'login_url': ' https://normal-insurance-system.com/login',
        'submit_url': ' https://normal-insurance-system.com/submit',
        'query_url': ' https://normal-insurance-system.com/query',
        'username': 'normal_user',
        'password': 'normal_pass',
    }
}

# 文件存储配置
FILE_CONFIG = {
    'base_dir': 'D:/保险保单/',
    'temp_dir': 'D:/保险保单/临时文件/',
    'log_dir': 'D:/保险保单/日志/',
}

# 在线表格配置
ONLINE_SHEET_CONFIG = {
    'type': 'feishu',  # feishu/dingtalk/excel
    'app_id': 'your_app_id',
    'app_secret': 'your_app_secret',
    'sheet_token': 'your_sheet_token',
    'sheet_id': 'your_sheet_id',
}

# 定时任务配置
SCHEDULE_CONFIG = {
    'mail_parse_times': ['12:00', '21:00'],
    'insurance_process_times': ['13:00', '22:00'],
    'policy_download_times': ['14:00', '23:00'],
    'mail_reply_times': ['14:00', '23:00'],
}

# 关键词配置
KEYWORDS = {
    'express_keywords': ['快递', '速递', 'express', 'courier', '顺丰', '圆通', '申通', '中通', '韵达'],
    'ignore_keywords': ['测试', 'test', 'demo'],
}
