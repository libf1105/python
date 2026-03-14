# 配置文件
# config.py
"""
配置文件 - 所有配置参数集中管理
"""

# 文件存储配置
FILE_CONFIG = {
    'base_dir': 'D:/自动投保/',
    'temp_dir': 'D:/自动投保/临时文件/',
    'log_dir': 'D:/自动投保/日志/',
    'excel_path': 'D:/自动投保/投保记录.xlsx',
    'attachment_dir': 'D:/自动投保/邮件附件/',
    'record_filename': '投保记录_20260314.xlsx',  # 指定要读取的投保记录文件名
}


EMAIL_CONFIG = {
    'imap_server': 'imap.qiye.163.com',  # IMAP服务器
    'imap_port': 993,                     # IMAP端口
    'smtp_server': 'smtp.qiye.163.com',  # SMTP服务器
    'smtp_port': 465,                     # SMTP端口
    'username': 'libingfeng@jctrans.net', # 邮箱账号
    'password': 'vh9w#4tw3$qd@B4G',          # 邮箱密码/授权码
    'folder': '"INBOX/04 &T92WaU4aUqGQ6A-"',               # 监控的文件夹' 
    'reply_body': '''尊敬的客户，

您好！

您的货物保险已成功投保，保单文件请见附件。

如有任何疑问，请随时联系我们。

谢谢！

此致
敬礼！''',  # 回复邮件正文
}

# 投保系统配置
INSURANCE_SYSTEMS = {
    '快递': {
        # 'login_url': ' https://express-insurance-system.com/login',
        'submit_url': ' http://www.moubaoins.com/home',
        'query_url': ' http://www.moubaoins.com/home',
        # 'username': 'express_user',
        # 'password': 'express_pass',
    },
    '非快递': {
        # 'login_url': ' https://normal-insurance-system.com/login',
        'submit_url': ' https://www.yunjiins.cn/policies/insure/intl',
        'query_url': ' https://www.yunjiins.cn/policies/intl',
        # 'username': 'normal_user',
        # 'password': 'normal_pass',
    }
}


# 关键词配置
KEYWORDS = {
    'express_keywords': ['快递', '速递', 'express', 'courier', '顺丰', '圆通', '申通', '中通', '韵达'],
    'ignore_keywords': ['测试', 'test', 'demo'],
    'fragile_keywords': ['易碎', '玻璃', '陶瓷', '陶器', '瓷器', '镜子', '镜面', '灯具', '灯饰', '工艺品', '陶艺', 'glass', 'ceramic', 'fragile', 'mirror'],
    'dangerous_keywords': ['危险品', '电池', '磁铁', '磁性', '化学品', '液体', '气体', '易燃', '易爆', '腐蚀', '有毒', '放射', 'battery', 'magnet', 'chemical', 'liquid', 'gas', 'flammable', 'explosive', 'corrosive', 'toxic', 'radioactive']
}
