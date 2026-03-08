# 测试用方法
# 将邮箱中的3封邮件标记为未读状态，方便重新进行邮件读取和解析


import imaplib
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

mail = imaplib.IMAP4_SSL('imap.qiye.163.com', 993)
mail.login('libingfeng@jctrans.net', '2LnX!cZ@BFe8.*n')
mail.select('"INBOX/04 &T92WaU4aUqGQ6A-"')

# 获取所有邮件
result, data = mail.search(None, 'ALL')
all_uids = data[0].split()
print(f'总邮件数: {len(all_uids)}')

# 标记最近3封为未读
if len(all_uids) >= 3:
    for uid in all_uids[-3:]:
        mail.store(uid, '-FLAGS', '\\Seen')
    print('已将最近3封邮件标记为未读')

# 检查未读邮件
result, data = mail.search(None, 'UNSEEN')
unseen = data[0].split()
print(f'未读邮件数: {len(unseen)}')

mail.logout()
