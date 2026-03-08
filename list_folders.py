import imaplib
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

EMAIL_CONFIG = {
    'imap_server': 'imap.qiye.163.com',
    'imap_port': 993,
    'username': 'libingfeng@jctrans.net',
    'password': '2LnX!cZ@BFe8.*n',
}

mail = imaplib.IMAP4_SSL(EMAIL_CONFIG['imap_server'], EMAIL_CONFIG['imap_port'])
mail.login(EMAIL_CONFIG['username'], EMAIL_CONFIG['password'])

status, folders = mail.list()
print("所有文件夹:")
for folder in folders:
    print(folder.decode('utf-8', errors='ignore'))

print("\n查找04开头的文件夹:")
for folder in folders:
    folder_str = folder.decode('utf-8', errors='ignore')
    if '04' in folder_str:
        print(f"找到: {folder_str}")

mail.logout()
