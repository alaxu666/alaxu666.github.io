#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置文件
"""

# 系统URL
LOGIN_URL = "https://plmcnprdawc.yanfeng.com:3000/#/showHome"

# 登录凭据（建议从环境变量或安全存储中读取）
USERNAME = "uxuxs004"
PASSWORD = "ABCabc%123"

# 文件路径配置
DOWNLOAD_DIR = r"C:\Users\uxuxs004\Downloads"
OUTPUT_DIR = r"C:\Users\uxuxs004\OneDrive - Yanfeng\work\产品安全\SSO\pythonProject"

# 搜索参数
BU_VALUE = "DFM"
CATEGORY_VALUE = "S"
PRODUCT_GROUP = "equipment"

# 等待时间配置
DEFAULT_WAIT_TIME = 10
DOWNLOAD_WAIT_TIME = 5
PAGE_LOAD_WAIT_TIME = 3

# 输出列配置
OUTPUT_COLUMNS = [
    'Project ID', 'Project Name', 'OEM Name', 'EBP Leader',
    'Product Engineer', 'SMTE', 'Project State',
    'Phase 1 Gate Exit-GO', 'Phase 2 Gate Exit-DVR',
    'Phase 3 Gate Exit-FPR', 'Phase 4 Gate Exit-CPA',
    'Phase 5 Gate Exit-PLR', '当前状态', 'PKR信息'
]

# 邮件配置
# 注意：由于Office 365安全策略，SMTP自动发送可能会失败
# 如果SMTP失败，系统会自动回退到打开邮件客户端手动发送
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SENDER_EMAIL = "uxuxs004@yanfeng.com"
SENDER_PASSWORD = "ABCabc%123"
RECIPIENT_EMAIL = "shirong.xu@yanfeng.com"
