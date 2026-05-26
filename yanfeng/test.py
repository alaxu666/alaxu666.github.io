#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
SMTP邮件发送测试脚本
功能：测试发件邮箱的SMTP配置是否正确，能否成功发送邮件
使用方法：修改下方的邮件配置，然后运行此脚本
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys

# ==================== 请修改以下配置 ====================
SMTP_SERVER = "smtp.office365.com"      # SMTP服务器地址
SMTP_PORT = 587                          # 端口（TLS通常为587）
SENDER_EMAIL = "shirong.xu@yanfeng.com"   # 发件人邮箱
SENDER_PASSWORD = "2045O2O5-7"    # 密码或应用专用密码
RECIPIENT_EMAIL = "shirong.xu@yanfeng.com"     # 测试收件人邮箱（可以填自己的邮箱）
# =======================================================

def send_test_email():
    """发送一封简单的测试邮件"""
    try:
        # 创建邮件内容
        subject = "SMTP测试邮件 - 来自Python脚本"
        body = """
        您好！

        这是一封通过 Python smtplib 发送的测试邮件。
        如果您收到了这封邮件，说明您的 SMTP 配置正确，可以用于自动发送邮件功能。

        发送时间: 测试运行时刻
        发件服务器: {}
        
        祝好！
        """.format(SMTP_SERVER)

        # 构造 MIME 邮件
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = subject
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECIPIENT_EMAIL

        # 连接服务器并发送
        print(f"正在连接 SMTP 服务器 {SMTP_SERVER}:{SMTP_PORT} ...")
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # 启用 TLS 加密
        print("已建立 TLS 连接")

        print("正在登录...")
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("登录成功")

        print("正在发送邮件...")
        server.send_message(msg)
        server.quit()

        print("=" * 50)
        print("✅ 邮件发送成功！")
        print(f"   发件人: {SENDER_EMAIL}")
        print(f"   收件人: {RECIPIENT_EMAIL}")
        print(f"   主题: {subject}")
        print("=" * 50)
        return True

    except smtplib.SMTPAuthenticationError:
        print("❌ 认证失败：用户名或密码错误")
        print("   提示：如果使用企业邮箱（如 Office 365），可能需要使用“应用专用密码”")
        return False
    except smtplib.SMTPConnectError:
        print("❌ 连接服务器失败，请检查 SMTP_SERVER 和 SMTP_PORT 是否正确")
        return False
    except smtplib.SMTPException as e:
        print(f"❌ SMTP 错误: {e}")
        return False
    except Exception as e:
        print(f"❌ 未知错误: {e}")
        return False

def main():
    print("=" * 50)
    print("SMTP 邮件发送测试工具")
    print("=" * 50)
    print("当前配置:")
    print(f"  SMTP_SERVER   = {SMTP_SERVER}")
    print(f"  SMTP_PORT     = {SMTP_PORT}")
    print(f"  SENDER_EMAIL  = {SENDER_EMAIL}")
    print(f"  RECIPIENT_EMAIL = {RECIPIENT_EMAIL}")
    print("-" * 50)

    # 简单检查配置是否被修改
    if SENDER_EMAIL == "your.name@yanfeng.com" or "your_" in SENDER_EMAIL:
        print("⚠️ 警告：您尚未修改发件人邮箱配置！")
        print("   请先编辑脚本开头的 SMTP_SERVER、SENDER_EMAIL、SENDER_PASSWORD 等变量")
        confirm = input("   是否仍要继续测试？(y/N): ")
        if confirm.lower() != 'y':
            print("测试已取消")
            sys.exit(0)

    if SENDER_PASSWORD == "your_app_password":
        print("⚠️ 警告：密码仍为默认值，请填写真实密码")
        confirm = input("   是否仍要继续测试？(y/N): ")
        if confirm.lower() != 'y':
            print("测试已取消")
            sys.exit(0)

    success = send_test_email()
    if not success:
        print("\n建议：")
        print("1. 确认密码是否正确（部分邮箱需使用“应用专用密码”）")
        print("2. 确认 SMTP 服务器地址和端口是否正确")
        print("3. 确认网络是否可以访问该 SMTP 服务器")
        print("4. 尝试使用 telnet 或 ping 测试连通性")
        sys.exit(1)
    else:
        print("\n测试通过！可以继续使用自动邮件发送功能。")
        sys.exit(0)

if __name__ == "__main__":
    main()