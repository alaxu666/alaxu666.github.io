#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PKR数据爬取脚本
功能：自动登录系统，爬取PKR相关数据并处理
"""

import os
import time
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import glob
import re
import urllib.parse
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
try:
    import pyperclip
    HAS_PYPERCLIP = True
    print("pyperclip已安装，启用剪贴板自动复制功能")
except ImportError:
    HAS_PYPERCLIP = False
    print("注意: pyperclip未安装，将跳过剪贴板复制功能")

# 导入配置和模块
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_loader import load_config_module
from YanfengAutoWork import setup_driver, PLMLogin, DL_Project_List, RD_Project_List, FindEBPLeader

config = load_config_module()
for name in dir(config):
    if name.isupper():
        globals()[name] = getattr(config, name)

class PKRDataCrawler:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.mail_driver = None
        self.download_dir = DOWNLOAD_DIR
        self.output_dir = OUTPUT_DIR
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.confirm_dir = os.path.join(self.script_dir, "PKR确认信息")
        os.makedirs(self.download_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.confirm_dir, exist_ok=True)

    def set_clipboard_html(self, html_content):
        """将HTML内容设置到Windows剪贴板，支持表格粘贴。"""
        try:
            import win32clipboard
            import win32con
        except ImportError:
            return False

        html_prefix = "<html><body><!--StartFragment-->"
        html_suffix = "<!--EndFragment--></body></html>"
        fragment = html_content
        html = html_prefix + fragment + html_suffix

        header = (
            "Version:0.9\r\n"
            "StartHTML:0000000000\r\n"
            "EndHTML:0000000000\r\n"
            "StartFragment:0000000000\r\n"
            "EndFragment:0000000000\r\n"
        )
        header_bytes = header.encode('utf-8')
        html_bytes = html.encode('utf-8')

        start_html = len(header_bytes)
        start_fragment = start_html + len(html_prefix.encode('utf-8'))
        end_fragment = start_fragment + len(fragment.encode('utf-8'))
        end_html = start_html + len(html_bytes)

        header = (
            "Version:0.9\r\n"
            f"StartHTML:{start_html:010d}\r\n"
            f"EndHTML:{end_html:010d}\r\n"
            f"StartFragment:{start_fragment:010d}\r\n"
            f"EndFragment:{end_fragment:010d}\r\n"
        )
        full_html = header.encode('utf-8') + html_bytes

        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
            win32clipboard.SetClipboardData(html_format, full_html)
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, fragment)
            return True
        finally:
            win32clipboard.CloseClipboard()

    def auto_paste_to_active_window(self):
        """向当前活动窗口发送 Ctrl+V 粘贴命令。"""
        try:
            import ctypes
            user32 = ctypes.windll.user32
            VK_CONTROL = 0x11
            VK_V = 0x56
            KEYEVENTF_KEYUP = 0x0002

            user32.keybd_event(VK_CONTROL, 0, 0, 0)
            user32.keybd_event(VK_V, 0, 0, 0)
            time.sleep(0.05)
            user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
            user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
            return True
        except Exception:
            return False

    def open_outlook_mail_window(self, recipient_emails, cc_emails, subject, html_content):
        """使用 Outlook COM 打开邮件窗口，按收件人列表逐个添加，避免 To/Cc 字符串被压成单人。"""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = subject

            recipients = [email.strip() for email in re.split(r'[;,]', recipient_emails) if email.strip()]
            cc_list = [email.strip() for email in re.split(r'[;,]', cc_emails) if email.strip()]

            for recipient in recipients:
                try:
                    rec = mail.Recipients.Add(recipient)
                    rec.Type = 1
                except Exception:
                    pass
            for cc in cc_list:
                try:
                    rec = mail.Recipients.Add(cc)
                    rec.Type = 2
                except Exception:
                    pass

            try:
                mail.Recipients.ResolveAll()
            except Exception:
                pass

            mail.HTMLBody = html_content
            mail.Display(False)
            print(f"Outlook 收件人列表最终递归添加: {recipients}")
            print(f"Outlook 抄送列表最终递归添加: {cc_list}")
            return True
        except Exception as e:
            print(f"Outlook COM 打开失败: {e}")
            return False

    # -------------------- 新增：网页版 Outlook 邮件发送 --------------------
    def send_web_mail(self, recipients_list, cc_list, subject, html_content):
        """
        使用 Edge 浏览器打开 Office 365 Outlook 网页版，自动填写邮件，保持浏览器打开。
        """
        print("=== 进入 send_web_mail ===")
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.edge.service import Service as EdgeService
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.keys import Keys
        from selenium.common.exceptions import TimeoutException, NoSuchElementException

        email = SENDER_EMAIL
        password = SENDER_PASSWORD
        print(f"发件人邮箱: {email}")

        driver_path = r"C:\XSR\githubPage\yanfeng\msedgedriver.exe"
        print(f"使用驱动: {driver_path}")

        options = EdgeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_experimental_option("detach", True)

        service = EdgeService(driver_path)
        self.mail_driver = webdriver.Edge(service=service, options=options)
        mail_wait = WebDriverWait(self.mail_driver, 60)  # 全局超时提高到60秒

        try:
            print("正在打开 Outlook 网页版...")
            self.mail_driver.get("https://outlook.office.com/mail/")
            time.sleep(5)

            # ---------- 登录 ----------
            email_input = mail_wait.until(EC.presence_of_element_located((By.NAME, "loginfmt")))
            email_input.clear()
            email_input.send_keys(email)
            next_btn = mail_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
            next_btn.click()
            print("已点击'下一步'")

            time.sleep(2)
            pwd_input = mail_wait.until(EC.presence_of_element_located((By.NAME, "passwd")))
            pwd_input.clear()
            pwd_input.send_keys(password)
            time.sleep(1)
            signin_btn = mail_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
            signin_btn.click()
            print("已点击'登录'")

            try:
                stay_btn = mail_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='是']")))
                stay_btn.click()
            except:
                pass

            # 等待邮箱主界面加载（“新邮件”出现即为加载完成）
            print("等待邮箱主界面加载...")
            mail_wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'新邮件')]")))
            print("登录成功，邮箱已加载")

            # ---------- 快速点击“新邮件” ----------
            print("快速点击'新邮件'按钮...")
            new_mail_clicked = False

            # 尝试3次点击
            for attempt in range(3):
                try:
                    selectors = [
                        "//span[contains(text(),'新邮件')]",
                        "//button[contains(text(),'新邮件')]",
                        "//div[@role='button' and contains(text(),'新邮件')]",
                        "//span[contains(text(),'New message')]"
                    ]
                    elem = None
                    for sel in selectors:
                        try:
                            elem = self.mail_driver.find_element(By.XPATH, sel)
                            if elem.is_enabled() and elem.is_displayed():
                                break
                        except:
                            continue

                    if elem and elem.is_enabled() and elem.is_displayed():
                        self.mail_driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(0.5)
                        try:
                            elem.click()
                        except:
                            try:
                                ActionChains(self.mail_driver).move_to_element(elem).click().perform()
                            except:
                                self.mail_driver.execute_script("arguments[0].click();", elem)
                        new_mail_clicked = True
                        print("  点击'新邮件'成功")
                        break
                except:
                    pass
                time.sleep(0.5)

            # 如果快速点击失败，使用快捷键 Ctrl+N
            if not new_mail_clicked:
                print("快速点击失败，立即使用快捷键 Ctrl+N...")
                try:
                    body = self.mail_driver.find_element(By.TAG_NAME, "body")
                    body.click()
                    time.sleep(0.5)
                except:
                    pass
                ActionChains(self.mail_driver).key_down(Keys.CONTROL).send_keys('n').key_up(Keys.CONTROL).perform()
                new_mail_clicked = True

            # ---------- 等待撰写面板加载（增强版） ----------
            if new_mail_clicked:
                print("等待新邮件撰写面板加载（最多60秒）...")
                
                # 先等待页面文档加载完成
                try:
                    WebDriverWait(self.mail_driver, 30).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                    print("  页面文档加载完成")
                except TimeoutException:
                    print("  ⚠️ 文档加载超时，但继续尝试")

                # 等待收件人输入框出现（超时60秒，轮询间隔1秒）
                try:
                    recipient_div = WebDriverWait(self.mail_driver, 60, poll_frequency=1).until(
                        EC.presence_of_element_located((By.XPATH, "//div[@aria-label='收件人' or @aria-label='To']"))
                    )
                    print("新邮件撰写面板已加载")
                except TimeoutException:
                    # 检查是否在新窗口
                    if len(self.mail_driver.window_handles) > 1:
                        self.mail_driver.switch_to.window(self.mail_driver.window_handles[-1])
                        print("切换到新窗口")
                        recipient_div = WebDriverWait(self.mail_driver, 60, poll_frequency=1).until(
                            EC.presence_of_element_located((By.XPATH, "//div[@aria-label='收件人' or @aria-label='To']"))
                        )
                        print("新窗口中找到收件人输入框")
                    else:
                        # 如果还没出现，尝试用其他元素检测（如主题输入框）
                        try:
                            subject_input = WebDriverWait(self.mail_driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='添加主题' or @placeholder='Add a subject']"))
                            )
                            print("检测到主题输入框，推测撰写面板已打开，但收件人未找到")
                            # 此时没有 recipient_div，跳过收件人填写，后面会尝试直接找收件人
                            recipient_div = None
                        except:
                            raise TimeoutError("未能检测到撰写面板，请手动检查")

                # 填写收件人（如果找到了）
                if recipient_div:
                    print("填写收件人...")
                    recipient_div.click()
                    recipients_str = '; '.join(recipients_list) if isinstance(recipients_list, list) else str(recipients_list)
                    try:
                        recipient_input = recipient_div.find_element(By.TAG_NAME, "input")
                        recipient_input.send_keys(recipients_str)
                        recipient_input.send_keys("\n")
                    except:
                        recipient_div.send_keys(recipients_str)
                    time.sleep(1)
                else:
                    # 如果没找到收件人，尝试直接找输入框（某些界面下可能不同）
                    try:
                        recipient_input = self.mail_driver.find_element(By.XPATH, "//div[@aria-label='收件人' or @aria-label='To']//input")
                        recipient_input.send_keys(recipients_str)
                        recipient_input.send_keys("\n")
                        print("通过备用方式填写收件人")
                    except:
                        print("⚠️ 未找到收件人输入框，请手动填写")

                # 抄送（如果提供）
                if cc_list:
                    print("尝试填写抄送...")
                    try:
                        # 1. 先尝试点击“抄送”链接（可能不存在，忽略）
                        try:
                            cc_link = self.mail_driver.find_element(By.XPATH, "//span[contains(text(),'抄送') or contains(text(),'Cc')]")
                            self.mail_driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cc_link)
                            time.sleep(0.3)
                            self.mail_driver.execute_script("arguments[0].click();", cc_link)
                            time.sleep(0.5)
                        except:
                            pass

                        # 2. 定位抄送输入区域
                        cc_div = WebDriverWait(self.mail_driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//div[@aria-label='抄送' or @aria-label='Cc']"))
                        )
                        # 滚动到可见并强制点击
                        self.mail_driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cc_div)
                        time.sleep(0.3)
                        self.mail_driver.execute_script("arguments[0].click();", cc_div)

                        # 输入抄送地址
                        cc_str = '; '.join(cc_list) if isinstance(cc_list, list) else str(cc_list)
                        try:
                            cc_input = cc_div.find_element(By.TAG_NAME, "input")
                            cc_input.send_keys(cc_str)
                            cc_input.send_keys("\n")
                        except:
                            cc_div.send_keys(cc_str)
                        print("  抄送已填写")
                    except Exception as e:
                        print(f"  抄送填写跳过: {e}")

                # 主题
                print("填写主题...")
                subject_input = WebDriverWait(self.mail_driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@placeholder='添加主题' or @placeholder='Add a subject']"))
                )
                subject_input.clear()
                subject_input.send_keys(subject)

                # 正文
                print("填写邮件正文...")
                body_div = WebDriverWait(self.mail_driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@role='textbox' and (@aria-label='邮件正文' or @aria-label='Message body')]"))
                )
                body_div.clear()
                self.mail_driver.execute_script("arguments[0].innerHTML = arguments[1];", body_div, html_content)
                print("邮件正文已填入")

                print("\n✅ 邮件已填写完成，请检查后手动点击发送。")
                print("浏览器将保持打开，您可以安全地关闭此终端窗口。")

            else:
                print("❌ 未能点击'新邮件'按钮，请手动操作。")

        except Exception as e:
            print(f"❌ 网页版邮件填写失败: {e}")
            import traceback
            traceback.print_exc()
            print("浏览器保持打开以便调试，请手动检查。")

    # -------------------- 邮件 HTML 内容构建 --------------------
    def _build_mail_html_content(self, incomplete_df):
        """根据 DataFrame 生成邮件 HTML 正文，使用内联样式确保 Outlook 支持"""
        html_content = """
        <html>
        <head>
            <meta charset="utf-8">
        </head>
        <body style="font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; color: #333;">
        <h3 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 10px;">PKR确认状态未完成提醒</h3>
        <table style="border-collapse: collapse; width: 100%; margin: 20px 0; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border: 1px solid #ddd;">
            <thead>
                <tr>
                    <th style="border: 1px solid #ddd; padding: 12px; text-align: left; background-color: #4CAF50; color: white; font-weight: bold;">Project Name</th>
                    <th style="border: 1px solid #ddd; padding: 12px; text-align: left; background-color: #4CAF50; color: white; font-weight: bold;">当前状态</th>
                    <th style="border: 1px solid #ddd; padding: 12px; text-align: left; background-color: #4CAF50; color: white; font-weight: bold;">未完成PKR信息</th>
                    <th style="border: 1px solid #ddd; padding: 12px; text-align: left; background-color: #4CAF50; color: white; font-weight: bold;">是否确认</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in incomplete_df.iterrows():
            project_name = str(row.get('Project Name', '')).strip()
            current_status = str(row.get('当前状态', '')).strip()
            incomplete_pkr = str(row.get('未完成PKR信息', '')).strip()
            confirm_status = str(row.get('是否确认', '')).strip()
            if not incomplete_pkr or incomplete_pkr == "nan":
                incomplete_pkr = "无"

            # 根据状态设置行背景色
            if confirm_status == "未递交(工程师)":
                row_style = "background-color: #ffebee;"
            elif confirm_status == "未确认(cao liang)":
                row_style = "background-color: #fff3e0;"
            else:
                row_style = ""

            # 状态文字颜色
            if confirm_status == "未递交(工程师)":
                status_style = "color: #d32f2f; font-weight: bold;"
            elif confirm_status == "未确认(cao liang)":
                status_style = "color: #f57c00; font-weight: bold;"
            else:
                status_style = ""

            html_content += f"""
            <tr style="{row_style}">
                <td style="border: 1px solid #ddd; padding: 12px; text-align: left;">{project_name}</td>
                <td style="border: 1px solid #ddd; padding: 12px; text-align: left;">{current_status}</td>
                <td style="border: 1px solid #ddd; padding: 12px; text-align: left;">{incomplete_pkr}</td>
                <td style="border: 1px solid #ddd; padding: 12px; text-align: left; {status_style}">{confirm_status}</td>
            </tr>
            """

        html_content += f"""
            </tbody>
        </table>
        <div style="margin-top: 20px; font-size: 12px; color: #666;">
            <p>总计 {len(incomplete_df)} 个项目需要关注</p>
            <p><span style="color: #d32f2f; font-weight: bold;">未递交(工程师)</span> - 需要工程师提交PKR</p>
            <p><span style="color: #f57c00; font-weight: bold;">未确认(cao liang)</span> - 需要曹亮确认</p>
        </div>
        </body>
        </html>
        """
        return html_content

    # -------------------- 重写 send_teams_message --------------------
    def send_teams_message(self, df1):
        print("=== 进入 send_teams_message ===")
        try:
            # 更新"是否确认"列（逻辑不变）
            for idx, row in df1.iterrows():
                incomplete_pkr = str(row.get('未完成PKR信息', '')).strip()
                confirmed_info = str(row.get('Confirmed', '')).strip()
                if incomplete_pkr == "完成":
                    if confirmed_info and "Confirmed" in confirmed_info:
                        confirmed_lines = confirmed_info.split('\n')
                        all_confirmed = True
                        for line in confirmed_lines:
                            if line.strip() and not line.endswith(',Confirmed'):
                                all_confirmed = False
                                break
                        if all_confirmed:
                            df1.at[idx, '是否确认'] = '完成确认'
                        else:
                            df1.at[idx, '是否确认'] = '未确认(cao liang)'
                    else:
                        df1.at[idx, '是否确认'] = '未确认(cao liang)'
                else:
                    df1.at[idx, '是否确认'] = '未递交(工程师)'

            incomplete_df = df1[df1['是否确认'] != '完成确认'].copy()
            incomplete_df = incomplete_df[incomplete_df['当前状态'] != 'Phase 1 Gate Exit-GO']

            print(f"未完成数据行数: {len(incomplete_df)}")

            if len(incomplete_df) == 0:
                print("没有需要关注的数据，不发送邮件")
                return

            print(f"找到 {len(incomplete_df)} 条需要关注的数据")

            recipient_emails = self.extract_unique_pkr_names(incomplete_df)
            if not recipient_emails:
                recipient_emails = self.collect_recipient_emails(df1)
            if not recipient_emails:
                recipient_emails = RECIPIENT_EMAIL

            recipient_list = list(dict.fromkeys(
                [email.strip() for email in re.split(r'[;,]', recipient_emails) if email.strip()]
            ))
            print(f"收件人列表: {recipient_list}")

            cc_list = ["liang.cao@yanfeng.com"]
            subject = "PKR确认状态未完成提醒"
            html_content = self._build_mail_html_content(incomplete_df)

            print("即将调用 send_web_mail...")
            self.send_web_mail(recipient_list, cc_list, subject, html_content)
            print("send_web_mail 调用完成（若无异常，浏览器应已打开）")

        except Exception as e:
            print(f"邮件发送过程出错: {e}")
            import traceback
            traceback.print_exc()

    # -------------------- 原有方法（全部保留） --------------------
    def navigate_to_xso_management(self):
        """导航到XSO Management页面"""
        try:
            xso_pkr_div = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='XSO & PKR']"))
            )
            xso_pkr_div.click()
            print("已点击XSO & PKR")
            time.sleep(PAGE_LOAD_WAIT_TIME * 2)
            iframe_content = self.wait.until(
                EC.presence_of_element_located((By.ID, "iframeContent"))
            )
            time.sleep(PAGE_LOAD_WAIT_TIME)
            iframe_content = self.driver.find_element(By.ID, "iframeContent")
            self.driver.switch_to.frame(iframe_content)
            print("已切换到iframeContent")
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    try:
                        xso_management = self.wait.until(
                            EC.element_to_be_clickable((By.LINK_TEXT, "XSO Management"))
                        )
                    except:
                        xso_management = self.wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//span[text()='XSO Management']/parent::a"))
                        )
                    xso_management.click()
                    print("已展开XSO Management")
                    break
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise e
                    print(f"第{attempt + 1}次尝试展开XSO Management失败，重试中...")
                    time.sleep(PAGE_LOAD_WAIT_TIME)
            program_dashboard = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Program Dashboard"))
            )
            program_dashboard.click()
            print("已点击Program Dashboard")
            time.sleep(PAGE_LOAD_WAIT_TIME * 2)
        except Exception as e:
            print(f"导航到XSO Management过程中出现错误: {e}")
            try:
                print("当前页面源码片段:")
                print(self.driver.page_source[:1000])
            except:
                pass
            raise

    def wait_for_project_list_download(self, existing_files=None, timeout=120, poll_interval=2):
        if existing_files is None:
            existing_files = set()
        pattern = os.path.join(self.download_dir, "Project_List*.xls*")
        end_time = time.time() + timeout
        last_size = {}
        while time.time() < end_time:
            all_files = [f for f in glob.glob(pattern) if not f.lower().endswith('.crdownload') and not f.lower().endswith('.part')]
            new_files = [f for f in all_files if f not in existing_files]
            if new_files:
                latest_file = max(new_files, key=os.path.getmtime)
                try:
                    size = os.path.getsize(latest_file)
                except OSError:
                    size = -1
                previous = last_size.get(latest_file)
                if previous == size and size > 0:
                    return latest_file
                last_size[latest_file] = size
            time.sleep(poll_interval)
        raise TimeoutError(f"等待Project_List下载完成超时: {pattern}")

    def download_xso_data(self):
        try:
            iframe = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-1-1"))
            )
            self.driver.switch_to.frame(iframe)
            print("已切换到XSO iframe")
            bu_select = Select(self.driver.find_element(By.NAME, "select_bu"))
            bu_select.select_by_value(BU_VALUE)
            print(f"已选择BU: {BU_VALUE}")
            category_select = Select(self.driver.find_element(By.NAME, "select_Category"))
            category_select.select_by_value(CATEGORY_VALUE)
            print(f"已选择Category: {CATEGORY_VALUE}")
            product_group_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='Product Group']")
            product_group_input.clear()
            product_group_input.send_keys(PRODUCT_GROUP)
            print(f"已输入Product Group: {PRODUCT_GROUP}")
            search_button = self.driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']")
            search_button.click()
            print("已点击搜索按钮")
            time.sleep(PAGE_LOAD_WAIT_TIME)
            existing_files = set(glob.glob(os.path.join(self.download_dir, "Project_List*.xls*")))
            export_button = self.driver.find_element(By.CSS_SELECTOR, "span[onclick='exportData()']")
            export_button.click()
            print("已点击导出按钮")
            download_file = self.wait_for_project_list_download(existing_files=existing_files, timeout=180)
            print(f"下载完成: {download_file}")
            self.driver.switch_to.default_content()
        except Exception as e:
            print(f"下载XSO数据过程中出现错误: {e}")
            raise

    def get_latest_downloaded_file(self):
        try:
            pattern = os.path.join(self.download_dir, "Project_List*.xls*")
            files = glob.glob(pattern)
            if not files:
                raise FileNotFoundError("未找到Project_List文件")
            latest_file = max(files, key=os.path.getmtime)
            print(f"找到最新文件: {latest_file}")
            return latest_file
        except Exception as e:
            print(f"获取最新下载文件过程中出现错误: {e}")
            raise

    def process_excel_data(self, file_path):
        try:
            df = pd.read_excel(file_path)
            print(f"已读取Excel文件，共{len(df)}行数据")
            today = datetime.now()
            current_week_start = today - timedelta(days=today.weekday())
            current_week_end = current_week_start + timedelta(days=6)
            next_week_start = current_week_start + timedelta(days=7)
            next_week_end = next_week_start + timedelta(days=6)
            print(f"当前周: {current_week_start.strftime('%Y-%m-%d')} 到 {current_week_end.strftime('%Y-%m-%d')}")
            print(f"下周: {next_week_start.strftime('%Y-%m-%d')} 到 {next_week_end.strftime('%Y-%m-%d')}")
            df = df[df['Category'] == 'S'].copy()
            print(f"筛选Category=S后，剩余{len(df)}行数据")
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA'
            ]
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            mask = False
            for col in date_columns:
                if col in df.columns:
                    col_mask = (
                        (df[col] >= current_week_start) & (df[col] <= current_week_end) |
                        (df[col] >= next_week_start) & (df[col] <= next_week_end)
                    )
                    mask = mask | col_mask
            df_filtered = df[mask].copy()
            print(f"筛选日期范围后，剩余{len(df_filtered)}行数据")
            expanded_rows = []
            for idx, row in df_filtered.iterrows():
                project_id = row['Project ID']
                for col in date_columns:
                    if col in df_filtered.columns and pd.notna(row[col]):
                        if (current_week_start <= row[col] <= current_week_end) or \
                           (next_week_start <= row[col] <= next_week_end):
                            new_row = row.copy()
                            new_row['当前状态'] = col
                            expanded_rows.append(new_row)
            if expanded_rows:
                df1 = pd.DataFrame(expanded_rows)
                print(f"展开多phase数据后，共{len(df1)}行数据")
            else:
                df1 = df_filtered.copy()
                df1['当前状态'] = ''
                print("未找到符合条件的数据")
            df1['PKR信息'] = ''
            df1['Confirmed'] = ''
            df1['是否确认'] = ''
            df1['未完成人名'] = ''
            df1.insert(0, 'id', range(1, len(df1) + 1))
            return df1
        except Exception as e:
            print(f"处理Excel数据过程中出现错误: {e}")
            raise

    def navigate_to_pkr_management(self):
        try:
            self.driver.switch_to.default_content()
            print("已切换回主文档")
            time.sleep(PAGE_LOAD_WAIT_TIME)
            iframe_content = self.wait.until(
                EC.presence_of_element_located((By.ID, "iframeContent"))
            )
            self.driver.switch_to.frame(iframe_content)
            print("已切换到iframeContent")
            pkr_management = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "PKR Management"))
            )
            pkr_management.click()
            print("已展开PKR Management")
            pkr_summary = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Program PKR Summary"))
            )
            pkr_summary.click()
            print("已点击Program PKR Summary")
            time.sleep(PAGE_LOAD_WAIT_TIME)
        except Exception as e:
            print(f"导航到PKR Management过程中出现错误: {e}")
            raise

    def extract_pkr_info(self, df1):
        try:
            df1 = df1.reset_index(drop=True)
            iframe = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-2-2"))
            )
            self.driver.switch_to.frame(iframe)
            print("已切换到PKR iframe")
            for idx, row in df1.iterrows():
                project_id = row['Project ID']
                current_status = row['当前状态']
                row_id = row.get('id', idx+1)
                print(f"处理行索引={idx}, id={row_id}, Project ID: {project_id}, 当前状态: {current_status}")
                if current_status == 'Phase 1 Gate Exit-GO':
                    print(f"Project ID {project_id} 当前状态为Phase 1，跳过PKR提取")
                    df1.at[idx, 'PKR信息'] = ''
                    df1.at[idx, 'Confirmed'] = ''
                    df1.at[idx, '是否确认'] = ''
                    continue
                aria_attr_map = {
                    'Phase 2 Gate Exit-DVR': 'result-table_dvr_html',
                    'Phase 3 Gate Exit-FPR': 'result-table_fpr_html',
                    'Phase 4 Gate Exit-CPA': 'result-table_cpa_html'
                }
                target_aria = aria_attr_map.get(current_status)
                if not target_aria:
                    print(f"未知的当前状态: {current_status}，跳过")
                    df1.at[idx, 'PKR信息'] = ''
                    df1.at[idx, 'Confirmed'] = ''
                    df1.at[idx, '是否确认'] = ''
                    continue
                try:
                    search_input = self.wait.until(
                        EC.presence_of_element_located((By.ID, "searchProject"))
                    )
                    search_input.clear()
                    search_input.send_keys(str(project_id))
                    search_button = self.driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']")
                    search_button.click()
                    time.sleep(PAGE_LOAD_WAIT_TIME)
                    lbl_span = self.wait.until(
                        EC.presence_of_element_located((By.CLASS_NAME, "lbl"))
                    )
                    self.driver.execute_script("arguments[0].click();", lbl_span)
                    time.sleep(2)
                    td_elements = self.driver.find_elements(
                        By.CSS_SELECTOR, f"td[aria-describedby='{target_aria}']"
                    )
                    pkr_info = ""
                    confirmed_info = ""
                    for td in td_elements:
                        div_scores = td.find_elements(By.CLASS_NAME, "div_score")
                        for div in div_scores:
                            onclick_attr = div.get_attribute("onclick")
                            if onclick_attr:
                                onclick_parts = onclick_attr.split(',')
                                if len(onclick_parts) > 1:
                                    last_part = onclick_parts[-1].strip()
                                    match = re.search(r"'([^_']+)_", last_part)
                                    if match:
                                        pkr_name = match.group(1).replace("%20", " ")
                                    else:
                                        pkr_name = "Total"
                                else:
                                    pkr_name = "Total"
                            else:
                                pkr_name = "Total"
                            pkr_score = div.text.strip()
                            line = f"{pkr_name},{pkr_score}"
                            if pkr_info:
                                pkr_info += "\n"
                            pkr_info += line
                            try:
                                i_elements = div.find_elements(By.XPATH, "preceding-sibling::i[1]")
                                if not i_elements:
                                    i_elements = td.find_elements(By.TAG_NAME, "i")
                                for i_tag in i_elements:
                                    confirmed_title = i_tag.get_attribute("title") or i_tag.text.strip()
                                    if confirmed_title:
                                        if confirmed_info:
                                            confirmed_info += "\n"
                                        confirmed_info += f"{pkr_name},{confirmed_title}"
                                    break
                            except Exception as e:
                                print(f"提取Confirmed信息时出错: {e}")
                                continue
                    formatted_pkr_info = ""
                    if pkr_info and pkr_info.strip():
                        lines = pkr_info.split('\n')
                        formatted_lines = []
                        for line in lines:
                            if line.strip():
                                parts = line.split(',', 1)
                                if len(parts) == 2:
                                    pkr_name, pkr_score = parts
                                    if pkr_name == "Total":
                                        formatted_line = f"最低打分：{pkr_score}；\n"
                                    else:
                                        formatted_line = f"{pkr_name}的PKR打分：{pkr_score}；\n"
                                    formatted_lines.append(formatted_line)
                        formatted_pkr_info = ''.join(formatted_lines)
                    df1.at[idx, 'PKR信息'] = formatted_pkr_info
                    df1.at[idx, 'Confirmed'] = confirmed_info
                    if confirmed_info and confirmed_info.strip():
                        confirmed_lines = [line.strip() for line in confirmed_info.split('\n') if line.strip()]
                        if not confirmed_lines:
                            df1.at[idx, '是否确认'] = ''
                        else:
                            all_confirmed = True
                            all_submitted = True
                            has_other = False
                            for line in confirmed_lines:
                                if ',' in line:
                                    pkr_name, status = line.rsplit(',', 1)
                                    status = status.strip()
                                    if status != 'Confirmed':
                                        all_confirmed = False
                                    if status != 'Submitted':
                                        all_submitted = False
                                    if status not in ['Confirmed', 'Submitted']:
                                        has_other = True
                            if all_confirmed and not has_other:
                                df1.at[idx, '是否确认'] = '完成确认'
                            elif all_submitted and not has_other:
                                df1.at[idx, '是否确认'] = '未确认(Cao liang)'
                            elif has_other:
                                df1.at[idx, '是否确认'] = '未评分(工程师)'
                            else:
                                df1.at[idx, '是否确认'] = '未确认(Cao liang)'
                    else:
                        if pkr_info and pkr_info.strip():
                            df1.at[idx, '是否确认'] = '未递交(工程师)'
                        else:
                            df1.at[idx, '是否确认'] = ''
                    print(f"Project ID {project_id} PKR信息: {pkr_info}")
                    print(f"Project ID {project_id} Confirmed: {confirmed_info}")
                    print(f"Project ID {project_id} 确认状态: {df1.at[idx, '是否确认']}")
                except Exception as e:
                    print(f"处理行索引={idx} (Project ID {project_id}) 时出错: {e}")
                    df1.at[idx, 'PKR信息'] = "处理失败"
                    df1.at[idx, 'Confirmed'] = ""
                    df1.at[idx, '是否确认'] = "错误"
                    continue
            self.driver.switch_to.default_content()
            return df1
        except Exception as e:
            print(f"提取PKR信息过程中出现错误: {e}")
            raise

    def analyze_incomplete_pkr(self, df1):
        try:
            print("正在分析未完成PKR信息...")
            def extract_incomplete_pkr_by_score(row):
                current_status = row.get('当前状态', '')
                if current_status == 'Phase 1 Gate Exit-GO':
                    return "phase1不适用"
                pkr_info = row.get('PKR信息', '')
                if pd.isna(pkr_info) or not str(pkr_info).strip():
                    return "完成"
                incomplete_pkr = []
                pkr_text = str(pkr_info).strip()
                pkr_lines = pkr_text.split('\n')
                pkr_items = []
                for line in pkr_lines:
                    line = line.strip()
                    if line:
                        pkr_items.append(line)
                for item in pkr_items:
                    item = item.strip()
                    if not item:
                        continue
                    if "最低打分：" in item:
                        continue
                    if "的PKR打分：" in item:
                        try:
                            name_part, score_part = item.split("的PKR打分：", 1)
                            pkr_name = name_part.strip()
                            pkr_score_str = score_part.strip()
                            pkr_score_str_clean = pkr_score_str.split('；')[0]
                            pkr_score = int(pkr_score_str_clean)
                            if pkr_score != 100:
                                incomplete_pkr.append(f"{pkr_name}的PKR打分：{pkr_score}；\n")
                        except (ValueError, IndexError):
                            continue
                if incomplete_pkr:
                    return ''.join(incomplete_pkr).strip()
                else:
                    return "完成"
            df1['未完成PKR信息'] = df1.apply(extract_incomplete_pkr_by_score, axis=1)
            print(f"已完成{len(df1)}行数据的未完成PKR信息分析")
            incomplete_count = len(df1[df1['未完成PKR信息'] != "完成"])
            complete_count = len(df1[df1['未完成PKR信息'] == "完成"])
            print(f"包含未完成PKR的项目数: {incomplete_count}")
            print(f"PKR全部完成的项目数: {complete_count}")
            return df1
        except Exception as e:
            print(f"分析未完成PKR信息过程中出现错误: {e}")
            df1['未完成PKR信息'] = "分析失败"
            return df1

    def sort_data_by_date(self, df1):
        try:
            print("正在按日期排序数据...")
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA',
                'Phase 5 Gate Exit-PLR'
            ]
            for col in date_columns:
                if col in df1.columns:
                    df1[col] = pd.to_datetime(df1[col], errors='coerce')
            df1['最早日期'] = pd.NaT
            for idx, row in df1.iterrows():
                earliest_date = None
                for col in date_columns:
                    if col in df1.columns and pd.notna(row[col]):
                        if earliest_date is None or row[col] < earliest_date:
                            earliest_date = row[col]
                df1.at[idx, '最早日期'] = earliest_date
            df1_sorted = df1.sort_values(by='最早日期', na_position='last')
            if '最早日期' in df1_sorted.columns:
                df1_sorted = df1_sorted.drop('最早日期', axis=1)
            print(f"数据排序完成，共{len(df1_sorted)}行数据")
            return df1_sorted
        except Exception as e:
            print(f"数据排序过程中出现错误: {e}")
            return df1

    def remove_phases_beyond_next_sunday(self, df1):
        today = datetime.now()
        current_sunday = today + timedelta(days=(6 - today.weekday()))
        next_sunday = current_sunday + timedelta(days=7)
        next_sunday = next_sunday.replace(hour=0, minute=0, second=0, microsecond=0)
        print(f"删除条件：Gate Exit 日期 > {next_sunday.strftime('%Y-%m-%d')} 的行将被移除")
        phase_column_map = {
            'Phase 1 Gate Exit-GO': 'Phase 1 Gate Exit-GO',
            'Phase 2 Gate Exit-DVR': 'Phase 2 Gate Exit-DVR',
            'Phase 3 Gate Exit-FPR': 'Phase 3 Gate Exit-FPR',
            'Phase 4 Gate Exit-CPA': 'Phase 4 Gate Exit-CPA',
        }
        original_len = len(df1)
        keep_mask = pd.Series([True] * original_len)
        for idx, row in df1.iterrows():
            current_status = row.get('当前状态', '')
            if current_status in phase_column_map:
                date_col = phase_column_map[current_status]
                if date_col in df1.columns and pd.notna(row[date_col]):
                    gate_date = row[date_col]
                    if gate_date > next_sunday:
                        keep_mask.iloc[idx] = False
        df1_filtered = df1[keep_mask].copy()
        print(f"删除了 {original_len - len(df1_filtered)} 行超过下周日的 Phase 数据，剩余 {len(df1_filtered)} 行")
        return df1_filtered

    def save_historical_incomplete_records(self, df1):
        historical_file = os.path.join(self.confirm_dir, "PKR历史未完成记录.xlsx")
        mask = (
            (df1['是否确认'] != "完成确认") &
            (df1['是否确认'].notna()) &
            (df1['是否确认'].astype(str).str.strip() != "") &
            (df1['当前状态'] != 'Phase 1 Gate Exit-GO')
        )
        incomplete_df = df1[mask].copy()
        if len(incomplete_df) > 0:
            self.save_excel_with_retry(incomplete_df, historical_file, index=False)
            print(f"已更新历史未完成记录，共 {len(incomplete_df)} 条，保存至 {historical_file}")
        else:
            if os.path.exists(historical_file):
                os.remove(historical_file)
                print("所有项目已完成确认，已清空历史未完成记录文件")

    def add_missing_historical_projects(self, df1, original_project_list_file):
        historical_file = os.path.join(self.confirm_dir, "PKR历史未完成记录.xlsx")
        if not os.path.exists(historical_file):
            return df1
        try:
            hist_df = pd.read_excel(historical_file)
            hist_project_ids = set(hist_df['Project ID'].astype(str).unique())
            current_project_ids = set(df1['Project ID'].astype(str).unique())
            missing_ids = hist_project_ids - current_project_ids
            if not missing_ids:
                return df1
            print(f"历史记录中发现 {len(missing_ids)} 个项目不在当前数据中，尝试重新加入")
            original_df = pd.read_excel(original_project_list_file)
            missing_rows = original_df[original_df['Project ID'].astype(str).isin(missing_ids)]
            if len(missing_rows) == 0:
                print("未能在原始文件找到缺失的项目，跳过")
                return df1
            new_rows = self.filter_data_by_date_range(missing_rows)
            df1 = pd.concat([df1, new_rows], ignore_index=True)
            print(f"已添加 {len(new_rows)} 行历史项目数据")
            return df1
        except Exception as e:
            print(f"处理历史未完成记录文件时出错: {e}")
            return df1

    def add_gate_week_info(self, df1):
        try:
            print("正在添加过门周信息...")
            today = datetime.now()
            current_week_start = today - timedelta(days=today.weekday())
            current_week_end = current_week_start + timedelta(days=6)
            next_week_start = current_week_start + timedelta(days=7)
            next_week_end = next_week_start + timedelta(days=6)
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA',
                'Phase 5 Gate Exit-PLR'
            ]
            df1['过门周'] = ''
            for idx, row in df1.iterrows():
                gate_week = ''
                for col in date_columns:
                    if col in df1.columns and pd.notna(row[col]):
                        if current_week_start <= row[col] <= current_week_end:
                            gate_week = '本周'
                            break
                        elif next_week_start <= row[col] <= next_week_end:
                            gate_week = '下周'
                            break
                df1.at[idx, '过门周'] = gate_week
            print("过门周信息添加完成")
            return df1
        except Exception as e:
            print(f"添加过门周信息过程中出现错误: {e}")
            return df1

    def save_excel_with_retry(self, df, file_path, **kwargs):
        while True:
            try:
                df.to_excel(file_path, **kwargs)
                return
            except PermissionError:
                print("文件已被锁定，尝试关闭文件后，按回车键再试")
                input()
            except Exception as e:
                message = str(e).lower()
                if "being used by another process" in message or "permission denied" in message or "locked" in message:
                    print("文件已被锁定，尝试关闭文件后，按回车键再试")
                    input()
                    continue
                raise

    def extract_unique_pkr_names(self, df1):
        try:
            print("正在提取未完成人名...")
            incomplete_pkr_names = set()
            for idx, row in df1.iterrows():
                confirm_status = row.get('是否确认', '')
                if confirm_status != "完成确认":
                    incomplete_pkr_info = row.get('未完成PKR信息', '')
                    if pd.notna(incomplete_pkr_info) and str(incomplete_pkr_info).strip() \
                            and incomplete_pkr_info != "完成" and incomplete_pkr_info != "phase1不适用":
                        pkr_text = str(incomplete_pkr_info).strip()
                        pkr_lines = pkr_text.split('\n')
                        for line in pkr_lines:
                            line = line.strip()
                            if line and "的PKR打分：" in line:
                                pkr_name = line.split("的PKR打分：")[0].strip()
                                if pkr_name and pkr_name != 'Total':
                                    if pkr_name == "Wang Hao23":
                                        pkr_name = "Wang23 Hao"
                                    incomplete_pkr_names.add(pkr_name)
            pkr_names_list = sorted(list(incomplete_pkr_names))
            if not pkr_names_list:
                print("未找到未完成PKR信息中的PKR名称")
                return ""
            print(f"从未完成PKR信息中找到{len(pkr_names_list)}个PKR名称: {', '.join(pkr_names_list)}")
            email_addresses = []
            for pkr_name in pkr_names_list:
                name_parts = pkr_name.split(' ')
                if len(name_parts) >= 2:
                    last_name = name_parts[0]
                    first_name = name_parts[1]
                    email = f"{first_name}.{last_name}@yanfeng.com"
                    email_addresses.append(email)
                else:
                    email = f"{pkr_name.replace(' ', '.')}@yanfeng.com"
                    email_addresses.append(email)
            email_addresses.sort()
            email_format = '; '.join(email_addresses)
            print(f"生成的邮箱格式: {email_format}")
            print("未完成人名提取和组合完成")
            return email_format
        except Exception as e:
            print(f"提取未完成人名过程中出现错误: {e}")
            import traceback
            traceback.print_exc()
            return ""

    def collect_recipient_emails(self, df1):
        try:
            if '未完成人名' not in df1.columns:
                return ""
            unique_emails = []
            seen = set()
            import re
            for _, row in df1.iterrows():
                raw_value = str(row.get('未完成人名', '')).strip()
                if not raw_value:
                    continue
                email_parts = re.split(r'[;,]', raw_value)
                for email in email_parts:
                    email = email.strip()
                    if not email or '@' not in email:
                        continue
                    normalized = email.lower()
                    if normalized not in seen:
                        seen.add(normalized)
                        unique_emails.append(email)
            return '; '.join(unique_emails)
        except Exception as e:
            print(f"收集收件人邮箱时出错: {e}")
            return ""

    def export_final_data(self, df1):
        try:
            df1 = self.analyze_incomplete_pkr(df1)
            df1 = self.add_gate_week_info(df1)
            df1 = self.sort_data_by_date(df1)
            email_string = self.extract_unique_pkr_names(df1)
            df1['未完成人名'] = ''
            if email_string and len(df1) > 0:
                df1.iloc[0, df1.columns.get_loc('未完成人名')] = email_string
            columns_to_keep = OUTPUT_COLUMNS + ['未完成PKR信息', '过门周', 'Confirmed', '是否确认', '未完成人名']
            existing_columns = [col for col in columns_to_keep if col in df1.columns]
            complete_df = df1[existing_columns].copy()
            complete_file = os.path.join(self.confirm_dir, "PKR完成情况（近两周）.xlsx")
            self.save_excel_with_retry(complete_df, complete_file, index=False)
            print(f"完整数据已导出到: {complete_file}")
            print(f"共导出{len(complete_df)}行完整数据")
            incomplete_mask = (
                (df1['是否确认'] != "完成确认") &
                (df1['当前状态'] != 'Phase 1 Gate Exit-GO')
            )
            incomplete_df = df1[incomplete_mask].copy()
            incomplete_final_df = incomplete_df[existing_columns].copy()
            incomplete_file = os.path.join(self.confirm_dir, "PKR未完成情况（近两周）.xlsx")
            self.save_excel_with_retry(incomplete_final_df, incomplete_file, index=False)
            print(f"未完成数据已导出到: {incomplete_file}")
            print(f"共导出{len(incomplete_final_df)}行未完成数据")
            print(f"筛选条件: 是否确认≠'完成确认' 且 当前状态≠'Phase 1 Gate Exit-GO'")
            output_file = incomplete_file
            self.save_historical_incomplete_records(df1)
            return output_file
        except Exception as e:
            print(f"导出数据过程中出现错误: {e}")
            raise

    def update_ebp_leader_info(self, df1):
        try:
            print("正在提取EBP Leader信息...")
            empty_ebp_projects = df1[
                df1['EBP Leader'].isna() |
                (df1['EBP Leader'].astype(str).str.strip() == '')
            ]
            print(f"找到{len(empty_ebp_projects)}个EBP Leader为空的项目")
            if len(empty_ebp_projects) == 0:
                print("没有需要补充EBP Leader信息的项目")
                return df1
            project_id_list = empty_ebp_projects['Project ID'].tolist()
            ebp_leader_df = FindEBPLeader(self.driver, self.wait, project_id_list)
            for _, row in ebp_leader_df.iterrows():
                project_id = row['Project ID']
                ebp_leader = row['EBP Leader']
                mask = df1['Project ID'] == project_id
                if mask.any():
                    df1.loc[mask, 'EBP Leader'] = ebp_leader
                    if ebp_leader:
                        print(f"已更新Project ID {project_id}的EBP Leader: {ebp_leader}")
            print("EBP Leader信息更新完成")
            return df1
        except Exception as e:
            print(f"更新EBP Leader信息过程中出现错误: {e}")
            raise

    def filter_data_by_date_range(self, df1):
        """筛选Phase 1-4 Gate Exit日期在本周和下周末内的数据"""
        try:
            today = datetime.now()
            current_week_start = today - timedelta(days=today.weekday())
            current_week_end = current_week_start + timedelta(days=6)
            next_week_start = current_week_start + timedelta(days=7)
            next_week_end = next_week_start + timedelta(days=6)
            print(f"当前周: {current_week_start.strftime('%Y-%m-%d')} 到 {current_week_end.strftime('%Y-%m-%d')}")
            print(f"下周: {next_week_start.strftime('%Y-%m-%d')} 到 {next_week_end.strftime('%Y-%m-%d')}")
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA'
            ]
            mask = False
            for col in date_columns:
                if col in df1.columns:
                    col_mask = (
                        (df1[col] >= current_week_start) & (df1[col] <= current_week_end) |
                        (df1[col] >= next_week_start) & (df1[col] <= next_week_end)
                    )
                    mask = mask | col_mask
            df_filtered = df1[mask].copy()
            print(f"筛选日期范围后，剩余{len(df_filtered)}行数据")
            processed_rows = []
            for _, row in df_filtered.iterrows():
                project_has_valid_phases = False
                for col in date_columns:
                    if col in df1.columns and pd.notna(row[col]):
                        if (row[col] >= current_week_start) and (row[col] <= current_week_end) or \
                           (row[col] >= next_week_start) and (row[col] <= next_week_end):
                            new_row = row.copy()
                            new_row['当前状态'] = col
                            processed_rows.append(new_row)
                            project_has_valid_phases = True
                if not project_has_valid_phases:
                    new_row = row.copy()
                    new_row['当前状态'] = ""
                    processed_rows.append(new_row)
            if processed_rows:
                result_df = pd.DataFrame(processed_rows)
            else:
                result_df = df_filtered.copy()
                if '当前状态' not in result_df.columns:
                    result_df['当前状态'] = ''
            print(f"最终处理完成，共{len(result_df)}行数据（多phase项目已分行处理）")
            return result_df
        except Exception as e:
            print(f"筛选日期范围数据过程中出现错误: {e}")
            raise

    def run(self):
        """运行完整的爬取流程"""
        try:
            print("开始PKR数据爬取流程...")
            self.driver, self.wait = setup_driver(self.download_dir)
            PLMLogin(self.driver, self.wait)
            latest_file = DL_Project_List(self.driver, self.wait, self.download_dir)
            self.original_project_list_file = latest_file
            df1 = RD_Project_List(latest_file)
            df1 = self.filter_data_by_date_range(df1)
            df1 = self.add_missing_historical_projects(df1, self.original_project_list_file)
            df1 = self.update_ebp_leader_info(df1)
            self.navigate_to_pkr_management()
            df1 = self.extract_pkr_info(df1)
            df1 = self.remove_phases_beyond_next_sunday(df1)
            output_file = self.export_final_data(df1)
            try:
                print("\n开始同步数据到飞书表格...")
                from PKR未完成情况同步到飞书 import sync_pkr_data_to_feishu
                sync_pkr_data_to_feishu()
                print("飞书表格同步完成！")
            except Exception as e:
                print(f"飞书同步失败: {e}")
                import traceback
                traceback.print_exc()
                print("请手动运行 'PKR未完成情况同步到飞书.py' 进行同步")
            self.send_teams_message(df1)
            print("PKR数据爬取流程完成！")
            print(f"结果已保存到: {output_file}")
        except Exception as e:
            print(f"运行过程中出现错误: {e}")
            raise
        finally:
            if self.driver:
                self.driver.quit()

if __name__ == "__main__":
    crawler = PKRDataCrawler()
    crawler.run()