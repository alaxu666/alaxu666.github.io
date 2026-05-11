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
from webdriver_manager.chrome import ChromeDriverManager
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

# 导入配置文件
from config import *

class PKRDataCrawler:
    def __init__(self):
        self.driver = None
        self.wait = None
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
        """使用 Outlook COM 打开含 HTML 正文的邮件窗口。"""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.To = recipient_emails
            mail.CC = cc_emails
            mail.HTMLBody = html_content
            mail.Display(False)
            return True
        except Exception as e:
            print(f"Outlook COM 打开失败: {e}")
            return False

    def setup_driver(self):
        """设置Selenium WebDriver"""
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_experimental_option('prefs', {
            'download.default_directory': self.download_dir,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        })

        # 使用webdriver-manager自动下载和管理ChromeDriver
        from selenium.webdriver.chrome.service import Service

        try:
            # 尝试使用webdriver-manager（适用于有网络访问的环境）
            service = Service(ChromeDriverManager().install())
        except Exception as e:
            print(f"ChromeDriverManager下载失败: {e}")
            print("尝试使用本地ChromeDriver...")

            # 尝试使用本地已安装的ChromeDriver
            try:
                # 常见ChromeDriver路径
                possible_paths = [
                    'chromedriver.exe',  # 当前目录
                    r'C:\chromedriver\chromedriver.exe',  # 常见安装路径
                    r'C:\Program Files\chromedriver\chromedriver.exe',
                    r'C:\Windows\chromedriver.exe'
                ]

                driver_path = None
                for path in possible_paths:
                    if os.path.exists(path):
                        driver_path = path
                        break

                if driver_path:
                    print(f"找到本地ChromeDriver: {driver_path}")
                    service = Service(driver_path)
                else:
                    print("未找到本地ChromeDriver，尝试让Selenium自动查找...")
                    # 让Selenium自动查找ChromeDriver
                    service = Service()

            except Exception as local_e:
                print(f"本地ChromeDriver也失败: {local_e}")
                raise Exception("无法初始化ChromeDriver，请确保已安装Chrome浏览器和ChromeDriver")

        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 10)

    def login(self):
        """登录系统"""
        try:
            print(f"正在打开登录页面: {LOGIN_URL}")
            self.driver.get(LOGIN_URL)

            # 等待页面加载
            time.sleep(PAGE_LOAD_WAIT_TIME)

            # 检查页面标题
            print(f"当前页面标题: {self.driver.title}")

            # 输入用户名
            print("正在定位用户名输入框...")
            username_input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']"))
            )
            username_input.clear()
            username_input.send_keys(USERNAME)
            print("已输入用户名")

            # 输入密码
            print("正在定位密码输入框...")
            password_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='Password']")
            password_input.clear()
            password_input.send_keys(PASSWORD)
            print("已输入密码")

            # 点击登录按钮
            print("正在定位登录按钮...")
            login_button = self.driver.find_element(By.CSS_SELECTOR, "button.sw-button.accent-caution")
            login_button.click()
            print("已点击登录按钮")

            # 等待页面跳转
            time.sleep(PAGE_LOAD_WAIT_TIME * 2)

            # 验证登录成功
            print("正在验证登录状态...")
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[icon-id='homeEdit']"))
            )
            print("登录成功！")

        except Exception as e:
            print(f"登录过程中出现错误: {e}")
            print(f"当前URL: {self.driver.current_url}")
            print(f"当前标题: {self.driver.title}")
            raise

    def navigate_to_xso_management(self):
        """导航到XSO Management页面"""
        try:
            # 点击XSO & PKR
            xso_pkr_div = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='XSO & PKR']"))
            )
            xso_pkr_div.click()
            print("已点击XSO & PKR")

            # 等待页面加载
            time.sleep(PAGE_LOAD_WAIT_TIME * 2)

            # 定位iframeContent
            iframe_content = self.wait.until(
                EC.presence_of_element_located((By.ID, "iframeContent"))
            )

            # 等待iframe内容加载
            time.sleep(PAGE_LOAD_WAIT_TIME)

            # 切换到iframe
            iframe_content = self.driver.find_element(By.ID, "iframeContent")
            self.driver.switch_to.frame(iframe_content)
            print("已切换到iframeContent")

            # 展开XSO Management - 增加重试机制
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # 尝试多种定位方式
                    try:
                        xso_management = self.wait.until(
                            EC.element_to_be_clickable((By.LINK_TEXT, "XSO Management"))
                        )
                    except:
                        # 如果LINK_TEXT失败，尝试XPath
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

            # 点击Program Dashboard
            program_dashboard = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Program Dashboard"))
            )
            program_dashboard.click()
            print("已点击Program Dashboard")

            # 等待iframe加载
            time.sleep(PAGE_LOAD_WAIT_TIME * 2)

        except Exception as e:
            print(f"导航到XSO Management过程中出现错误: {e}")
            # 打印页面源码帮助调试
            try:
                print("当前页面源码片段:")
                print(self.driver.page_source[:1000])
            except:
                pass
            raise

    def wait_for_project_list_download(self, existing_files=None, timeout=120, poll_interval=2):
        """等待最新Project_List下载完成并返回文件路径"""
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
        """下载XSO数据"""
        try:
            # 切换到目标iframe
            iframe = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-1-1"))
            )
            self.driver.switch_to.frame(iframe)
            print("已切换到XSO iframe")

            # 选择BU
            bu_select = Select(self.driver.find_element(By.NAME, "select_bu"))
            bu_select.select_by_value(BU_VALUE)
            print(f"已选择BU: {BU_VALUE}")

            # 选择Category
            category_select = Select(self.driver.find_element(By.NAME, "select_Category"))
            category_select.select_by_value(CATEGORY_VALUE)
            print(f"已选择Category: {CATEGORY_VALUE}")

            # 输入Product Group
            product_group_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='Product Group']")
            product_group_input.clear()
            product_group_input.send_keys(PRODUCT_GROUP)
            print(f"已输入Product Group: {PRODUCT_GROUP}")

            # 点击搜索按钮
            search_button = self.driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']")
            search_button.click()
            print("已点击搜索按钮")

            # 等待搜索结果加载
            time.sleep(PAGE_LOAD_WAIT_TIME)

            # 记录当前已有Project_List文件，避免误判旧文件
            existing_files = set(glob.glob(os.path.join(self.download_dir, "Project_List*.xls*")))

            # 点击导出按钮（实际上是span元素）
            export_button = self.driver.find_element(By.CSS_SELECTOR, "span[onclick='exportData()']")
            export_button.click()
            print("已点击导出按钮")

            # 等待下载文件出现并完成
            download_file = self.wait_for_project_list_download(existing_files=existing_files, timeout=180)
            print(f"下载完成: {download_file}")

            # 切换回默认内容
            self.driver.switch_to.default_content()

        except Exception as e:
            print(f"下载XSO数据过程中出现错误: {e}")
            raise

    def get_latest_downloaded_file(self):
        """获取最新下载的Project_List文件"""
        try:
            # 查找所有Project_List开头的xls文件
            pattern = os.path.join(self.download_dir, "Project_List*.xls*")
            files = glob.glob(pattern)

            if not files:
                raise FileNotFoundError("未找到Project_List文件")

            # 按修改时间排序，获取最新的文件
            latest_file = max(files, key=os.path.getmtime)
            print(f"找到最新文件: {latest_file}")
            return latest_file

        except Exception as e:
            print(f"获取最新下载文件过程中出现错误: {e}")
            raise

    def process_excel_data(self, file_path):
        """处理Excel数据"""
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            print(f"已读取Excel文件，共{len(df)}行数据")

            # 获取当前日期和本周、下周的日期范围
            today = datetime.now()
            current_week_start = today - timedelta(days=today.weekday())
            current_week_end = current_week_start + timedelta(days=6)
            next_week_start = current_week_start + timedelta(days=7)
            next_week_end = next_week_start + timedelta(days=6)

            print(f"当前周: {current_week_start.strftime('%Y-%m-%d')} 到 {current_week_end.strftime('%Y-%m-%d')}")
            print(f"下周: {next_week_start.strftime('%Y-%m-%d')} 到 {next_week_end.strftime('%Y-%m-%d')}")

            # 筛选Category为"S"的数据
            df = df[df['Category'] == 'S'].copy()
            print(f"筛选Category=S后，剩余{len(df)}行数据")

            # 定义日期列
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA'
            ]

            # 转换日期列
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')

            # 筛选日期在本周或下周的数据
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

            # 为每个Project ID的多个phase创建多行数据
            expanded_rows = []

            for idx, row in df_filtered.iterrows():
                project_id = row['Project ID']

                # 检查每个phase的日期是否在本周或下周
                for col in date_columns:
                    if col in df_filtered.columns and pd.notna(row[col]):
                        if (current_week_start <= row[col] <= current_week_end) or \
                           (next_week_start <= row[col] <= next_week_end):

                            # 创建新行，复制原数据
                            new_row = row.copy()
                            new_row['当前状态'] = col  # 设置当前状态为对应的phase
                            expanded_rows.append(new_row)

            # 创建新的DataFrame
            if expanded_rows:
                df1 = pd.DataFrame(expanded_rows)
                print(f"展开多phase数据后，共{len(df1)}行数据")
            else:
                # 如果没有找到符合条件的数据，创建空DataFrame
                df1 = df_filtered.copy()
                df1['当前状态'] = ''
                print("未找到符合条件的数据")

            # 添加PKR信息列、Confirmed列和是否确认列
            df1['PKR信息'] = ''
            df1['Confirmed'] = ''
            df1['是否确认'] = ''
            df1['未完成人名'] = ''
            df1.insert(0, 'id', range(1, len(df1) + 1))   # 新增：添加唯一id列

            return df1

        except Exception as e:
            print(f"处理Excel数据过程中出现错误: {e}")
            raise

    def extract_ebp_leader_info(self, df1):
        """提取EBP Leader信息"""
        try:
            print("正在提取EBP Leader信息...")

            # 切换回主文档
            self.driver.switch_to.default_content()
            print("已切换回主文档")

            # 等待页面加载
            time.sleep(PAGE_LOAD_WAIT_TIME)

            # 定位iframeContent
            iframe_content = self.wait.until(
                EC.presence_of_element_located((By.ID, "iframeContent"))
            )

            # 切换到iframe
            self.driver.switch_to.frame(iframe_content)
            print("已切换到iframeContent")

            # 展开XSO Management
            xso_management = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='XSO Management']/parent::a"))
            )
            xso_management.click()
            print("已展开XSO Management")

            # 点击Program Dashboard
            program_dashboard = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Program Dashboard']/parent::a"))
            )
            program_dashboard.click()
            print("已点击Program Dashboard")

            # 等待iframe加载
            time.sleep(PAGE_LOAD_WAIT_TIME * 2)

            # 切换到XSO iframe
            xso_iframe = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-1-1"))
            )
            self.driver.switch_to.frame(xso_iframe)
            print("已切换到XSO iframe")

            # 设置筛选条件（与下载数据时保持一致）
            # 选择BU
            bu_select = Select(self.driver.find_element(By.NAME, "select_bu"))
            bu_select.select_by_value(BU_VALUE)
            print(f"已选择BU: {BU_VALUE}")

            # 选择Category
            category_select = Select(self.driver.find_element(By.NAME, "select_Category"))
            category_select.select_by_value(CATEGORY_VALUE)
            print(f"已选择Category: {CATEGORY_VALUE}")

            # 输入Product Group
            product_group_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='Product Group']")
            product_group_input.clear()
            product_group_input.send_keys(PRODUCT_GROUP)
            print(f"已输入Product Group: {PRODUCT_GROUP}")

            # 执行初始搜索以加载数据表格
            initial_search_button = self.driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']")
            initial_search_button.click()
            print("已执行初始搜索")
            time.sleep(PAGE_LOAD_WAIT_TIME)

            # 筛选出EBP Leader为空的项目
            empty_ebp_projects = df1[
                df1['EBP Leader'].isna() |
                (df1['EBP Leader'].astype(str).str.strip() == '')
            ]

            print(f"找到{len(empty_ebp_projects)}个EBP Leader为空的项目")

            if len(empty_ebp_projects) == 0:
                print("没有需要补充EBP Leader信息的项目")
                return df1

            # 记录处理结果统计
            success_count = 0
            failed_count = 0
            failed_projects = []

            # 对每个EBP Leader为空的项目进行处理
            for idx, row in empty_ebp_projects.iterrows():
                project_id = row['Project ID']
                print(f"处理Project ID: {project_id}的EBP Leader信息")

                try:
                    # 定位搜索输入框
                    search_input = self.wait.until(
                        EC.presence_of_element_located((By.ID, "searchProject"))
                    )
                    search_input.clear()
                    search_input.send_keys(str(project_id))
                    print(f"已输入Project ID: {project_id}")

                    # 点击搜索按钮
                    search_button = self.driver.find_element(
                        By.CSS_SELECTOR, "button[onclick='searchproject()']"
                    )
                    search_button.click()

                    # 等待搜索结果加载完成 - 等待表格出现
                    try:
                        self.wait.until(
                            EC.presence_of_element_located((By.ID, "grid-table"))
                        )
                        time.sleep(2)  # 额外等待确保数据完全加载
                    except TimeoutException:
                        print(f"Project ID {project_id}搜索结果表格加载超时")
                        continue

                    # 查找EBP Leader元素
                    ebp_elements = self.driver.find_elements(
                        By.CSS_SELECTOR, "td[aria-describedby='grid-table_real_pm_name']"
                    )

                    if ebp_elements:
                        ebp_title = ebp_elements[0].get_attribute("title")
                        if ebp_title:
                            df1.at[idx, 'EBP Leader'] = ebp_title.strip()
                            print(f"Project ID {project_id}的EBP Leader已更新为: {ebp_title.strip()}")
                            success_count += 1
                        else:
                            print(f"Project ID {project_id}的EBP Leader元素没有title属性，元素数量: {len(ebp_elements)}")
                            # 尝试获取文本内容作为备选
                            ebp_text = ebp_elements[0].text.strip()
                            if ebp_text:
                                df1.at[idx, 'EBP Leader'] = ebp_text
                                print(f"Project ID {project_id}使用文本内容作为EBP Leader: {ebp_text}")
                                success_count += 1
                            else:
                                failed_count += 1
                                failed_projects.append(project_id)
                    else:
                        print(f"Project ID {project_id}未找到EBP Leader元素，可能搜索无结果")
                        # 检查是否有"无数据"或类似提示
                        no_data_elements = self.driver.find_elements(By.CSS_SELECTOR, ".ui-jqgrid-btable .nodata")
                        if no_data_elements:
                            print(f"Project ID {project_id}搜索结果: 无数据")
                        else:
                            print(f"Project ID {project_id}搜索结果表格可能为空或结构不同")
                        failed_count += 1
                        failed_projects.append(project_id)

                except Exception as e:
                    print(f"处理Project ID {project_id}时出错: {e}")
                    failed_count += 1
                    failed_projects.append(project_id)
                    continue

            # 输出统计结果
            print(f"\nEBP Leader提取统计:")
            print(f"成功提取: {success_count} 个")
            print(f"提取失败: {failed_count} 个")
            if failed_projects:
                print(f"失败的Project ID: {', '.join(map(str, failed_projects))}")

            # 切换回默认内容
            self.driver.switch_to.default_content()
            print("EBP Leader信息提取完成")

            return df1

        except Exception as e:
            print(f"提取EBP Leader信息过程中出现错误: {e}")
            return df1

    def navigate_to_pkr_management(self):
        """导航到PKR Management页面"""
        try:
            # 切换回主文档
            self.driver.switch_to.default_content()
            print("已切换回主文档")

            # 等待页面加载
            time.sleep(PAGE_LOAD_WAIT_TIME)

            # 定位iframeContent
            iframe_content = self.wait.until(
                EC.presence_of_element_located((By.ID, "iframeContent"))
            )

            # 切换到iframe
            self.driver.switch_to.frame(iframe_content)
            print("已切换到iframeContent")

            # 展开PKR Management
            pkr_management = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "PKR Management"))
            )
            pkr_management.click()
            print("已展开PKR Management")

            # 点击Program PKR Summary
            pkr_summary = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Program PKR Summary"))
            )
            pkr_summary.click()
            print("已点击Program PKR Summary")

            # 等待iframe加载
            time.sleep(PAGE_LOAD_WAIT_TIME)

        except Exception as e:
            print(f"导航到PKR Management过程中出现错误: {e}")
            raise

    def extract_pkr_info(self, df1):
        """提取PKR信息 - 基于id逐行处理，避免重复Project ID导致错乱"""
        try:
            # ========== 关键修改：重置索引，保证连续性 ==========
            df1 = df1.reset_index(drop=True)
            # 如果id列不连续或需要重新编号（可选，不影响逻辑）
            # df1['id'] = range(1, len(df1) + 1)
            # =================================================

            # 切换到PKR iframe
            iframe = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-2-2"))
            )
            self.driver.switch_to.frame(iframe)
            print("已切换到PKR iframe")

            # 逐行处理（idx 是重置后的连续索引）
            for idx, row in df1.iterrows():
                project_id = row['Project ID']
                current_status = row['当前状态']
                row_id = row.get('id', idx+1)  # 获取id列值，若没有则用idx+1
                print(f"处理行索引={idx}, id={row_id}, Project ID: {project_id}, 当前状态: {current_status}")

                # Phase 1 直接跳过，PKR信息保持为空
                if current_status == 'Phase 1 Gate Exit-GO':
                    print(f"Project ID {project_id} 当前状态为Phase 1，跳过PKR提取")
                    # 确保这些列被清空（防止历史残留）
                    df1.at[idx, 'PKR信息'] = ''
                    df1.at[idx, 'Confirmed'] = ''
                    df1.at[idx, '是否确认'] = ''
                    continue

                # 根据当前状态确定对应的 aria-describedby 属性
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
                    # 1. 输入Project ID
                    search_input = self.wait.until(
                        EC.presence_of_element_located((By.ID, "searchProject"))
                    )
                    search_input.clear()
                    search_input.send_keys(str(project_id))

                    # 2. 点击搜索按钮
                    search_button = self.driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']")
                    search_button.click()

                    # 等待搜索结果加载
                    time.sleep(PAGE_LOAD_WAIT_TIME)

                    # 3. 点击class="lbl"的span
                    lbl_span = self.wait.until(
                        EC.presence_of_element_located((By.CLASS_NAME, "lbl"))
                    )
                    self.driver.execute_script("arguments[0].click();", lbl_span)
                    time.sleep(2)  # 等待弹窗数据展开

                    # 4. 查找对应 aria-describedby 的 td 标签
                    td_elements = self.driver.find_elements(
                        By.CSS_SELECTOR, f"td[aria-describedby='{target_aria}']"
                    )

                    pkr_info = ""
                    confirmed_info = ""

                    # 5. 遍历每个td，提取div_score
                    for td in td_elements:
                        div_scores = td.find_elements(By.CLASS_NAME, "div_score")
                        for div in div_scores:
                            # 提取PKR名称
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

                            # 提取Confirmed信息（i标签）
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

                    # 格式化PKR信息
                    formatted_pkr_info = ""
                    if pkr_info and pkr_info.strip():
                        lines = pkr_info.split('\n')
                        formatted_lines = []
                        for line in lines:
                            if line.strip():
                                parts = line.split(',', 1)  # 只分割第一个逗号
                                if len(parts) == 2:
                                    pkr_name, pkr_score = parts
                                    # 替换格式
                                    if pkr_name == "Total":
                                        # Total改为"最低打分："，分号后加换行符
                                        formatted_line = f"最低打分：{pkr_score}；\n"
                                    else:
                                        # 将PKRname和PKRscore之间的逗号替换为"的PKR打分："，分号后加换行符
                                        formatted_line = f"{pkr_name}的PKR打分：{pkr_score}；\n"

                                    formatted_lines.append(formatted_line)

                        formatted_pkr_info = ''.join(formatted_lines)

                    # 保存结果到df1的当前行（使用重置后的idx）
                    df1.at[idx, 'PKR信息'] = formatted_pkr_info
                    df1.at[idx, 'Confirmed'] = confirmed_info

                    # 判断是否确认状态（基于Confirmed列数据）
                    if confirmed_info and confirmed_info.strip():
                        confirmed_lines = [line.strip() for line in confirmed_info.split('\n') if line.strip()]

                        if not confirmed_lines:
                            df1.at[idx, '是否确认'] = ''
                        else:
                            # 分析每一行的状态
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

                            # 根据分析结果设置确认状态
                            if all_confirmed and not has_other:
                                df1.at[idx, '是否确认'] = '完成确认'
                            elif all_submitted and not has_other:
                                df1.at[idx, '是否确认'] = '未确认(Cao liang)'
                            elif has_other:
                                df1.at[idx, '是否确认'] = '未评分(工程师)'
                            else:
                                df1.at[idx, '是否确认'] = '未确认(Cao liang)'  # 混合状态
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
        """分析未完成PKR信息 - 基于PKR score"""
        try:
            print("正在分析未完成PKR信息...")

            def extract_incomplete_pkr_by_score(row):
                # 检查当前状态是否为Phase 1
                current_status = row.get('当前状态', '')
                if current_status == 'Phase 1 Gate Exit-GO':
                    return "phase1不适用"

                pkr_info = row.get('PKR信息', '')
                if pd.isna(pkr_info) or not str(pkr_info).strip():
                    return "完成"

                incomplete_pkr = []

                # 解析格式化后的PKR信息
                # 格式示例: "最低打分：100；\nWang Hao23的PKR打分：100；\nLu Dinghao的PKR打分：95；\n"
                pkr_text = str(pkr_info).strip()

                # 按换行符分割，然后处理每一行
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

                    # 跳过"最低打分"项，只处理具体的PKR人员
                    if "最低打分：" in item:
                        continue

                    # 解析"PKRname的PKR打分：score"格式
                    if "的PKR打分：" in item:
                        try:
                            # 分割PKR名称和分数
                            name_part, score_part = item.split("的PKR打分：", 1)
                            pkr_name = name_part.strip()
                            pkr_score_str = score_part.strip()

                            # 提取分数（应该是分号前的数字）
                            pkr_score_str_clean = pkr_score_str.split('；')[0]  # 移除分号
                            pkr_score = int(pkr_score_str_clean)

                            # 如果分数≠100，则添加到未完成列表
                            if pkr_score != 100:
                                # 保持与PKR信息相同的格式
                                incomplete_pkr.append(f"{pkr_name}的PKR打分：{pkr_score}；\n")
                        except (ValueError, IndexError):
                            # 如果解析失败，跳过此项
                            continue

                if incomplete_pkr:
                    # 使用与PKR信息相同的格式，每行后都有换行符
                    return ''.join(incomplete_pkr).strip()
                else:
                    return "完成"

            # 添加未完成PKR信息列
            df1['未完成PKR信息'] = df1.apply(extract_incomplete_pkr_by_score, axis=1)

            print(f"已完成{len(df1)}行数据的未完成PKR信息分析")

            # 统计信息
            incomplete_count = len(df1[df1['未完成PKR信息'] != "完成"])
            complete_count = len(df1[df1['未完成PKR信息'] == "完成"])
            print(f"包含未完成PKR的项目数: {incomplete_count}")
            print(f"PKR全部完成的项目数: {complete_count}")

            return df1

        except Exception as e:
            print(f"分析未完成PKR信息过程中出现错误: {e}")
            # 如果分析失败，添加默认列
            df1['未完成PKR信息'] = "分析失败"
            return df1

    def sort_data_by_date(self, df1):
        """按日期从早到晚排序数据"""
        try:
            print("正在按日期排序数据...")

            # 定义日期列
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA',
                'Phase 5 Gate Exit-PLR'
            ]

            # 转换日期列并找到每个项目的最早日期
            df1['最早日期'] = pd.NaT

            for idx, row in df1.iterrows():
                earliest_date = None
                for col in date_columns:
                    if col in df1.columns and pd.notna(row[col]):
                        if earliest_date is None or row[col] < earliest_date:
                            earliest_date = row[col]

                df1.at[idx, '最早日期'] = earliest_date

            # 按最早日期排序（从早到晚）
            df1_sorted = df1.sort_values(by='最早日期', na_position='last')

            # 删除临时列
            if '最早日期' in df1_sorted.columns:
                df1_sorted = df1_sorted.drop('最早日期', axis=1)

            print(f"数据排序完成，共{len(df1_sorted)}行数据")
            return df1_sorted

        except Exception as e:
            print(f"数据排序过程中出现错误: {e}")
            return df1

    def manage_historical_pkr_records(self, df1):
        """管理PKR历史未完成记录"""
        try:
            historical_file = os.path.join(self.confirm_dir, "PKR历史未完成记录.xlsx")

            # 读取现有的历史记录
            historical_df = None
            if os.path.exists(historical_file):
                try:
                    historical_df = pd.read_excel(historical_file)
                    print(f"读取到{len(historical_df)}条历史未完成记录")
                except Exception as e:
                    print(f"读取历史记录文件失败: {e}")
                    historical_df = None

            # 从当前数据中提取未完成确认的项目（基于是否确认字段）
            # 同时排除"当前状态"为"Phase 1 Gate Exit-GO"的项目
            incomplete_projects = df1[
                (df1['是否确认'] != "完成确认") &
                (df1['当前状态'] != 'Phase 1 Gate Exit-GO')
            ].copy()

            if historical_df is not None and len(historical_df) > 0:
                # 合并历史记录和当前未完成项目
                # 使用Project ID作为唯一标识
                all_incomplete = pd.concat([historical_df, incomplete_projects], ignore_index=True)

                # 去重，保留最新的记录
                all_incomplete = all_incomplete.drop_duplicates(subset=['Project ID'], keep='last')

                print(f"合并后共有{len(all_incomplete)}条未完成记录")
            else:
                all_incomplete = incomplete_projects
                print(f"新增{len(all_incomplete)}条未完成记录")

            # 注意：不在主输出中添加历史记录，只在历史记录文件中维护
            # 历史记录文件会单独保存，不会混入主输出

            return df1, all_incomplete, historical_file

        except Exception as e:
            print(f"管理历史记录过程中出现错误: {e}")
            return df1, pd.DataFrame(), ""

    def add_gate_week_info(self, df1):
        """添加过门周信息"""
        try:
            print("正在添加过门周信息...")

            # 获取当前日期和本周、下周的日期范围
            today = datetime.now()
            current_week_start = today - timedelta(days=today.weekday())
            current_week_end = current_week_start + timedelta(days=6)
            next_week_start = current_week_start + timedelta(days=7)
            next_week_end = next_week_start + timedelta(days=6)

            # 定义日期列
            date_columns = [
                'Phase 1 Gate Exit-GO',
                'Phase 2 Gate Exit-DVR',
                'Phase 3 Gate Exit-FPR',
                'Phase 4 Gate Exit-CPA',
                'Phase 5 Gate Exit-PLR'
            ]

            # 添加过门周列
            df1['过门周'] = ''

            for idx, row in df1.iterrows():
                gate_week = ''

                # 检查每个日期列
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

    def update_historical_records(self, historical_df, current_df, historical_file):
        """更新历史记录文件"""
        try:
            if historical_df is None or len(historical_df) == 0:
                return

            print("正在更新历史记录...")

            # 从当前数据中找出已完成的历史项目
            completed_projects = current_df[
                (current_df['未完成PKR信息'] == "完成") &
                (current_df['Project ID'].isin(historical_df['Project ID']))
            ]

            if len(completed_projects) > 0:
                completed_ids = set(completed_projects['Project ID'])
                print(f"发现{len(completed_ids)}个历史项目已完成PKR")

                # 从历史记录中移除已完成的项目
                updated_historical = historical_df[~historical_df['Project ID'].isin(completed_ids)]
                removed_count = len(historical_df) - len(updated_historical)
                print(f"从历史记录中移除{removed_count}个已完成项目")
            else:
                updated_historical = historical_df
                print("没有发现新的已完成项目")

            # 保存更新后的历史记录
            if len(updated_historical) > 0:
                updated_historical.to_excel(historical_file, index=False)
                print(f"历史记录已更新，保存{len(updated_historical)}条记录到: {historical_file}")
            else:
                # 如果历史记录为空，删除文件
                if os.path.exists(historical_file):
                    os.remove(historical_file)
                    print("历史记录已清空，删除文件")

        except Exception as e:
            print(f"更新历史记录过程中出现错误: {e}")

    def extract_unique_pkr_names(self, df1):
        """提取未完成PKR信息中的PKR名称，返回邮箱格式字符串（不修改df1）"""
        try:
            print("正在提取未完成人名...")
            incomplete_pkr_names = set()

            # 只处理"是否确认"列中不是"完成确认"的数据
            for idx, row in df1.iterrows():
                confirm_status = row.get('是否确认', '')
                if confirm_status != "完成确认":
                    incomplete_pkr_info = row.get('未完成PKR信息', '')
                    if pd.notna(incomplete_pkr_info) and str(incomplete_pkr_info).strip() \
                            and incomplete_pkr_info != "完成" and incomplete_pkr_info != "phase1不适用":

                        # 解析新的PKR信息格式
                        # 格式示例: "Lu Dinghao的PKR打分：95；\nXu Shirong的PKR打分：90；"
                        pkr_text = str(incomplete_pkr_info).strip()

                        # 按换行符分割
                        pkr_lines = pkr_text.split('\n')
                        for line in pkr_lines:
                            line = line.strip()
                            if line and "的PKR打分：" in line:
                                # 提取PKR名称
                                pkr_name = line.split("的PKR打分：")[0].strip()
                                if pkr_name and pkr_name != 'Total':
                                    # 特殊处理：Wang Hao23 -> Wang23 Hao
                                    if pkr_name == "Wang Hao23":
                                        pkr_name = "Wang23 Hao"
                                    incomplete_pkr_names.add(pkr_name)

            pkr_names_list = sorted(list(incomplete_pkr_names))
            if not pkr_names_list:
                print("未找到未完成PKR信息中的PKR名称")
                return ""

            print(f"从未完成PKR信息中找到{len(pkr_names_list)}个PKR名称: {', '.join(pkr_names_list)}")

            # 生成邮箱格式：名.姓@yanfeng.com
            email_addresses = []
            for pkr_name in pkr_names_list:
                # 分割姓和名
                name_parts = pkr_name.split(' ')
                if len(name_parts) >= 2:
                    # 假设格式为"姓 名"，转换为"名.姓@yanfeng.com"
                    last_name = name_parts[0]  # 姓
                    first_name = name_parts[1]  # 名
                    email = f"{first_name}.{last_name}@yanfeng.com"
                    email_addresses.append(email)
                else:
                    # 如果无法分割，使用原格式
                    email = f"{pkr_name.replace(' ', '.')}@yanfeng.com"
                    email_addresses.append(email)

            # 按邮箱地址排序
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

    def export_final_data(self, df1):
        """导出最终数据"""
        try:
            # 分析未完成PKR信息
            df1 = self.analyze_incomplete_pkr(df1)

            # 管理历史记录
            df1, historical_df, historical_file = self.manage_historical_pkr_records(df1)

            # 添加过门周信息
            df1 = self.add_gate_week_info(df1)

            # 按日期排序（从早到晚）
            df1 = self.sort_data_by_date(df1)

            # ---------- 关键修改：提取邮箱字符串并写入第一行 ----------
            email_string = self.extract_unique_pkr_names(df1)   # 只获取字符串，不改df1
            # 清空整个“未完成人名”列（确保只有第一行有值）
            df1['未完成人名'] = ''
            if email_string and len(df1) > 0:
                df1.iloc[0, df1.columns.get_loc('未完成人名')] = email_string
            # ------------------------------------------------------

            # 选择需要保留的列
            columns_to_keep = OUTPUT_COLUMNS + ['未完成PKR信息', '过门周', 'Confirmed', '是否确认', '未完成人名']
            existing_columns = [col for col in columns_to_keep if col in df1.columns]

            # 1. 导出PKR完成情况（近两周）.xlsx - 所有数据
            complete_df = df1[existing_columns].copy()
            complete_file = os.path.join(self.confirm_dir, "PKR完成情况（近两周）.xlsx")
            complete_df.to_excel(complete_file, index=False)
            print(f"完整数据已导出到: {complete_file}")
            print(f"共导出{len(complete_df)}行完整数据")

            # 2. 筛选出未完成的PKR数据（与邮件发送的筛选条件一致）
            # 条件："是否确认"列不是"完成确认"，且"当前状态"列不是"Phase 1 Gate Exit-GO"
            incomplete_mask = (
                (df1['是否确认'] != "完成确认") &
                (df1['当前状态'] != 'Phase 1 Gate Exit-GO')
            )
            incomplete_df = df1[incomplete_mask].copy()
            incomplete_final_df = incomplete_df[existing_columns].copy()

            # 导出未完成数据
            incomplete_file = os.path.join(self.confirm_dir, "PKR未完成情况（近两周）.xlsx")
            incomplete_final_df.to_excel(incomplete_file, index=False)
            print(f"未完成数据已导出到: {incomplete_file}")
            print(f"共导出{len(incomplete_final_df)}行未完成数据")
            print(f"筛选条件: 是否确认≠'完成确认' 且 当前状态≠'Phase 1 Gate Exit-GO'")

            # 返回未完成文件的路径（用于后续同步到飞书）
            output_file = incomplete_file

            # 导出历史未完成记录（如有）
            if len(historical_df) > 0:
                historical_df.to_excel(historical_file, index=False)
                print(f"历史未完成记录已保存到: {historical_file}")

            # 更新历史记录（移除已完成的项目）
            self.update_historical_records(historical_df, df1, historical_file)

            return output_file

        except Exception as e:
            print(f"导出数据过程中出现错误: {e}")
            raise

    def send_teams_message(self, df1):
        """自动发送邮件，包含未完成确认的数据"""
        try:
            # 首先更新"是否确认"列
            for idx, row in df1.iterrows():
                incomplete_pkr = str(row.get('未完成PKR信息', '')).strip()
                confirmed_info = str(row.get('Confirmed', '')).strip()

                if incomplete_pkr == "完成":
                    # 检查Confirmed列是否全部confirmed
                    if confirmed_info and "Confirmed" in confirmed_info:
                        # 检查是否所有PKR都confirmed
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
                    # 如果未完成PKR信息不是"完成"，则显示"未递交(工程师)"
                    df1.at[idx, '是否确认'] = '未递交(工程师)'

            # 筛选出需要关注的数据（排除"完成确认"的）
            incomplete_df = df1[df1['是否确认'] != '完成确认'].copy()

            # 删除"当前状态"为"Phase 1 Gate Exit-GO"的行
            incomplete_df = incomplete_df[incomplete_df['当前状态'] != 'Phase 1 Gate Exit-GO']

            if len(incomplete_df) == 0:
                print("没有需要关注的数据，不需要发送邮件")
                return

            print(f"找到{len(incomplete_df)}条需要关注的数据")

            # 获取收件人（未完成人名第一行）并格式化邮箱地址
            recipient_emails = ""
            if '未完成人名' in df1.columns and len(df1) > 0:
                raw_emails = str(df1.iloc[0].get('未完成人名', '')).strip()

                if raw_emails:
                    # 分割邮箱地址（支持逗号和分号分隔）
                    import re
                    email_list = [email.strip() for email in re.split(r'[;,]', raw_emails) if email.strip()]
                    formatted_emails = []

                    for email in email_list:
                        if '@yanfeng.com' in email:
                            # 提取邮箱用户名部分
                            username = email.split('@')[0]
                            # 将Lu.Gongchao格式转换为Gongchao.lu
                            if '.' in username:
                                parts = username.split('.')
                                if len(parts) == 2:
                                    # 首字母大写，其余小写
                                    first_name = parts[1].capitalize()
                                    last_name = parts[0].capitalize()
                                    formatted_username = f"{first_name}.{last_name}"
                                    formatted_email = f"{formatted_username}@yanfeng.com"
                                    formatted_emails.append(formatted_email)
                                else:
                                    formatted_emails.append(email)
                            else:
                                formatted_emails.append(email)
                        else:
                            formatted_emails.append(email)

                    # 用分号连接格式化后的邮箱地址
                    recipient_emails = '; '.join(formatted_emails)

            if not recipient_emails:
                recipient_emails = RECIPIENT_EMAIL  # 如果没有未完成人名，使用默认收件人

            # 创建邮件
            msg = MIMEMultipart('alternative')
            msg['Subject'] = 'PKR确认状态未完成提醒'
            msg['From'] = SENDER_EMAIL
            msg['To'] = recipient_emails
            msg['Cc'] = "liang.cao@yanfeng.com"

            # 创建HTML版本的消息内容（包含四列）
            html_content = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }}
                    table {{ border-collapse: collapse; width: 100%; margin: 20px 0; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
                    th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                    th {{ background-color: #4CAF50; color: white; font-weight: bold; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                    tr:hover {{ background-color: #e8f5e8; }}
                    tr.urgent {{ background-color: #ffebee; }}
                    tr.warning {{ background-color: #fff3e0; }}
                    h3 {{ color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; }}
                    .footer {{ margin-top: 20px; font-size: 12px; color: #666; }}
                    .status-submit {{ color: #d32f2f; font-weight: bold; }}
                    .status-confirm {{ color: #f57c00; font-weight: bold; }}
                </style>
            </head>
            <body>
            <h3>PKR确认状态未完成提醒</h3>
            <table>
            <tr>
                <th>Project Name</th>
                <th>当前状态</th>
                <th>未完成PKR信息</th>
                <th>是否确认</th>
            </tr>
            """

            for _, row in incomplete_df.iterrows():
                project_name = str(row.get('Project Name', '')).strip()
                current_status = str(row.get('当前状态', '')).strip()
                incomplete_pkr = str(row.get('未完成PKR信息', '')).strip()
                confirm_status = str(row.get('是否确认', '')).strip()

                if not incomplete_pkr or incomplete_pkr == "nan":
                    incomplete_pkr = "无"

                # 根据状态设置行样式
                row_class = ""
                status_class = ""
                if confirm_status == "未递交(工程师)":
                    row_class = "urgent"
                    status_class = "status-submit"
                elif confirm_status == "未确认(cao liang)":
                    row_class = "warning"
                    status_class = "status-confirm"

                html_content += f"""
            <tr class="{row_class}">
                <td>{project_name}</td>
                <td>{current_status}</td>
                <td>{incomplete_pkr}</td>
                <td class="{status_class}">{confirm_status}</td>
            </tr>
                """

            html_content += f"""
            </table>
            <div class="footer">
                <p>总计 {len(incomplete_df)} 个项目需要关注</p>
                <p><span class="status-submit">未递交(工程师)</span> - 需要工程师提交PKR</p>
                <p><span class="status-confirm">未确认(cao liang)</span> - 需要曹亮确认</p>
            </div>
            </body>
            </html>
            """

            # 只添加HTML版本，不添加纯文本版本
            html_part = MIMEText(html_content, 'html', 'utf-8')
            msg.attach(html_part)

            # 发送邮件
            try:
                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                server.starttls()  # 启用TLS加密
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.send_message(msg)
                server.quit()
                print(f"邮件已成功发送到 {recipient_emails}，抄送 liang.cao@yanfeng.com")
            except Exception as e:
                print(f"SMTP发送失败，尝试使用备选方案: {e}")
                print("正在使用备选方案打开邮件客户端...")

                # 设置邮件参数
                subject = "PKR确认状态未完成提醒"
                cc_email = "liang.cao@yanfeng.com"

                # 尝试使用 Outlook COM 直接打开邮件窗口
                if self.open_outlook_mail_window(recipient_emails, cc_email, subject, html_content):
                    print("已在 Outlook 中打开邮件窗口，请检查后发送")
                    return
                print("Outlook COM 打开失败，正在使用原始备选方案...")

                # 复制HTML内容到剪贴板，优先使用HTML格式
                clipboard_success = False
                if self.set_clipboard_html(html_content):
                    clipboard_success = True
                    print("已将HTML表格复制到剪贴板（HTML格式）")
                elif HAS_PYPERCLIP:
                    try:
                        pyperclip.copy(html_content)
                        clipboard_success = True
                        print("已将HTML内容复制到剪贴板（纯文本备选）")
                    except Exception as clipboard_error:
                        print(f"复制到剪贴板失败: {clipboard_error}")
                else:
                    print("无法写入HTML剪贴板，且pyperclip未安装，跳过剪贴板复制")

                # 创建HTML文件
                html_file = os.path.join(self.confirm_dir, "PKR未完成提醒.html")
                with open(html_file, 'w', encoding='utf-8') as f:
                    f.write(html_content)

                # 构建邮件URL（包含抄送）
                subject_encoded = urllib.parse.quote("PKR确认状态未完成提醒")

                # 简化邮件正文，只提示用户从浏览器复制
                email_body = "请将剪贴板内容粘贴到邮件正文中。"
                body = urllib.parse.quote(email_body)

                # 构建包含抄送的mailto URL
                mailto_url = f"mailto:{recipient_emails}?subject={subject_encoded}&body={body}&cc={cc_email}"

                # 打开邮件客户端
                os.startfile(mailto_url)

                # 等待邮件窗口就绪，然后尝试自动粘贴
                time.sleep(2)
                paste_success = False
                if clipboard_success:
                    paste_success = self.auto_paste_to_active_window()

                # 打开HTML文件以备份/预览
                os.startfile(html_file)

                print("已打开邮件客户端和HTML预览文件")
                if paste_success:
                    print("已尝试自动粘贴表格到当前邮件窗口，请检查发送窗口正文")
                elif clipboard_success:
                    print("HTML表格已复制到剪贴板，请在邮件正文中手动按Ctrl+V粘贴")
                print(f"收件人: {recipient_emails}")
                print(f"已添加抄送: {cc_email}")

        except Exception as e:
            print(f"发送邮件过程中出现错误: {e}")

    def run(self):
        """运行完整的爬取流程"""
        try:
            print("开始PKR数据爬取流程...")

            # 设置WebDriver
            self.setup_driver()

            # 登录系统
            self.login()

            # 导航到XSO Management并下载数据
            self.navigate_to_xso_management()
            self.download_xso_data()

            # 获取最新下载的文件
            latest_file = self.get_latest_downloaded_file()

            # 处理Excel数据
            df1 = self.process_excel_data(latest_file)

            # 提取EBP Leader信息
            df1 = self.extract_ebp_leader_info(df1)

            # 导航到PKR Management
            self.navigate_to_pkr_management()

            # 提取PKR信息
            df1 = self.extract_pkr_info(df1)

            # 导出最终数据
            output_file = self.export_final_data(df1)

            # 自动同步到飞书表格
            try:
                print("\n开始同步数据到飞书表格...")
                import sys
                import os
                # 确保导入路径正确
                current_dir = os.path.dirname(os.path.abspath(__file__))
                if current_dir not in sys.path:
                    sys.path.insert(0, current_dir)

                from PKR未完成情况同步到飞书 import sync_pkr_data_to_feishu
                sync_pkr_data_to_feishu()
                print("飞书表格同步完成！")
            except Exception as e:
                print(f"飞书同步失败: {e}")
                import traceback
                traceback.print_exc()
                print("请手动运行 'PKR未完成情况同步到飞书.py' 进行同步")

            # 发送Teams消息提醒
            self.send_teams_message(df1)

            print("PKR数据爬取流程完成！")
            print(f"结果已保存到: {output_file}")

            # 保持浏览器打开
            print("浏览器将保持打开状态，按回车键关闭...")
            input()

        except Exception as e:
            print(f"运行过程中出现错误: {e}")
            raise
        finally:
            if self.driver:
                self.driver.quit()


if __name__ == "__main__":
    crawler = PKRDataCrawler()
    crawler.run()