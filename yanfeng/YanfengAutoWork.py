#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Yanfeng Auto Work 工具库
包含常用的自动化操作函数
"""

import re  # 需要在文件开头导入，或者函数内导入
import os
import time
import glob
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def PLMLogin(driver, wait):
    """
    PLM系统登录函数

    Args:
        driver: Selenium WebDriver实例
        wait: WebDriverWait实例

    Returns:
        bool: 登录成功返回True，失败抛出异常
    """
    try:
        # 从配置中获取登录信息
        from config_loader import load_config_module
        config = load_config_module()
        LOGIN_URL = config.LOGIN_URL
        USERNAME = config.USERNAME
        PASSWORD = config.PASSWORD
        PAGE_LOAD_WAIT_TIME = getattr(config, 'PAGE_LOAD_WAIT_TIME', 3)

        print(f"正在打开登录页面: {LOGIN_URL}")
        driver.get(LOGIN_URL)

        # 等待页面加载
        time.sleep(PAGE_LOAD_WAIT_TIME)

        # 检查页面标题
        print(f"当前页面标题: {driver.title}")

        # 输入用户名
        print("正在定位用户名输入框...")
        username_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']"))
        )
        username_input.clear()
        username_input.send_keys(USERNAME)
        print("已输入用户名")

        # 输入密码
        print("正在定位密码输入框...")
        password_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Password']")
        password_input.clear()
        password_input.send_keys(PASSWORD)
        print("已输入密码")

        # 点击登录按钮
        print("正在定位登录按钮...")
        login_button = driver.find_element(By.CSS_SELECTOR, "button.sw-button.accent-caution")
        login_button.click()
        print("已点击登录按钮")

        # 等待页面跳转
        time.sleep(PAGE_LOAD_WAIT_TIME * 2)

        # 等待页面跳转完成
        time.sleep(PAGE_LOAD_WAIT_TIME * 2)
        print("登录流程完成")
        return True

    except Exception as e:
        print(f"登录过程中出现错误: {e}")
        print(f"当前URL: {driver.current_url}")
        print(f"当前标题: {driver.title}")
        raise


def DL_Project_List(driver, wait, download_dir=None):
    """
    下载Project_List数据（带4小时缓存检查）

    Args:
        driver: Selenium WebDriver实例
        wait: WebDriverWait实例
        download_dir: 下载目录，默认为None（使用配置或~/Downloads）

    Returns:
        str: Project_List文件的完整路径（来自缓存或新下载）
    """
    try:
        from config_loader import load_config_module
        config = load_config_module()

        if download_dir is None:
            download_dir = getattr(config, 'DOWNLOAD_DIR', os.path.expanduser('~/Downloads'))

        # ---------- 新增：检查本地缓存文件 ----------
        # 匹配 Project_List 相关文件（xls 或 xlsx），排除临时下载文件
        pattern = os.path.join(download_dir, "Project_List*.xls*")
        all_files = [f for f in glob.glob(pattern) 
                     if not f.lower().endswith('.crdownload') 
                     and not f.lower().endswith('.part')]
        
        if all_files:
            # 按最后修改时间排序，获取最新文件
            latest_file = max(all_files, key=os.path.getmtime)
            file_mtime = os.path.getmtime(latest_file)  # 秒为单位
            now = time.time()
            age_seconds = now - file_mtime
            age_hours = age_seconds / 3600.0

            if age_seconds <= 4 * 3600:   # 4小时以内
                print(f"找到可用的缓存文件: {latest_file}")
                print(f"文件修改时间距今 {age_hours:.2f} 小时（≤4小时），将直接使用缓存，不再重新下载。")
                return latest_file
            else:
                print(f"找到缓存文件 {latest_file}，但距今 {age_hours:.2f} 小时（>4小时），将重新下载最新数据。")
        else:
            print("未找到任何 Project_List 缓存文件，将执行下载。")
        # ---------- 缓存检查结束 ----------

        # 以下为原有下载逻辑
        print("正在导航到XSO Management...")

        # 点击XSO & PKR
        xso_pkr_div = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='XSO & PKR']"))
        )
        xso_pkr_div.click()
        print("已点击XSO & PKR")

        # ... 后续代码完全保持原样 ...
        # 定位iframeContent
        iframe_content = wait.until(
            EC.presence_of_element_located((By.ID, "iframeContent"))
        )
        # 等待iframe内容加载
        time.sleep(2)
        # 切换到iframe
        iframe_content = driver.find_element(By.ID, "iframeContent")
        driver.switch_to.frame(iframe_content)
        print("已切换到iframeContent")

        # 展开XSO Management - 增加重试机制
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # 尝试多种定位方式
                try:
                    xso_management = wait.until(
                        EC.element_to_be_clickable((By.LINK_TEXT, "XSO Management"))
                    )
                except:
                    # 如果LINK_TEXT失败，尝试XPath
                    xso_management = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//span[text()='XSO Management']/parent::a"))
                    )

                xso_management.click()
                print("已展开XSO Management")
                break
            except Exception as e:
                print(f"尝试 {attempt + 1} 次展开XSO Management失败: {e}")
                if attempt == max_retries - 1:
                    raise
                time.sleep(2)

        # 点击Program Dashboard
        program_dashboard = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Program Dashboard"))
        )
        program_dashboard.click()
        print("已点击Program Dashboard")

        # 等待iframe加载
        time.sleep(6)  # 增加等待时间

        # 切换到XSO iframe
        iframe = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-1-1"))
        )
        driver.switch_to.frame(iframe)
        print("已切换到XSO iframe")

        # 选择BU
        print("正在选择BU...")
        bu_select = Select(driver.find_element(By.NAME, "select_bu"))
        bu_select.select_by_value(getattr(config, 'BU_VALUE', 'DFM'))
        print(f"已选择BU: {getattr(config, 'BU_VALUE', 'DFM')}")

        # 选择Category
        print("正在选择Category...")
        category_select = Select(driver.find_element(By.NAME, "select_Category"))
        category_select.select_by_value(getattr(config, 'CATEGORY_VALUE', 'S'))
        print(f"已选择Category: {getattr(config, 'CATEGORY_VALUE', 'S')}")

        # 输入Product Group
        print("正在输入Product Group...")
        product_group_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Product Group']")
        product_group_input.clear()
        product_group_input.send_keys(getattr(config, 'PRODUCT_GROUP', 'equipment'))
        print(f"已输入Product Group: {getattr(config, 'PRODUCT_GROUP', 'equipment')}")

        # 点击搜索按钮
        search_button = driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']")
        search_button.click()
        print("已点击搜索按钮")

        # 等待搜索结果加载
        time.sleep(getattr(config, 'PAGE_LOAD_WAIT_TIME', 3))

        # 记录当前已有Project_List文件，避免误判旧文件
        existing_files = set(glob.glob(os.path.join(download_dir, "Project_List*.xls*")))

        # 点击导出按钮（实际上是span元素）
        export_button = driver.find_element(By.CSS_SELECTOR, "span[onclick='exportData()']")
        export_button.click()
        print("已点击导出按钮")

        # 等待下载文件出现并完成
        download_file = wait_for_project_list_download(download_dir, existing_files, timeout=180)
        print(f"下载完成: {download_file}")

        # 切换回默认内容
        driver.switch_to.default_content()

        return download_file

    except Exception as e:
        print(f"下载XSO数据过程中出现错误: {e}")
        raise


def wait_for_project_list_download(download_dir, existing_files=None, timeout=120, poll_interval=2):
    """
    等待最新Project_List下载完成并返回文件路径

    Args:
        download_dir: 下载目录
        existing_files: 已存在的文件集合
        timeout: 超时时间（秒）
        poll_interval: 检查间隔（秒）

    Returns:
        str: 下载完成的文件路径
    """
    if existing_files is None:
        existing_files = set()

    pattern = os.path.join(download_dir, "Project_List*.xls*")

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # 获取所有匹配的文件（排除临时下载文件）
            all_files = [f for f in glob.glob(pattern) if not f.lower().endswith('.crdownload') and not f.lower().endswith('.part')]

            # 过滤掉已存在的文件
            new_files = [f for f in all_files if f not in existing_files]

            if new_files:
                # 按修改时间排序，获取最新的文件
                latest_file = max(new_files, key=os.path.getctime)
                print(f"找到新下载文件: {latest_file}")

                # 等待文件完全写入
                initial_size = os.path.getsize(latest_file)
                time.sleep(2)
                final_size = os.path.getsize(latest_file)

                if initial_size == final_size and initial_size > 0:
                    return latest_file

        except Exception as e:
            print(f"检查下载文件时出错: {e}")

        time.sleep(poll_interval)

    raise TimeoutError(f"等待Project_List下载完成超时: {pattern}")


def RD_Project_List(file_path):
    """
    读取Project_List Excel文件并返回处理后的DataFrame

    Args:
        file_path: Excel文件路径

    Returns:
        pandas.DataFrame: 处理后的数据框（只筛选Category=S，不进行日期筛选）
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        print(f"已读取Excel文件，共{len(df)}行数据")

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

        # 转换日期列为datetime格式，便于后续处理
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        print(f"数据预处理完成，共{len(df)}行Category=S的数据")
        return df

    except Exception as e:
        print(f"处理Excel数据过程中出现错误: {e}")
        raise


def setup_driver(download_dir=None):
    """
    设置Edge浏览器WebDriver

    Args:
        download_dir: 下载目录

    Returns:
        tuple: (driver, wait) WebDriver和WebDriverWait实例
    """
    from config_loader import load_config_module
    config = load_config_module()

    if download_dir is None:
        download_dir = getattr(config, 'DOWNLOAD_DIR', os.path.expanduser('~/Downloads'))

    print("正在设置Edge浏览器...")

    edge_options = EdgeOptions()
    edge_options.add_argument('--ignore-certificate-errors')
    edge_options.add_argument('--ignore-ssl-errors')
    edge_options.add_argument('--start-maximized')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')

    # 设置下载选项
    edge_options.add_experimental_option('prefs', {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True
    })

    try:
        # 尝试自动下载EdgeDriver
        print("尝试使用EdgeDriverManager...")
        edge_service = EdgeService(EdgeChromiumDriverManager().install())
    except Exception as e:
        print(f"EdgeDriverManager下载失败: {e}")
        print("尝试使用本地EdgeDriver...")
        try:
            # 尝试常见EdgeDriver路径
            possible_paths = [
                'msedgedriver.exe',  # 当前目录
                r'C:\edgedriver\msedgedriver.exe',  # 常见安装路径
                r'C:\Program Files\edgedriver\msedgedriver.exe',
                r'C:\Windows\msedgedriver.exe'
            ]

            driver_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    driver_path = path
                    break

            if driver_path:
                print(f"找到本地EdgeDriver: {driver_path}")
                edge_service = EdgeService(driver_path)
            else:
                print("未找到本地EdgeDriver，让Selenium自动查找...")
                edge_service = EdgeService()

        except Exception as local_e:
            print(f"本地EdgeDriver也失败: {local_e}")
            raise Exception("无法初始化EdgeDriver，请确保已安装Edge浏览器")

    driver = webdriver.Edge(service=edge_service, options=edge_options)
    wait = WebDriverWait(driver, 10)
    print("成功初始化Edge浏览器")

    return driver, wait

def FindEBPLeader(driver, wait, project_id_list):
    """
    从本地缓存表或 PLM 系统中提取 EBP Leader 信息
    优先读取本地“项目信息表.xlsx”，未找到的再通过网页搜索获取
    匹配规则：将传入的 Project ID 中的数字部分与本地表的数字列进行比对
    """
    # ---------- 辅助函数：提取数字部分 ----------
    def extract_digits(pid):
        """提取字符串中的连续数字，若无则返回原字符串"""
        match = re.search(r'\d+', str(pid))
        return match.group(0) if match else str(pid).strip()

    # ---------- 1. 加载本地缓存表 ----------
    script_dir = os.path.dirname(os.path.abspath(__file__))
    local_excel_path = os.path.join(script_dir, "项目信息表", "项目信息表.xlsx")
    local_mapping = {}  # 数字部分 -> ebp_leader

    if os.path.exists(local_excel_path):
        try:
            local_df = pd.read_excel(local_excel_path)
            if 'Project ID' in local_df.columns and 'EBP Leader' in local_df.columns:
                for _, row in local_df.iterrows():
                    pid_raw = str(row['Project ID']).strip()
                    # 本地表的 Project ID 可能是数字或字符串，同样提取数字部分作为 key
                    pid_digits = extract_digits(pid_raw)
                    leader = str(row['EBP Leader']).strip() if pd.notna(row['EBP Leader']) else ''
                    if pid_digits:  # 确保有内容
                        local_mapping[pid_digits] = leader
                print(f"已加载本地缓存表，共 {len(local_mapping)} 条记录")
            else:
                print("警告：本地项目信息表缺少 'Project ID' 或 'EBP Leader' 列，将忽略本地缓存")
        except Exception as e:
            print(f"警告：读取本地项目信息表失败 - {e}，将完全通过网页查询")
    else:
        print(f"未找到本地缓存文件 {local_excel_path}，将完全通过网页查询")

    # ---------- 2. 分离已缓存和需查询的项目 ----------
    need_search = []
    results = []
    for pid in project_id_list:
        pid_str = str(pid).strip()
        pid_digits = extract_digits(pid_str)  # 提取数字部分用于匹配
        if pid_digits in local_mapping and local_mapping[pid_digits]:
            # 本地已有有效数据，直接使用
            results.append({'Project ID': pid, 'EBP Leader': local_mapping[pid_digits]})
            print(f"Project ID {pid} (数字部分 {pid_digits}) 从本地缓存获取 EBP Leader: {local_mapping[pid_digits]}")
        else:
            need_search.append(pid)
            results.append({'Project ID': pid, 'EBP Leader': ''})  # 占位

    if not need_search:
        print("所有 Project ID 均在本地缓存中找到，无需网页查询")
        return pd.DataFrame(results)

    print(f"需要网页查询的项目数量: {len(need_search)}")

    # ---------- 3. 网页查询未缓存的项目 ----------
    try:
        from config_loader import load_config_module
        config = load_config_module()
        PAGE_LOAD_WAIT_TIME = getattr(config, 'PAGE_LOAD_WAIT_TIME', 3)
        BU_VALUE = getattr(config, 'BU_VALUE', 'DFM')
        CATEGORY_VALUE = getattr(config, 'CATEGORY_VALUE', 'S')
        PRODUCT_GROUP = getattr(config, 'PRODUCT_GROUP', 'equipment')

        print("正在导航到 XSO Management 提取 EBP Leader 信息...")

        driver.switch_to.default_content()
        time.sleep(PAGE_LOAD_WAIT_TIME)

        # 检查是否已在 XSO Management 页面，若不在则先点击跳转
        try:
            driver.find_element(By.XPATH, '//span[@class="menu-text" and text()="XSO Management"]')
            print("当前已在 XSO Management 页面")
        except NoSuchElementException:
            print("当前不在 XSO Management 页面，尝试点击 XSO & PKR 跳转...")
            try:
                xso_pkr_div = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div[title="XSO & PKR"]')))
                xso_pkr_div.click()
                print("已点击 XSO & PKR")
                time.sleep(PAGE_LOAD_WAIT_TIME)
            except Exception as nav_e:
                print(f"点击 XSO & PKR 失败: {nav_e}")

        iframe_content = wait.until(EC.presence_of_element_located((By.ID, "iframeContent")))
        driver.switch_to.frame(iframe_content)
        print("已切换到 iframeContent")

        # 展开 XSO Management
        try:
            menu_item = wait.until(EC.presence_of_element_located((By.ID, "menu-1")))
            if "open" in menu_item.get_attribute("class").split():
                print("XSO Management 已经展开")
            else:
                xso_management = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='XSO Management']/parent::a")))
                xso_management.click()
                print("已展开 XSO Management")
        except Exception:
            xso_management = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='XSO Management']/parent::a")))
            xso_management.click()

        program_dashboard = wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='Program Dashboard']/parent::a")))
        driver.execute_script("arguments[0].click();", program_dashboard)
        print("已点击 Program Dashboard")

        time.sleep(PAGE_LOAD_WAIT_TIME * 2)

        xso_iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "iframe_tab_menu-1-1")))
        driver.switch_to.frame(xso_iframe)
        print("已切换到 XSO iframe")

        # 设置筛选条件
        Select(driver.find_element(By.NAME, "select_bu")).select_by_value(BU_VALUE)
        Select(driver.find_element(By.NAME, "select_Category")).select_by_value(CATEGORY_VALUE)
        driver.find_element(By.CSS_SELECTOR, "input[placeholder='Product Group']").send_keys(PRODUCT_GROUP)
        driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']").click()
        time.sleep(PAGE_LOAD_WAIT_TIME)

        # 逐个查询未缓存的项目
        success_count = 0
        failed_count = 0
        failed_projects = []

        for project_id in need_search:
            print(f"正在网页查询 Project ID: {project_id}")
            ebp_leader = ""
            try:
                search_input = wait.until(EC.presence_of_element_located((By.ID, "searchProject")))
                search_input.clear()
                search_input.send_keys(str(project_id))  # 保持原样发送（可改为数字部分，根据需求决定）

                driver.find_element(By.CSS_SELECTOR, "button[onclick='searchproject()']").click()
                wait.until(EC.presence_of_element_located((By.ID, "grid-table")))
                time.sleep(2)

                ebp_elements = driver.find_elements(By.CSS_SELECTOR, "td[aria-describedby='grid-table_real_pm_name']")
                if ebp_elements:
                    title = ebp_elements[0].get_attribute("title")
                    if title:
                        ebp_leader = title.strip()
                    else:
                        ebp_leader = ebp_elements[0].text.strip()
                    print(f"网页查询成功: {project_id} -> {ebp_leader}")
                    success_count += 1
                else:
                    print(f"未找到 EBP Leader: {project_id}")
                    failed_count += 1
                    failed_projects.append(project_id)

                # 更新 results
                for item in results:
                    if str(item['Project ID']) == str(project_id):
                        item['EBP Leader'] = ebp_leader
                        break

            except Exception as e:
                print(f"网页查询失败 {project_id}: {e}")
                failed_count += 1
                failed_projects.append(project_id)

        print(f"\n网页查询统计: 成功 {success_count}, 失败 {failed_count}")
        if failed_projects:
            print(f"失败列表: {failed_projects}")

    except Exception as e:
        print(f"网页查询过程出现错误: {e}")

    return pd.DataFrame(results)


if __name__ == "__main__":
    # 测试代码
    print("YanfengAutoWork模块测试")

    try:
        # 测试浏览器设置
        driver, wait = setup_driver()
        print("浏览器设置成功")

        # 测试登录
        PLMLogin(driver, wait)
        print("登录测试成功")

        # 测试下载
        download_file = DL_Project_List(driver, wait)
        print(f"下载测试成功: {download_file}")

        # 测试读取
        df = RD_Project_List(download_file)
        print(f"读取测试成功，数据形状: {df.shape}")

        # 关闭浏览器
        driver.quit()
        print("所有测试完成")

    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc()