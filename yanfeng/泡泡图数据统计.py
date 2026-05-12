# -*- coding: utf-8 -*-
import ast
import glob
import importlib.util
import os
import re
import shutil
import sys
import time
from datetime import datetime
from difflib import SequenceMatcher
from config_loader import get_config_path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

config_path = get_config_path()

CONFIG_KEYS = {
    'PKR_LOGIN_URL': 'https://plmcnprdawc.yanfeng.com:3000/#/showHome',
    'PKR_USERNAME': 'uxuxs004',
    'PKR_PASSWORD': 'ABCabc%123',
    'EBP_LOGIN_URL': 'https://ebp.yanfeng.com/login',
    'EBP_USERNAME': 'uxuxs004',
    'EBP_PASSWORD': '2045O2O5-7'
}


def parse_config_py(path: str) -> dict:
    cfg = {}
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            match = re.match(r'^([A-Z0-9_]+)\s*=\s*(.+)$', line)
            if not match:
                continue
            key, value_text = match.groups()
            try:
                value = ast.literal_eval(value_text.strip())
            except Exception:
                value = value_text.strip().strip('"').strip("'")
            cfg[key] = value
    return cfg


def append_config_py_keys(path: str, defaults: dict):
    existing = parse_config_py(path)
    lines = []
    with open(path, encoding='utf-8') as f:
        lines = f.readlines()
    appended = False
    with open(path, 'a', encoding='utf-8') as f:
        for key, default in defaults.items():
            if key in existing:
                continue
            if key == 'PKR_LOGIN_URL' and 'LOGIN_URL' in existing:
                default = existing['LOGIN_URL']
            if key == 'PKR_USERNAME' and 'USERNAME' in existing:
                default = existing['USERNAME']
            if key == 'PKR_PASSWORD' and 'PASSWORD' in existing:
                default = existing['PASSWORD']
            f.write(f"\n{key} = {repr(default)}\n")
            appended = True
    return appended

append_config_py_keys(config_path, CONFIG_KEYS)

config_spec = importlib.util.spec_from_file_location('yanfeng_config', config_path)
config_module = importlib.util.module_from_spec(config_spec)
config_spec.loader.exec_module(config_module)
DOWNLOAD_DIR = config_module.DOWNLOAD_DIR

PKR_LOGIN_URL = getattr(config_module, 'PKR_LOGIN_URL', getattr(config_module, 'LOGIN_URL', CONFIG_KEYS['PKR_LOGIN_URL']))
PKR_USERNAME = getattr(config_module, 'PKR_USERNAME', getattr(config_module, 'USERNAME', CONFIG_KEYS['PKR_USERNAME']))
PKR_PASSWORD = getattr(config_module, 'PKR_PASSWORD', getattr(config_module, 'PASSWORD', CONFIG_KEYS['PKR_PASSWORD']))
EBP_LOGIN_URL = getattr(config_module, 'EBP_LOGIN_URL', CONFIG_KEYS['EBP_LOGIN_URL'])
EBP_USERNAME = getattr(config_module, 'EBP_USERNAME', CONFIG_KEYS['EBP_USERNAME'])
EBP_PASSWORD = getattr(config_module, 'EBP_PASSWORD', CONFIG_KEYS['EBP_PASSWORD'])

pkr_path = os.path.join(script_dir, 'PKR数据爬取.py')
if not os.path.exists(pkr_path):
    raise FileNotFoundError(f'未找到PKR数据爬取脚本: {pkr_path}')
pkr_spec = importlib.util.spec_from_file_location('pkr_data_crawler', pkr_path)
pkr_module = importlib.util.module_from_spec(pkr_spec)
pkr_spec.loader.exec_module(pkr_module)

for key, value in {
    'LOGIN_URL': PKR_LOGIN_URL,
    'USERNAME': PKR_USERNAME,
    'PASSWORD': PKR_PASSWORD
}.items():
    if hasattr(pkr_module, key):
        setattr(pkr_module, key, value)

PKRDataCrawler = pkr_module.PKRDataCrawler


def parse_month_input(text: str) -> datetime:
    text = str(text).strip()
    if len(text) == 4 and text.isdigit():
        year = 2000 + int(text[:2])
        month = int(text[2:])
    elif len(text) == 6 and text.isdigit():
        year = int(text[:4])
        month = int(text[4:])
    else:
        raise ValueError(f"无法识别的月份格式: {text}")
    if month < 1 or month > 12:
        raise ValueError(f"无法识别的月份: {month}")
    return datetime(year, month, 1)


def get_latest_project_list_file(download_dir: str) -> str:
    pattern = os.path.join(download_dir, "Project_List*.xls*")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(f"未找到Project_List文件，搜索路径: {pattern}")
    latest_file = max(files, key=os.path.getmtime)
    return latest_file


def to_datetime_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors='coerce')


def filter_phase3_by_month(df: pd.DataFrame, month_dt: datetime) -> pd.DataFrame:
    if 'Phase 3 Gate Exit-FPR' not in df.columns:
        return pd.DataFrame(columns=df.columns)
    df = df.copy()
    df['Phase 3 Gate Exit-FPR'] = to_datetime_series(df['Phase 3 Gate Exit-FPR'])

    # 筛选日期在指定月份的数据
    date_mask = (df['Phase 3 Gate Exit-FPR'].dt.year == month_dt.year) & (
        df['Phase 3 Gate Exit-FPR'].dt.month == month_dt.month)

    # 筛选Category列为"S"的数据
    category_mask = df['Category'] == 'S'

    # 同时满足两个条件
    combined_mask = date_mask & category_mask

    return df.loc[combined_mask].copy()


def normalize_id(id_value):
    if pd.isna(id_value):
        return ''
    return str(id_value).strip()


def merge_phase3_into_ppt(dfPPT: pd.DataFrame, dfPhase3: pd.DataFrame) -> pd.DataFrame:
    dfPPT = dfPPT.copy()
    dfPhase3 = dfPhase3.copy()
    dfPPT['Project ID'] = dfPPT['Project ID'].astype(str).str.strip()
    dfPhase3['Project ID'] = dfPhase3['Project ID'].astype(str).str.strip()

    ids_phase3 = set(dfPhase3['Project ID'].dropna())
    ids_ppt = set(dfPPT['Project ID'].dropna())

    dfPPT = dfPPT[dfPPT['Project ID'].isin(ids_phase3)].copy()

    missing_ids = ids_phase3 - ids_ppt
    if missing_ids:
        append_rows = []
        for pid in missing_ids:
            row = dfPhase3[dfPhase3['Project ID'] == pid]
            if row.empty:
                continue
            row = row.iloc[0]
            new_row = {col: None for col in dfPPT.columns}
            for col in ['Project ID', 'Project Name', 'EBP Leader', 'Product Engineer', 'Phase 3 Gate Exit-FPR']:
                if col in row.index:
                    new_row[col] = row[col]
            append_rows.append(new_row)
        if append_rows:
            dfPPT = pd.concat([dfPPT, pd.DataFrame(append_rows)], ignore_index=True, sort=False)
    return dfPPT


def update_from_history(dfPPT: pd.DataFrame, dfLS: pd.DataFrame) -> pd.DataFrame:
    dfPPT = dfPPT.copy()
    dfLS = dfLS.copy()
    if 'Project ID' not in dfPPT.columns or 'Project ID' not in dfLS.columns:
        return dfPPT
    dfPPT['Project ID'] = dfPPT['Project ID'].astype(str).str.strip()
    dfLS['Project ID'] = dfLS['Project ID'].astype(str).str.strip()
    common = set(dfPPT['Project ID']) & set(dfLS['Project ID'])
    if not common:
        return dfPPT
    dfLS_indexed = dfLS.set_index('Project ID')
    cols = [col for col in ['Design_release', 'Project Name Part', 'score', 'Formal EQU', 'Detailed_ver_crosstab_ID'] if col in dfLS_indexed.columns]
    for pid in common:
        values = dfLS_indexed.loc[pid, cols]
        if isinstance(values, pd.DataFrame):
            values = values.iloc[0]
        if isinstance(values, pd.Series):
            for col in cols:
                if col not in dfPPT.columns:
                    continue
                mask = (dfPPT['Project ID'] == pid)
                if col in dfPPT.columns:
                    target_series = dfPPT.loc[mask, col]
                    empty_mask = target_series.isna() | (target_series.astype(str).str.strip() == '')
                    if not empty_mask.any():
                        continue
                    dfPPT.loc[mask & empty_mask, col] = values[col]
    return dfPPT


def extract_between_commas(text: str, start_idx: int, end_idx: int) -> str:
    if not isinstance(text, str):
        return ''
    parts = [part.strip() for part in text.split(',')]
    if start_idx < 0 or end_idx > len(parts) or start_idx >= end_idx:
        return ''
    return ','.join(parts[start_idx:end_idx]).strip()


def fill_project_name_part(dfPPT: pd.DataFrame) -> pd.DataFrame:
    df = dfPPT.copy()
    if 'Project Name Part' not in df.columns or 'Project Name' not in df.columns:
        return df
    for idx, row in df.loc[df['Project Name Part'].isna() | (df['Project Name Part'].astype(str).str.strip() == ''), :].iterrows():
        project_name = row.get('Project Name', '')
        if not isinstance(project_name, str) or ',' not in project_name:
            continue
        parts = [p.strip() for p in project_name.split(',')]
        if len(parts) < 4:
            continue
        df.loc[idx, 'Project Name Part'] = ','.join(parts[2:-3]).strip() if len(parts) > 5 else ''
    return df


def fill_missing_score(dfPPT: pd.DataFrame) -> pd.DataFrame:
    df = dfPPT.copy()
    score_columns = [col for col in df.columns if col.strip().lower() == 'score']
    for col in score_columns:
        df[col] = df[col].fillna(100)
    return df


def compute_weeks_after_gate(dfPPT: pd.DataFrame) -> pd.DataFrame:
    df = dfPPT.copy()
    if 'Phase 3 Gate Exit-FPR' in df.columns and 'Design_release' in df.columns:
        df['Phase 3 Gate Exit-FPR'] = to_datetime_series(df['Phase 3 Gate Exit-FPR'])
        df['Design_release'] = to_datetime_series(df['Design_release'])
        def calc(row):
            a = row['Phase 3 Gate Exit-FPR']
            b = row['Design_release']
            if pd.isna(a) or pd.isna(b):
                return None
            weeks = (b - a).days / 7.0
            return round(weeks, 2)
        df['Max weeks after gate'] = df.apply(calc, axis=1)
    return df


def find_local_edge_driver() -> str | None:
    env_path = os.environ.get('EDGE_DRIVER_PATH')
    if env_path and os.path.exists(env_path):
        return env_path
    candidates = [
        r'C:\Windows\System32\msedgedriver.exe',
        r'C:\Program Files\msedgedriver.exe',
        r'C:\Program Files (x86)\msedgedriver.exe',
        r'C:\Tools\msedgedriver.exe',
        r'C:\drivers\msedgedriver.exe',
        r'.\msedgedriver.exe'
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    cache_dirs = [
        os.path.join(os.path.expanduser('~'), '.cache', 'selenium'),
        os.path.join(os.path.expanduser('~'), '.cache', 'webdriver'),
        os.path.join(os.path.expanduser('~'), '.cache', 'msedgedriver')
    ]
    for cache_dir in cache_dirs:
        pattern = os.path.join(cache_dir, '**', 'msedgedriver.exe')
        matches = glob.glob(pattern, recursive=True)
        if matches:
            return max(matches, key=os.path.getmtime)
    which_path = shutil.which('msedgedriver') or shutil.which('msedgedriver.exe')
    if which_path:
        return which_path
    return None


def build_edge_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    driver_path = None
    try:
        driver_path = ChromeDriverManager().install()
    except Exception as ex:
        raise RuntimeError(
            f'无法下载ChromeDriver。原始错误: {ex}'
        ) from ex

    service = webdriver.chrome.service.Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver


def wait_for_text_element(driver, text, timeout=20):
    wait = WebDriverWait(driver, timeout)
    xpath = f"//*[normalize-space(text())='{text}']"
    return wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))


def find_closest_text_match(source: str, candidates: pd.Series) -> pd.DataFrame:
    source = str(source or '').strip()
    scores = candidates.fillna('').astype(str).apply(lambda x: SequenceMatcher(None, source, x.strip()).ratio())
    return pd.DataFrame({'text': candidates, 'score': scores})


def parse_oem_from_name_part(text: str) -> str:
    if not isinstance(text, str):
        return ''
    match = re.match(r'\s*([^,]+)', text)
    return match.group(1).strip() if match else text.strip()


def parse_model_and_device(text: str):
    if not isinstance(text, str):
        return '', ''
    parts = [p.strip() for p in text.split(',')]
    model = parts[1] if len(parts) > 1 else ''
    device = parts[2] if len(parts) > 2 else ''
    return model, device


def fetch_formal_equ_from_ebp(dfPPT: pd.DataFrame) -> pd.DataFrame:
    driver = None
    try:
        driver = build_edge_driver()
        driver.get(EBP_LOGIN_URL)
        wait = WebDriverWait(driver, 20)

        # 登录
        account_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="account"]')))
        account_input.clear()
        account_input.send_keys(EBP_USERNAME)
        password_input = driver.find_element(By.CSS_SELECTOR, 'input[formcontrolname="password"]')
        password_input.clear()
        password_input.send_keys(EBP_PASSWORD)
        password_input.send_keys(Keys.RETURN)

        # 等待登录成功
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//span[text()="EBP Program"]'))
        )
        print("登录成功")

        # 导航到My Program List
        ebp_program_span = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//span[text()="EBP Program"]'))
        )
        ebp_program_span.click()

        my_program_list_span = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//span[text()="My Program List"]'))
        )
        my_program_list_span.click()
        print("已进入My Program List页面")
        time.sleep(3)

        # 处理每个缺失EQU的项目
        for idx, row in dfPPT.loc[dfPPT['Formal EQU'].isna() | (dfPPT['Formal EQU'].astype(str).str.strip() == ''), :].iterrows():
            project_name_part = str(row.get('Project Name Part', ''))
            if not project_name_part.strip():
                continue

            print(f"处理第 {idx+1} 行数据: {project_name_part[:50]}...")

            try:
                # 确保在My Program List页面
                my_program_list_span = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//span[text()="My Program List"]'))
                )
                my_program_list_span.click()
                time.sleep(2)

                # 提取OEM
                if "," in project_name_part:
                    oem = re.split(r',', project_name_part)[0].strip()
                    print(f"提取的OEM: {oem}")
                else:
                    oem = project_name_part.strip()
                    print(f"未找到分隔符，使用完整名称作为OEM: {oem}")

# ===== 过滤器输入逻辑（修复版）=====
                try:
                    # 1. 精确定位过滤容器
                    filter_container = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[col-id="display_name_filter"]'))
                    )

                    # 2. JavaScript安全点击过滤按钮
                    filter_btn = filter_container.find_element(By.CSS_SELECTOR, 'span[data-ref="eFilterButton"]')
                    driver.execute_script("arguments[0].click();", filter_btn)
                    print("✓ 已点击过滤按钮")
                    time.sleep(1.5)

                    # 3. 使用成功的策略定位输入框
                    input_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "div.ag-popup input[type='text']")))

                    # 4. JavaScript强制设置值 + 触发事件
                    driver.execute_script("""
                        arguments[0].value = '';
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """, input_field, oem)
                    print(f"✓ 已通过JS设置过滤值: {oem}")

                    # 5. 显式触发搜索
                    input_field.send_keys(Keys.RETURN)
                    time.sleep(2.5)

                    # 6. 验证输入生效
                    current_val = driver.execute_script("return arguments[0].value;", input_field)
                    if current_val != oem:
                        print(f"⚠ 警告：输入框实际值 '{current_val}' 与目标 '{oem}' 不符")

                except Exception as e:
                    print(f"✗ 过滤操作失败 (OEM={oem}): {str(e)}")
                    continue

                # ===== 提取项目列表 =====
                df_OEMprojects = pd.DataFrame(columns=["项目全名", "车型代号零部件名", "加工设备"])

                try:
                    project_list_container = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ag-body-viewport.ag-layout-normal.ag-row-animation'))
                    )

                    # 提取所有可见span
                    project_spans = project_list_container.find_elements(By.XPATH, './/span[not(contains(@class, "ag-hidden")) and string-length(text()) > 5]')

                    for span in project_spans:
                        full_name = span.text.strip()
                        if full_name.count(',') >= 4:
                            parts = [p.strip() for p in full_name.split(',')]
                            vehicle_part = parts[3] if len(parts) > 3 else ""
                            equipment = parts[4] if len(parts) > 4 else ""

                            new_row = pd.DataFrame([{
                                "项目全名": full_name,
                                "车型代号零部件名": vehicle_part,
                                "加工设备": equipment
                            }])
                            df_OEMprojects = pd.concat([df_OEMprojects, new_row], ignore_index=True)

                    print(f"✓ 有效项目数: {len(df_OEMprojects)}")
                except Exception as e:
                    print(f"✗ 项目列表提取失败: {str(e)}")
                    continue

                if df_OEMprojects.empty:
                    print("⚠ 无有效项目数据，跳过当前行")
                    continue

                # ===== 双层相似度匹配 =====
                current_pnp = str(row["Project Name Part"]).strip()
                pnp_parts = [p.strip() for p in current_pnp.split(',') if p.strip()]
                target_vehicle = pnp_parts[1] if len(pnp_parts) > 1 else ""
                target_equipment = pnp_parts[2] if len(pnp_parts) > 2 else ""

                def calc_sim(s1, s2):
                    if not s1 or not s2:
                        return 0.0
                    clean1 = re.sub(r'\s+', ' ', s1.strip().lower())
                    clean2 = re.sub(r'\s+', ' ', s2.strip().lower())
                    return SequenceMatcher(None, clean1, clean2).ratio()

                df_OEMprojects["车型相似度"] = df_OEMprojects["车型代号零部件名"].apply(
                    lambda x: calc_sim(x, target_vehicle))
                df_OEMprojects["设备相似度"] = df_OEMprojects["加工设备"].apply(
                    lambda x: calc_sim(x, target_equipment) if target_equipment else 0.0)

                df_OEMprojects = df_OEMprojects.sort_values(
                    ["车型相似度", "设备相似度"],
                    ascending=[False, False]
                ).reset_index(drop=True)

                ProjectName = df_OEMprojects.iloc[0]["项目全名"]
                top_vehicle_sim = df_OEMprojects.iloc[0]["车型相似度"]
                top_equipment_sim = df_OEMprojects.iloc[0]["设备相似度"]

                print(f"🔍 匹配结果 | 车型: '{target_vehicle}'(目标) vs '{df_OEMprojects.iloc[0]['车型代号零部件名']}'(匹配) | "
                    f"车型相似度: {top_vehicle_sim:.2%} | 设备相似度: {top_equipment_sim:.2%}")

                # ===== 点击匹配的项目 =====
                try:
                    safe_name = ProjectName.replace("'", "\\'").replace("\n", " ").replace("\r", " ").strip()

                    click_script = f"""
                    const targetText = '{safe_name}';
                    const spans = document.querySelectorAll('span');
                    for (let span of spans) {{
                        if (span.textContent.trim() === targetText &&
                            window.getComputedStyle(span).visibility !== 'hidden' &&
                            !span.closest('.ag-hidden')) {{
                            span.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                            setTimeout(() => {{
                                if (span.offsetParent !== null) {{
                                    span.click();
                                }}
                            }}, 400);
                            return true;
                        }}
                    }}
                    return false;
                    """

                    clicked = driver.execute_script(click_script)
                    if not clicked:
                        partial_script = f"""
                        const targetText = '{safe_name}';
                        const spans = document.querySelectorAll('span');
                        for (let span of spans) {{
                            if (span.textContent.includes('{safe_name.split(',')[0]}') &&
                                span.textContent.includes('{safe_name.split(',')[1] or ''}') &&
                                window.getComputedStyle(span).visibility !== 'hidden') {{
                                span.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                                setTimeout(() => {{ span.click(); }}, 400);
                                return true;
                            }}
                        }}
                        return false;
                        """
                        clicked = driver.execute_script(partial_script)
                        if not clicked:
                            raise Exception("JS点击失败：未找到匹配元素")

                    print(f"✓ 已通过JS成功点击项目: {ProjectName[:50]}...")
                    time.sleep(3.5)

                except Exception as e:
                    print(f"✗ 点击项目失败: {str(e)}")
                    continue

                # ===== 提取EQU =====
                try:
                    budget_tab = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "Budget")]'))
                    )
                    budget_tab.click()
                    print("已点击 Budget 标签")
                    time.sleep(3)

                    equ_label = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "EQU Total:")]'))
                    )

                    parent_div = equ_label.find_element(By.XPATH, './ancestor::div[1]')
                    equ_value_div = parent_div.find_element(By.CSS_SELECTOR, 'div.flex-1')
                    equ_value = equ_value_div.text.strip()

                    print(f"找到的EQU Total值: {equ_value}")

                    try:
                        equ_value_float = float(equ_value)
                        dfPPT.at[idx, "Formal EQU"] = equ_value_float
                        print(f"成功保存EQU值: {equ_value_float}")
                    except ValueError:
                        print(f"警告: 无法将 '{equ_value}' 转换为数字，保存为字符串")
                        dfPPT.at[idx, "Formal EQU"] = equ_value

                    dfPPT.at[idx, "Detailed_ver_crosstab_name"] = ProjectName

                except Exception as e:
                    print(f"✗ 提取EQU Total值失败: {str(e)}")

                # 返回项目列表页面
                try:
                    my_program_list_span = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//span[text()="My Program List"]'))
                    )
                    my_program_list_span.click()
                    time.sleep(2)
                except Exception as e:
                    print(f"返回My Program List失败: {str(e)}")
                    continue

            except Exception as e:
                print(f"处理行 {idx} 时发生错误: {e}")
                continue

        return dfPPT
    finally:
        if driver:
            driver.quit()


def refresh_history(dfPPT: pd.DataFrame, dfLS: pd.DataFrame) -> pd.DataFrame:
    dfPPT = dfPPT.copy()
    if 'Project ID' not in dfPPT.columns or 'Project ID' not in dfLS.columns:
        return dfLS
    dfPPT['Project ID'] = dfPPT['Project ID'].astype(str).str.strip()
    dfLS['Project ID'] = dfLS['Project ID'].astype(str).str.strip()
    combined = pd.concat([dfLS[~dfLS['Project ID'].isin(dfPPT['Project ID'])], dfPPT], ignore_index=True, sort=False)
    combined = combined.drop_duplicates(subset=['Project ID'], keep='last')
    return combined


def save_dataframe(df: pd.DataFrame, path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            df.to_excel(path, index=False)
            print(f"成功保存文件: {path}")
            return
        except PermissionError as e:
            if attempt < max_retries - 1:
                print(f"文件被占用，{attempt + 1}秒后重试... (错误: {e})")
                time.sleep(1)
            else:
                # 如果重试失败，保存到临时文件
                import tempfile
                temp_dir = tempfile.gettempdir()
                temp_filename = f"泡泡图当月数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                temp_path = os.path.join(temp_dir, temp_filename)
                df.to_excel(temp_path, index=False)
                print(f"文件被占用，已保存到临时文件: {temp_path}")
                print("请手动将临时文件复制到目标位置并关闭Excel程序。")


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    month_input = input('请输入要统计的月份，例如2605：').strip()
    Month = parse_month_input(month_input)
    print(f'统计月份：{Month.strftime("%Y-%m")}')

    crawler = PKRDataCrawler()
    try:
        crawler.setup_driver()
        crawler.login()
        crawler.navigate_to_xso_management()
        crawler.download_xso_data()
    finally:
        if crawler.driver:
            crawler.driver.quit()

    project_list_file = get_latest_project_list_file(DOWNLOAD_DIR)
    df_project_list = pd.read_excel(project_list_file)
    dfPhase3 = filter_phase3_by_month(df_project_list, Month)

    dfPPT_path = os.path.join(script_dir, '泡泡图DATA', '泡泡图当月数据.xlsx')
    dfLS_path = os.path.join(script_dir, '泡泡图DATA', '泡泡图历史数据.xlsx')
    dfPPT = pd.read_excel(dfPPT_path)
    dfLS = pd.read_excel(dfLS_path)

    dfPPT = merge_phase3_into_ppt(dfPPT, dfPhase3)
    dfPPT = update_from_history(dfPPT, dfLS)
    dfPPT = fill_project_name_part(dfPPT)
    dfPPT = fill_missing_score(dfPPT)
    dfPPT = compute_weeks_after_gate(dfPPT)

    if dfPPT['Formal EQU'].isna().any() or (dfPPT['Formal EQU'].astype(str).str.strip() == '').any():
        dfPPT = fetch_formal_equ_from_ebp(dfPPT)

    dfLS_updated = refresh_history(dfPPT, dfLS)
    save_dataframe(dfLS_updated, dfLS_path)
    save_dataframe(dfPPT, dfPPT_path)
    print('处理完成，已覆盖泡泡图历史数据.xlsx和泡泡图当月数据.xlsx')


if __name__ == '__main__':
    main()
