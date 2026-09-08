#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
飞书项目信息表更新脚本
功能：从PLM系统获取项目数据，与飞书表格数据对比并更新
"""

import os
import sys
import re
import pandas as pd
from datetime import datetime, timedelta
import requests
import json
import platform
import subprocess
import traceback

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 从 YanfengAutoWork 导入所需的函数
from YanfengAutoWork import setup_driver, PLMLogin, DL_Project_List, RD_Project_List, FindEBPLeader
from config_loader import load_config_module


def get_feishu_access_token(app_id, app_secret):
    """获取飞书访问令牌"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    payload = {"app_id": app_id, "app_secret": app_secret}

    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        return response.json().get("tenant_access_token")
    else:
        raise Exception(f"获取飞书访问令牌失败: {response.text}")


def get_feishu_table_data(app_token, table_id, view_id, access_token):
    """获取飞书表格数据（返回所有字段）"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    all_records = []
    page_token = None

    while True:
        params = {"view_id": view_id}
        if page_token:
            params["page_token"] = page_token

        response = requests.get(url, headers=headers, params=params)
        data = response.json()
        if response.status_code == 200 and data.get("code") == 0:
            records = data.get("data", {}).get("items", [])
            all_records.extend(records)
            if data.get("data", {}).get("has_more"):
                page_token = data.get("data", {}).get("page_token")
            else:
                break
        else:
            raise Exception(f"获取飞书表格数据失败: {response.status_code} {response.text}")

    return all_records


def convert_timestamp_to_date(value):
    """将毫秒时间戳转换为 YYYY-MM-DD 字符串，非时间戳返回原值"""
    if isinstance(value, (int, float)):
        # 简单判断：13位毫秒时间戳范围（约 2000-2100 年）
        if 1e12 < value < 1e13:
            try:
                dt = datetime.fromtimestamp(value / 1000)
                return dt.strftime('%Y-%m-%d')
            except (ValueError, OSError):
                pass
    return value


def convert_to_timestamp(date_str):
    """将多种格式的日期字符串转换为毫秒时间戳"""
    if not date_str or date_str is None:
        return None
    date_str = str(date_str).strip()
    if not date_str:
        return None

    # 清理常见全角字符
    date_str = date_str.replace('／', '/').replace('—', '-')
    if '年' in date_str:
        date_str = date_str.replace('年', '/').replace('月', '/').replace('日', '')
    date_str = date_str.replace(' ', '')

    formats = [
        "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d",
        "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S",
        "%Y年%m月%d日", "%Y%m%d"
    ]
    for fmt in formats:
        try:
            if fmt == "%Y%m%d" and not date_str.isdigit():
                continue
            dt = datetime.strptime(date_str, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    print(f"警告: 日期字符串 '{date_str}' 无法转换为毫秒时间戳")
    return None


def update_feishu_record(app_token, table_id, record_id, fields, access_token):
    """更新飞书表格记录，fields 的 key 为字段名称"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {"fields": fields}
    response = requests.put(url, headers=headers, json=payload)
    if response.status_code != 200:
        raise Exception(f"更新飞书记录失败: {response.status_code} {response.text}")
    data = response.json()
    if data.get("code") != 0:
        raise Exception(f"更新飞书记录失败: {data.get('code')} {data.get('msg')}")


def add_feishu_record(app_token, table_id, fields, access_token):
    """添加飞书表格记录，fields 的 key 为字段名称"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {"fields": fields}
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code != 200:
        raise Exception(f"添加飞书记录失败: {response.status_code} {response.text}")
    data = response.json()
    if data.get("code") != 0:
        raise Exception(f"添加飞书记录失败: {data.get('code')} {data.get('msg')}")

def save_excel_with_retry(df, filepath, sheet_name='Sheet1'):
    """保存DataFrame到Excel，如果文件被锁定则提示用户关闭后重试"""
    while True:
        try:
            df.to_excel(filepath, index=False, sheet_name=sheet_name)
            print(f"成功保存文件: {filepath}")
            break
        except PermissionError:
            print(f"错误: 文件 '{filepath}' 已被其他程序锁定（例如Excel已打开）。")
            input("请关闭该文件后，按回车键重试...")
        except Exception as e:
            print(f"保存文件时发生未知错误: {e}")
            input("请检查后按回车键重试...")

def process_project_data():
    """主函数：处理项目数据"""
    driver = None
    wait = None
    try:
        print("开始飞书项目信息表更新流程...")

        # 1. 加载配置
        config = load_config_module()
        APP_ID = config.APP_ID
        APP_SECRET = config.APP_SECRET
        APP_TOKEN = config.APP_TOKEN
        TABLE_ID_P = config.TABLE_ID_P
        VIEW_ID_P = config.VIEW_ID_P

        print("配置加载完成")

        # 2. 设置浏览器和登录PLM系统
        print("正在设置浏览器和登录PLM系统...")
        driver, wait = setup_driver()
        PLMLogin(driver, wait)
        print("PLM系统登录成功")

        # 3. 下载项目信息表
        print("正在下载项目信息表...")
        download_file = DL_Project_List(driver, wait)
        print(f"项目信息表下载完成: {download_file}")

        # 4. 读取项目信息表数据
        print("正在读取项目信息表数据...")
        raw_data = RD_Project_List(download_file)
        if isinstance(raw_data, tuple) and len(raw_data) == 2:
            df_raw = raw_data[0]
            print(f"读取到{len(df_raw)}行原始数据")
        else:
            df_raw = raw_data
            print(f"读取到{len(df_raw)}行原始数据")

        # 5. 数据筛选和格式化处理（PLM数据）
        print("正在筛选和格式化PLM数据...")
        required_columns = [
            'Project ID', 'Project Name', 'EBP Leader', 'Product Engineer',
            'Project State', 'Phase 2 Gate Exit-DVR', 'Phase 3 Gate Exit-FPR',
            'Phase 4 Gate Exit-CPA'
        ]
        available_columns = [col for col in required_columns if col in df_raw.columns]
        dfPL = df_raw[available_columns].copy()

        if 'Project ID' in dfPL.columns:
            dfPL['Project ID'] = dfPL['Project ID'].astype(str).str.extract(r'(\d+)')[0]

        date_columns_in_pl = ['Phase 2 Gate Exit-DVR', 'Phase 3 Gate Exit-FPR', 'Phase 4 Gate Exit-CPA']
        for col in date_columns_in_pl:
            if col in dfPL.columns:
                dfPL[col] = pd.to_datetime(dfPL[col], errors='coerce').dt.strftime('%Y-%m-%d')
        print(f"PLM数据处理完成，共{len(dfPL)}行")

        # 6. 获取飞书访问令牌
        print("正在获取飞书访问令牌...")
        access_token = get_feishu_access_token(APP_ID, APP_SECRET)
        print("飞书访问令牌获取成功")

        # 7. 读取飞书表格数据（获取所有字段）
        print("正在读取飞书表格数据...")
        feishu_records = get_feishu_table_data(APP_TOKEN, TABLE_ID_P, VIEW_ID_P, access_token)

        feishu_data = []
        for record in feishu_records:
            record_id = record.get('record_id')
            fields = record.get('fields', {})
            row = {'record_id': record_id}
            for field_name, field_value in fields.items():
                row[field_name] = convert_timestamp_to_date(field_value)
            feishu_data.append(row)

        dfPLFS = pd.DataFrame(feishu_data)
        dfPLFS = dfPLFS.fillna('')
        print(f"飞书表格数据读取完成，共{len(dfPLFS)}行，字段：{list(dfPLFS.columns)}")

        # 确保必要字段存在
        if 'Phase 4 score' not in dfPLFS.columns:
            dfPLFS['Phase 4 score'] = ''
        if 'Phase 4 Gate Exit-CPA' not in dfPLFS.columns:
            dfPLFS['Phase 4 Gate Exit-CPA'] = ''

        # 8. 补充EBP Leader信息（使用飞书已有数据 + PLM系统）
        print("正在对比和补充EBP Leader信息...")
        for idx, pl_row in dfPL.iterrows():
            project_id = pl_row['Project ID']
            if pd.isna(pl_row.get('EBP Leader')) or str(pl_row.get('EBP Leader')).strip() == '':
                feishu_match = dfPLFS[dfPLFS['Project ID'] == project_id]
                if not feishu_match.empty and feishu_match.iloc[0].get('EBP Leader'):
                    dfPL.at[idx, 'EBP Leader'] = feishu_match.iloc[0]['EBP Leader']
                    print(f"从飞书数据补充Project ID {project_id}的EBP Leader: {feishu_match.iloc[0]['EBP Leader']}")

        empty_ebp_projects = dfPL[
            dfPL['EBP Leader'].isna() |
            (dfPL['EBP Leader'].astype(str).str.strip() == '')
        ]['Project ID'].tolist()
        if empty_ebp_projects:
            print(f"发现{len(empty_ebp_projects)}个项目的EBP Leader需要从PLM系统获取")
            ebp_leader_result = FindEBPLeader(driver, wait, empty_ebp_projects)
            if ebp_leader_result is not None and not ebp_leader_result.empty:
                for _, row in ebp_leader_result.iterrows():
                    project_id = row['Project ID']
                    ebp_leader = row['EBP Leader']
                    if ebp_leader:
                        dfPL.loc[dfPL['Project ID'] == project_id, 'EBP Leader'] = ebp_leader
                        print(f"从PLM系统获取Project ID {project_id}的EBP Leader: {ebp_leader}")

        # 9. 对比并更新飞书数据（核心字段）
        print("正在对比并更新飞书数据...")
        compare_fields = [
            'Project Name', 'EBP Leader', 'Product Engineer', 'Project State',
            'Phase 2 Gate Exit-DVR', 'Phase 3 Gate Exit-FPR', 'Phase 4 Gate Exit-CPA'
        ]
        updated_count = 0
        added_count = 0

        for col in compare_fields:
            if col not in dfPLFS.columns:
                dfPLFS[col] = ''

        for _, pl_row in dfPL.iterrows():
            project_id = str(pl_row['Project ID']).strip()
            if not project_id:
                continue

            mask = dfPLFS['Project ID'].astype(str).str.strip() == project_id
            feishu_match = dfPLFS[mask]

            if not feishu_match.empty:
                feishu_idx = feishu_match.index[0]
                record_id = dfPLFS.at[feishu_idx, 'record_id']
                fields_to_update = {}

                for field in compare_fields:
                    pl_raw = pl_row.get(field)
                    if pd.isna(pl_raw):
                        pl_value = ''
                    else:
                        pl_value = str(pl_raw).strip()
                    
                    feishu_value = dfPLFS.at[feishu_idx, field]
                    if pd.isna(feishu_value):
                        feishu_value = ''
                    else:
                        feishu_value = str(feishu_value).strip()
                    
                    if pl_value != feishu_value:
                        if field in date_columns_in_pl:
                            converted = convert_to_timestamp(pl_value)
                            if converted is not None:
                                fields_to_update[field] = converted
                        else:
                            fields_to_update[field] = pl_value

                if fields_to_update:
                    update_feishu_record(APP_TOKEN, TABLE_ID_P, record_id, fields_to_update, access_token)
                    updated_count += 1
                    print(f"更新Project ID {project_id}，变更字段：{list(fields_to_update.keys())}")
                    # 同步本地
                    for fld, val in fields_to_update.items():
                        if fld in date_columns_in_pl and isinstance(val, int):
                            dfPLFS.at[feishu_idx, fld] = datetime.fromtimestamp(val/1000).strftime('%Y-%m-%d')
                        else:
                            dfPLFS.at[feishu_idx, fld] = str(val) if val is not None else ''
            else:
                # 新增记录
                fields_to_add = {'Project ID': project_id}
                for field in compare_fields:
                    pl_raw = pl_row.get(field)
                    if pd.isna(pl_raw):
                        value = ''
                    else:
                        value = str(pl_raw).strip()
                    if field in date_columns_in_pl:
                        conv = convert_to_timestamp(value)
                        if conv is not None:
                            fields_to_add[field] = conv
                    else:
                        fields_to_add[field] = value
                add_feishu_record(APP_TOKEN, TABLE_ID_P, fields_to_add, access_token)
                added_count += 1
                print(f"新增Project ID {project_id}")
                new_row = {'record_id': f'new_{project_id}', 'Project ID': project_id}
                for fld in compare_fields:
                    new_row[fld] = str(pl_row.get(fld, '')) if pd.notna(pl_row.get(fld)) else ''
                dfPLFS = pd.concat([dfPLFS, pd.DataFrame([new_row])], ignore_index=True)

        print(f"飞书核心数据更新完成：更新{updated_count}条，新增{added_count}条")

        # ========== 新增功能：填充“项目号”、“工作令号”、“项目名”空值并同步飞书 ==========
        print("正在检查和补充项目号、工作令号、项目名...")
        
        # 确保 dfPLFS 中有这些中文字段列
        for col in ['项目号', '工作令号', '项目名']:
            if col not in dfPLFS.columns:
                dfPLFS[col] = ''
        
        # 辅助函数：提取工作令号（从第二个逗号之后查找连续6位数字）
        def extract_work_order(project_name):
            if not isinstance(project_name, str):
                return '000000'
            commas = [i for i, ch in enumerate(project_name) if ch == ',']
            if len(commas) < 2:
                # 如果逗号不足2个，回退到全局搜索
                match = re.search(r'\d{6}', project_name)
                return match.group(0) if match else '000000'
            start = commas[1] + 1
            sub_str = project_name[start:]
            match = re.search(r'\d{6}', sub_str)
            if match:
                return match.group(0)
            return '000000'
        
        # 辅助函数：提取项目名（第二个逗号到第五个逗号之间的文本，删除逗号并去除空格）
        def extract_project_name(project_name):
            if not isinstance(project_name, str):
                return ''
            commas = [i for i, ch in enumerate(project_name) if ch == ',']
            if len(commas) < 5:
                print(f"警告: Project Name '{project_name}' 中逗号不足5个，无法提取项目名")
                return ''
            start = commas[1] + 1
            end = commas[4]
            extracted = project_name[start:end].strip()
            extracted = extracted.replace(',', '')
            return extracted.strip()
        
        # 遍历 dfPLFS，填充空值并记录需要飞书更新的行
        updates_for_feishu = []  # 存储 (record_id, fields_dict)
        for idx, row in dfPLFS.iterrows():
            record_id = row.get('record_id')
            if not record_id:
                continue
            
            project_id = row.get('Project ID', '')
            project_name = row.get('Project Name', '')
            if not project_id:
                continue
            
            fields_to_update = {}
            
            # 项目号为空
            if pd.isna(row.get('项目号')) or str(row.get('项目号')).strip() == '':
                new_value = str(project_id).strip()
                fields_to_update['项目号'] = new_value
                dfPLFS.at[idx, '项目号'] = new_value
                print(f"Project ID {project_id} 项目号为空，填充为 {new_value}")
            
            # 工作令号为空
            if pd.isna(row.get('工作令号')) or str(row.get('工作令号')).strip() == '':
                new_value = extract_work_order(project_name)
                fields_to_update['工作令号'] = new_value
                dfPLFS.at[idx, '工作令号'] = new_value
                print(f"Project ID {project_id} 工作令号为空，填充为 {new_value}")
            
            # 项目名为空
            if pd.isna(row.get('项目名')) or str(row.get('项目名')).strip() == '':
                new_value = extract_project_name(project_name)
                fields_to_update['项目名'] = new_value
                dfPLFS.at[idx, '项目名'] = new_value
                print(f"Project ID {project_id} 项目名为空，填充为 {new_value}")
            
            if fields_to_update:
                updates_for_feishu.append((record_id, fields_to_update))
        
        # 批量更新飞书（逐条）
        for record_id, fields in updates_for_feishu:
            try:
                update_feishu_record(APP_TOKEN, TABLE_ID_P, record_id, fields, access_token)
                print(f"已同步飞书记录 {record_id}，更新字段：{list(fields.keys())}")
            except Exception as e:
                print(f"同步飞书记录 {record_id} 失败: {e}")
        
        print(f"项目号/工作令号/项目名补充完成，共更新飞书 {len(updates_for_feishu)} 条记录")

        # ========== 继续原有流程：生成发运前检查计划和导出 ==========
        # 生成发运前检查计划（基于 Phase 4 score 和 Phase 4 Gate Exit-CPA）
        print("正在生成发运前检查计划...")
        today = datetime.now()
        days_until_sunday = (6 - today.weekday()) % 7
        if days_until_sunday == 0:
            days_until_sunday = 7
        next_week_end = today + timedelta(days=days_until_sunday)

        date_col = 'Phase 4 Gate Exit-CPA'
        score_col = 'Phase 4 score'
        if date_col not in dfPLFS.columns:
            date_col = 'phase 4 gate exit-cpa'
        if score_col not in dfPLFS.columns:
            score_col = 'phase 4 score'
        
        if date_col in dfPLFS.columns:
            dfPLFS['cpa_date'] = pd.to_datetime(dfPLFS[date_col], errors='coerce')
            if score_col not in dfPLFS.columns:
                dfPLFS[score_col] = ''
            shipping_check = dfPLFS[
                (dfPLFS[score_col].isna() | (dfPLFS[score_col].astype(str).str.strip() == '')) &
                (dfPLFS['cpa_date'].notna()) &
                (dfPLFS['cpa_date'] <= next_week_end)
            ].copy()
            shipping_check.drop(columns=['cpa_date'], errors='ignore', inplace=True)
        else:
            print(f"警告: 未找到日期列 {date_col}，跳过发运前检查计划生成")
            shipping_check = pd.DataFrame()

        # 定义期望的列顺序（首字母大写的 Phase 2/3/4 score）
        desired_columns = [
            'Project ID', 'Project Name', 'EBP Leader', 'Product Engineer', 'Project State',
            'Phase 2 Gate Exit-DVR', 'Phase 3 Gate Exit-FPR', 'Phase 4 Gate Exit-CPA',
            '项目号', '工作令号', '项目名', 'Phase 2 score', 'Phase 3 score', 'Phase 4 score'
        ]

        # 准备导出的完整项目信息表（去除临时列）
        dfPLFS_export = dfPLFS.drop(columns=['record_id', 'cpa_date'], errors='ignore')
        
        # 大小写不敏感列映射
        col_map = {}
        df_cols_lower = {col.lower(): col for col in dfPLFS_export.columns}
        for desired in desired_columns:
            if desired in dfPLFS_export.columns:
                col_map[desired] = desired
            else:
                lower_desired = desired.lower()
                if lower_desired in df_cols_lower:
                    col_map[desired] = df_cols_lower[lower_desired]
                else:
                    print(f"警告: 期望的列 '{desired}' 在数据中不存在，已跳过")
        existing_columns = [col_map[d] for d in desired_columns if d in col_map]
        dfPLFS_export = dfPLFS_export[existing_columns]

        # 处理发运前检查计划：确保即使为空也生成含列头的空表
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "项目信息表")
        os.makedirs(output_dir, exist_ok=True)

        # 使用相同的列顺序
        shipping_columns = existing_columns
        if shipping_check.empty:
            shipping_check = pd.DataFrame(columns=shipping_columns)
        else:
            # 确保列顺序一致（只保留存在于 shipping_columns 中的列）
            shipping_check = shipping_check[[col for col in shipping_columns if col in shipping_check.columns]]
        
        shipping_file = os.path.join(output_dir, "发运前检查计划.xlsx")
        save_excel_with_retry(shipping_check, shipping_file)
        
        if not shipping_check.empty:
            print(f"发运前检查计划已导出：{shipping_file}")
            # 自动打开Excel文件
            try:
                if platform.system() == 'Windows':
                    os.startfile(shipping_file)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', shipping_file])
                else:  # Linux
                    subprocess.run(['xdg-open', shipping_file])
                print(f"已自动打开文件：{shipping_file}")
            except Exception as e:
                print(f"自动打开文件失败：{e}")
        else:
            print(f"发运前检查计划为空，已生成空表：{shipping_file}")

        project_info_file = os.path.join(output_dir, "项目信息表.xlsx")
        save_excel_with_retry(dfPLFS_export, project_info_file)
        print(f"项目信息表已导出：{project_info_file}")

        print("飞书项目信息表更新流程完成！")

    except Exception as e:
        print(f"飞书项目信息表更新过程中出现错误: {e}")
        traceback.print_exc()
        raise
    finally:
        if driver is not None:
            try:
                driver.quit()
                print("浏览器已关闭")
            except Exception as close_error:
                print(f"关闭浏览器时出现错误: {close_error}")


if __name__ == "__main__":
    process_project_data()