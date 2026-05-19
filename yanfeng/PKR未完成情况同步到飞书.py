import requests
import pandas as pd
import time
from datetime import datetime
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_loader import load_config_module

config = load_config_module()
APP_ID = config.APP_ID
APP_SECRET = config.APP_SECRET
APP_TOKEN = config.APP_TOKEN
TABLE_ID = config.TABLE_ID
VIEW_ID = config.VIEW_ID
EXCEL_PATH = config.EXCEL_PATH

# ==================== 获取 token ====================
def get_tenant_access_token(app_id, app_secret):
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    payload = {"app_id": app_id, "app_secret": app_secret}
    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取 token 失败: {data}")
    return data["tenant_access_token"]

# ==================== 获取字段详细配置（包括选项列表） ====================
def get_field_details(access_token):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields"
    headers = {"Authorization": f"Bearer {access_token}"}
    all_fields = []
    page_token = None
    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"获取字段失败: {data}")
        items = data["data"]["items"]
        all_fields.extend(items)
        if not data["data"].get("has_more"):
            break
        page_token = data["data"]["page_token"]
    
    # 整理字段信息（同时支持ID和名称查找）
    field_info_by_id = {}
    field_info_by_name = {}
    for f in all_fields:
        field_info_by_id[f["field_id"]] = {
            "name": f["field_name"],
            "type": f["type"],
            "property": f.get("property", {}),
            "id": f["field_id"]
        }
        field_info_by_name[f["field_name"]] = field_info_by_id[f["field_id"]]

    # 返回两个映射
    return field_info_by_id, field_info_by_name

# ==================== 将Excel值转换为飞书API接受的值 ====================
def convert_value(value, field_info):
    """根据字段类型转换值，返回转换后的值，如果无法转换则返回None"""
    if pd.isna(value):
        return None
    
    field_type = field_info["type"]
    # 文本（type=1）
    if field_type == 1:
        return str(value)
    # 数字（type=2）
    elif field_type == 2:
        try:
            return float(value)
        except:
            return None
    # 单选（type=3）
    elif field_type == 3:
        options = field_info["property"].get("options", [])
        # 根据选项文本匹配ID
        val_str = str(value).strip()
        for opt in options:
            if opt["name"] == val_str:
                return opt["id"]
        # 未匹配到选项，返回None并警告
        print(f"  警告: 单选字段 '{field_info['name']}' 值 '{val_str}' 不在选项中，可用选项: {[opt['name'] for opt in options]}")
        return None
    # 多选（type=4）
    elif field_type == 4:
        # 假设Excel中用逗号分隔多个选项
        options = field_info["property"].get("options", [])
        val_str = str(value).strip()
        selected_ids = []
        for part in val_str.split(','):
            part = part.strip()
            for opt in options:
                if opt["name"] == part:
                    selected_ids.append(opt["id"])
                    break
        return selected_ids if selected_ids else None
    # 日期（type=5）
    elif field_type == 5:  # 日期字段
        # 1. 已经是数字（毫秒时间戳）
        if isinstance(value, (int, float)):
            return int(value)

        # 2. Python 或 Pandas 的 datetime 对象
        if isinstance(value, (datetime, pd.Timestamp)):
            return int(value.timestamp() * 1000)

        # 3. 处理字符串（这是您的数据所在的分支）
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return None

            # 3.1 首先尝试用正则提取日期部分，忽略时间部分
            import re
            match = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', value)
            if match:
                year, month, day = map(int, match.groups())
                try:
                    # 创建一个表示当地日期的 datetime 对象，时间部分为 00:00:00
                    dt = datetime(year, month, day)
                    return int(dt.timestamp() * 1000)  # 转换为毫秒时间戳
                except ValueError as e:
                    print(f"  警告: 日期字段 '{field_info['name']}' 的值 '{value}' 包含无效日期: {e}")
                    return None

            # 3.2 处理其他可能遇到的常见格式（按优先级排序）
            formats = [
                "%Y-%m-%d %H:%M:%S",  # 完整的日期时间
                "%Y/%m/%d %H:%M:%S",  # 另一种分隔符
                "%Y-%m-%d",           # 只有日期
                "%Y/%m/%d",           # 另一种分隔符
                "%Y.%m.%d"            # 点号分隔符
            ]

            for fmt in formats:
                try:
                    dt = datetime.strptime(value, fmt)
                    return int(dt.timestamp() * 1000)
                except ValueError:
                    continue  # 这个格式不行，继续尝试下一个

        # 如果走到这里，说明所有转换都失败了
        print(f"  警告: 日期字段 '{field_info['name']}' 的值 '{value}' 无法转换为毫秒级时间戳，已跳过")
        return None
    
    # 复选框（type=7）
    elif field_type == 7:
        val_str = str(value).lower()
        return val_str in ['true', '1', 'yes', '是']
    # 人员（type=8）
    elif field_type == 8:
        # 飞书人员字段需要传入用户的open_id或user_id列表
        # 这里简单处理：如果Excel中是邮箱或工号，需要查找映射，跳过复杂逻辑，先保持原值并警告
        print(f"  警告: 人员字段 '{field_info['name']}' 需要特殊处理，当前直接传字符串可能失败")
        return str(value)
    # 其他类型暂时按文本处理
    else:
        return str(value)

# ==================== 获取所有现有记录 ====================
def get_all_existing_records(access_token):
    """获取飞书表格中的所有记录，返回记录ID列表"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"page_size": 100, "view_id": VIEW_ID}
    all_records = []
    page_token = None
    while True:
        if page_token:
            params["page_token"] = page_token
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"获取现有记录失败: {data}")

        # 检查数据结构
        if "data" not in data or "items" not in data["data"]:
            # 表格为空时可能没有items字段，这是正常的
            break

        items = data["data"]["items"]
        all_records.extend(items)
        if not data["data"].get("has_more"):
            break
        page_token = data["data"]["page_token"]

    # 返回所有记录的ID
    record_ids = [rec["record_id"] for rec in all_records]
    print(f"OK 获取到 {len(record_ids)} 条现有记录")
    return record_ids

# ==================== 获取现有记录的 Project ID -> record_id 映射 ====================
def get_existing_mapping(access_token, project_id_field_id):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"page_size": 100, "view_id": VIEW_ID}
    all_records = []
    page_token = None
    while True:
        if page_token:
            params["page_token"] = page_token
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"获取现有记录失败: {data}")

        # 检查数据结构
        if "data" not in data or "items" not in data["data"]:
            # 表格为空时可能没有items字段，这是正常的
            break

        items = data["data"]["items"]
        all_records.extend(items)
        if not data["data"].get("has_more"):
            break
        page_token = data["data"]["page_token"]

    mapping = {}
    for rec in all_records:
        fields = rec.get("fields", {})
        if project_id_field_id in fields:
            proj_id_val = fields[project_id_field_id]
            if proj_id_val:
                mapping[str(proj_id_val)] = rec["record_id"]
    print(f"OK 已建立 {len(mapping)} 个现有 Project ID -> record_id 映射")
    return mapping

# ==================== 批量删除记录 ====================
def batch_delete_records(access_token, record_ids):
    if not record_ids:
        return
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/batch_delete"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    batch_size = 100
    for i in range(0, len(record_ids), batch_size):
        batch = record_ids[i:i+batch_size]
        payload = {"records": batch}
        resp = requests.post(url, headers=headers, json=payload)
        resp_data = resp.json()
        if resp.status_code != 200 or resp_data.get("code") != 0:
            print(f"ERROR 批量删除失败: {resp_data}")
            return
        else:
            print(f"OK 已删除 {len(batch)} 条记录")
        time.sleep(0.5)

# ==================== 批量操作（带详细错误输出） ====================
def batch_create_records(access_token, creations):
    if not creations:
        return
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/batch_create"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    batch_size = 100
    for i in range(0, len(creations), batch_size):
        batch = creations[i:i+batch_size]
        payload = {"records": batch}
        resp = requests.post(url, headers=headers, json=payload)
        resp_data = resp.json()
        if resp.status_code != 200 or resp_data.get("code") != 0:
            print(f"ERROR 批量新增失败: {resp_data}")
            return
        else:
            print(f"OK 已新增 {len(batch)} 条记录")
        time.sleep(0.5)

def batch_update_records(access_token, updates):
    if not updates:
        return
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/batch_update"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    batch_size = 100
    for i in range(0, len(updates), batch_size):
        batch = updates[i:i+batch_size]
        payload = {"records": batch}
        resp = requests.post(url, headers=headers, json=payload)
        resp_data = resp.json()
        if resp.status_code != 200 or resp_data.get("code") != 0:
            print(f"ERROR 批量更新失败: {resp_data}")
            return
        else:
            print(f"OK 已更新 {len(batch)} 条记录")
        time.sleep(0.5)

# ==================== 主程序 ====================
def main():
    print("1. 获取访问凭证...")
    token = get_tenant_access_token(APP_ID, APP_SECRET)
    print("OK 获取成功")
    
    print("2. 获取飞书表格字段详细信息...")
    field_info_by_id, field_info_by_name = get_field_details(token)
    print(f"已获取 {len(field_info_by_id)} 个字段")

    
    # 确认 Project ID 字段存在
    project_id_field = None
    for fid, info in field_info_by_id.items():
        if info["name"] == "Project ID":
            project_id_field = fid
            break
    if not project_id_field:
        raise Exception("表格中没有 'Project ID' 字段")
    print(f"OK Project ID 字段ID: {project_id_field}")
    
    print("3. 读取 Excel 文件...")
    df = pd.read_excel(EXCEL_PATH)
    df = df.dropna(how="all")
    print(f"OK Excel 中共有 {len(df)} 行数据")
    
    # 构建 Excel 列名到字段ID的映射
    col_to_field_id = {}
    for col in df.columns:
        for fid, info in field_info_by_id.items():
            if info["name"] == col:
                col_to_field_id[col] = fid
                break
    print(f"OK 匹配到 {len(col_to_field_id)} 个字段")
    
    print("4. 获取所有现有记录用于删除...")
    all_record_ids = get_all_existing_records(token)
    print(f"5. 发现 {len(all_record_ids)} 条现有记录")
    if all_record_ids:
        print(f"   删除所有现有记录...")
        batch_delete_records(token, all_record_ids)
    else:
        print("   没有现有记录需要删除")

    print("6. 获取现有记录映射...")
    existing_mapping = get_existing_mapping(token, project_id_field)

    creations = []

    # 检查Excel是否为空
    if len(df) == 0:
        print("7. Excel文件为空，创建默认更新记录...")
        # 创建一条特殊记录
        default_fields = {}
        current_date = datetime.now()

        # 为每个字段设置默认值
        for col_name, field_id in col_to_field_id.items():
            if field_id not in field_info_by_id:
                continue

            field_info = field_info_by_id[field_id]
            field_type = field_info["type"]

            if col_name == "未完成人名":
                # 日期时间格式
                default_fields[field_info["name"]] = current_date.strftime("%Y-%m-%d %H:%M:%S")
            elif col_name in ["Phase 1 Gate Exit-GO", "Phase 2 Gate Exit-DVR", "Phase 3 Gate Exit-FPR", "Phase 4 Gate Exit-CPA", "Phase 5 Gate Exit-PLR"]:
                # 日期字段，只填日期
                if field_type == 5:  # 日期类型
                    dt = datetime(current_date.year, current_date.month, current_date.day)
                    default_fields[field_info["name"]] = int(dt.timestamp() * 1000)
                else:
                    default_fields[field_info["name"]] = current_date.strftime("%Y-%m-%d")
            else:
                # 其他字段填入"表格更新日期"
                default_fields[field_info["name"]] = "表格更新日期"

        creations.append({"fields": default_fields})
        print("   已创建默认更新记录")
    else:
        print(f"7. 处理Excel数据 ({len(df)} 行)...")

    for idx, row in df.iterrows():
        proj_id_raw = row.get("Project ID")
        if pd.isna(proj_id_raw):
            continue
        proj_id = str(proj_id_raw)
        
        fields_by_name = {}  # 使用字段名而不是ID
        skip_fields = []
        for col_name, value in row.items():
            if col_name not in col_to_field_id:
                continue
            field_id = col_to_field_id[col_name]
            # 检查字段ID是否存在于field_info中
            if field_id not in field_info_by_id:
                print(f"  警告: Excel列 '{col_name}' 映射到不存在的字段ID '{field_id}'")
                skip_fields.append(col_name)
                continue
            converted = convert_value(value, field_info_by_id[field_id])
            if converted is None:
                skip_fields.append(col_name)
            else:
                # 使用字段名而不是字段ID
                field_name = field_info_by_id[field_id]["name"]
                fields_by_name[field_name] = converted

        if skip_fields:
            print(f"  项目 {proj_id} 跳过字段: {skip_fields}")

        if not fields_by_name:
            print(f"  项目 {proj_id} 没有有效字段，跳过")
            continue


        if proj_id in existing_mapping:
            updates.append({
                "record_id": existing_mapping[proj_id],
                "fields": fields_by_name
            })
        else:
            creations.append({
                "fields": fields_by_name
            })
    
    print(f"\nSUMMARY 待处理: 新增 {len(creations)} 条")

    if creations:
        print("\n7. 批量新增...")
        batch_create_records(token, creations)
    
    print("\nDONE 同步完成")

def sync_pkr_data_to_feishu():
    """
    将PKR未完成情况数据同步到飞书表格
    这个函数可以被其他模块调用
    """
    main()

if __name__ == "__main__":
    sync_pkr_data_to_feishu()