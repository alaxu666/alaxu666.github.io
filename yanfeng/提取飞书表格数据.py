#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
提取飞书表格数据
功能：读取飞书多维表格的字段类型，根据类型提取完整数据，导出到Excel
支持飞书字段类型：
  1=文本  2=数字  3=单选  4=多选  5=日期  7=复选框  8=人员
  11=超链接  13=创建时间  14=更新时间  17=附件/图片  18=关联记录
  19=公式(查找引用)  20=公式(计算结果)  3001=按钮
"""

import os
import sys
import re
import json
import io
import requests
import pandas as pd
from datetime import datetime

# 修复 Windows 控制台 GBK 编码问题（表格名称可能含 emoji 等字符）
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_loader import load_config_module


# ==================== 飞书API基础函数 ====================

def get_tenant_access_token(app_id, app_secret):
    """获取飞书tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    payload = {"app_id": app_id, "app_secret": app_secret}
    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取token失败: {data}")
    return data["tenant_access_token"]


def get_table_names(app_token, table_id, access_token):
    """
    获取多维表格名称（app名称）和当前表格名称（table名称）
    返回: str, 格式为 "app名称_table_name"
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    # 1. 获取多维表格应用名称
    url_app = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}"
    resp_app = requests.get(url_app, headers=headers)
    resp_app.raise_for_status()
    data_app = resp_app.json()
    if data_app.get("code") != 0:
        raise Exception(f"获取多维表格名称失败: {data_app}")
    app_name = data_app["data"]["app"]["name"]

    # 2. 获取当前表格名称（从表格列表中匹配 table_id）
    url_tables = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables"
    table_name = ""
    page_token = None
    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token
        resp_tables = requests.get(url_tables, headers=headers, params=params)
        resp_tables.raise_for_status()
        data_tables = resp_tables.json()
        if data_tables.get("code") != 0:
            raise Exception(f"获取表格列表失败: {data_tables}")
        for t in data_tables.get("data", {}).get("items", []):
            if t.get("table_id") == table_id:
                table_name = t.get("name", "")
                break
        if table_name or not data_tables.get("data", {}).get("has_more"):
            break
        page_token = data_tables.get("data", {}).get("page_token")

    # 3. 拼接：app名称_table_name
    if table_name:
        return f"{app_name}_{table_name}"
    return app_name


def get_table_fields(app_token, table_id, access_token):
    """
    获取表格所有字段的详细信息（字段ID、名称、类型、选项等）
    返回: list of dict, 每个dict包含 field_id, field_name, type, property
    """
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
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
        page_token = data["data"].get("page_token")

    return all_fields


def get_table_records(app_token, table_id, access_token):
    """
    获取表格所有记录
    返回: list of dict, 每个dict包含 record_id 和 fields
    """
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    headers = {"Authorization": f"Bearer {access_token}"}
    all_records = []
    page_token = None

    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"获取记录失败: {data}")
        items = data.get("data", {}).get("items", [])
        all_records.extend(items)
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data.get("data", {}).get("page_token")

    return all_records


# ==================== 字段值解析 ====================

def parse_rich_text(text_list):
    """
    解析飞书富文本字段值
    格式: [{"type": "text", "text": "内容"}, {"type": "mention", ...}, ...]
    也兼容纯字符串或数字
    """
    if text_list is None:
        return ""
    if isinstance(text_list, str):
        return text_list
    if isinstance(text_list, (int, float)):
        # 去除不必要的 .0
        if isinstance(text_list, float) and text_list == int(text_list):
            return str(int(text_list))
        return str(text_list)
    if not isinstance(text_list, list):
        return str(text_list)

    parts = []
    for item in text_list:
        if isinstance(item, dict):
            t = item.get("type", "")
            if t == "text":
                parts.append(item.get("text", ""))
            elif t == "mention":
                parts.append(item.get("text", ""))
            elif t == "equation":
                parts.append(item.get("text", ""))
            else:
                parts.append(item.get("text", str(item)))
        else:
            parts.append(str(item))
    return "".join(parts)


def parse_single_select(options):
    """解析飞书单选字段值（type=3），返回选项名称"""
    if options is None:
        return ""
    if isinstance(options, list):
        if len(options) == 0:
            return ""
        if isinstance(options[0], dict):
            return options[0].get("name", str(options[0]))
        return str(options[0])
    if isinstance(options, dict):
        return options.get("name", str(options))
    return str(options)


def parse_multi_select(options):
    """解析飞书多选字段值（type=4），多个选项用逗号分隔"""
    if options is None:
        return ""
    if not isinstance(options, list):
        return str(options)
    names = []
    for opt in options:
        if isinstance(opt, dict):
            names.append(opt.get("name", str(opt)))
        else:
            names.append(str(opt))
    return ", ".join(names)


def parse_date(timestamp_ms):
    """
    解析飞书日期/时间戳字段值（type=5, 13, 14）
    毫秒时间戳 -> YYYY-MM-DD 格式
    """
    if timestamp_ms is None:
        return ""
    if isinstance(timestamp_ms, str):
        return timestamp_ms
    try:
        ts = int(timestamp_ms) / 1000
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
    except (ValueError, OSError, OverflowError):
        return str(timestamp_ms)


def parse_number(value):
    """
    解析飞书数字字段值（type=2, 20）
    保持数值精度，去除不必要的 .0
    """
    if value is None:
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    if isinstance(value, str):
        return value
    return str(value)


def parse_checkbox(value):
    """解析飞书复选框字段值（type=7）"""
    if value is True or str(value).lower() in ("true", "1", "yes", "是"):
        return "是"
    return "否"


def parse_user_list(users):
    """
    解析飞书人员字段值（type=8, 11中的超链接人员）
    返回人员名称，多人用逗号分隔
    """
    if users is None:
        return ""
    if not isinstance(users, list):
        return str(users)
    names = []
    for user in users:
        if isinstance(user, dict):
            name = user.get("name", "")
            if not name:
                name = user.get("en_name", "")
            if not name:
                name = user.get("user_id", user.get("email", str(user)))
            names.append(name)
        else:
            names.append(str(user))
    return ", ".join(names)


def parse_attachment(attachments):
    """
    解析飞书附件/图片字段值（type=17）
    返回文件名列表，用逗号分隔
    """
    if attachments is None:
        return ""
    if not isinstance(attachments, list):
        return str(attachments)
    names = []
    for att in attachments:
        if isinstance(att, dict):
            name = att.get("name", att.get("file_token", str(att)))
            names.append(name)
        else:
            names.append(str(att))
    return ", ".join(names)


def parse_link(linked_records):
    """
    解析飞书关联记录字段值（type=18）
    格式: [{"record_ids": [...], "text": "显示文本", "table_id": "..."}]
    返回显示文本
    """
    if linked_records is None:
        return ""
    if not isinstance(linked_records, list):
        return str(linked_records)
    texts = []
    for rec in linked_records:
        if isinstance(rec, dict):
            text = rec.get("text", "")
            if not text:
                record_ids = rec.get("record_ids", [])
                text = ", ".join(str(rid) for rid in record_ids)
            texts.append(text)
        else:
            texts.append(str(rec))
    return ", ".join(texts)


def parse_formula(value):
    """
    解析飞书公式字段值（type=19, 20）
    公式结果可能是富文本数组、数字或字符串
    """
    if value is None:
        return ""
    if isinstance(value, list):
        return parse_rich_text(value)
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value)


# ==================== 字段类型映射 ====================

# field_type -> (类型名称, 解析函数)
FIELD_TYPE_HANDLERS = {
     1: ("文本",     parse_rich_text),
     2: ("数字",     parse_number),
     3: ("单选",     parse_single_select),
     4: ("多选",     parse_multi_select),
     5: ("日期",     parse_date),
     7: ("复选框",   parse_checkbox),
     8: ("人员",     parse_user_list),
    11: ("超链接",   parse_user_list),  # type=11 实际返回人员dict列表
    13: ("创建时间", parse_date),
    14: ("更新时间", parse_date),
    17: ("附件",     parse_attachment),
    18: ("关联记录", parse_link),
    19: ("公式",     parse_formula),
    20: ("公式",     parse_formula),
    3001: ("按钮",   parse_rich_text),
}


def convert_field_value(raw_value, field_type):
    """
    根据字段类型将飞书API返回的原始值转换为可读字符串
    """
    if raw_value is None:
        return ""

    handler = FIELD_TYPE_HANDLERS.get(field_type)
    if handler:
        return handler[1](raw_value)

    # 未知类型，通用处理
    if isinstance(raw_value, (dict, list)):
        return json.dumps(raw_value, ensure_ascii=False)
    return str(raw_value)


# ==================== 主程序 ====================

def main():
    print("=" * 60)
    print("提取飞书表格数据")
    print("=" * 60)

    # 1. 加载配置
    print("\n[1/5] 加载配置...")
    config = load_config_module()
    APP_ID = config.APP_ID
    APP_SECRET = config.APP_SECRET
    TABLE_LINK = config.TABLE_LINK
    print(f"  TABLE_LINK: {TABLE_LINK}")

    # 从TABLE_LINK解析app_token和table_id
    app_token_match = re.search(r'/base/([A-Za-z0-9]+)', TABLE_LINK)
    table_id_match = re.search(r'table=([A-Za-z0-9]+)', TABLE_LINK)

    if not app_token_match or not table_id_match:
        raise ValueError(f"无法从TABLE_LINK解析app_token和table_id: {TABLE_LINK}")

    app_token = app_token_match.group(1)
    table_id = table_id_match.group(1)
    print(f"  app_token: {app_token}")
    print(f"  table_id:  {table_id}")

    # 2. 获取访问令牌
    print("\n[2/5] 获取飞书访问令牌...")
    access_token = get_tenant_access_token(APP_ID, APP_SECRET)
    print("  获取成功")

    # 3. 获取多维表格名称、当前表格名称和字段信息
    print("\n[3/5] 获取多维表格信息...")
    file_name = get_table_names(app_token, table_id, access_token)
    print(f"  导出文件名: {file_name}")
    fields = get_table_fields(app_token, table_id, access_token)
    print(f"  共获取 {len(fields)} 个字段")

    # 构建字段映射: field_name -> field_type
    # 注意：飞书API返回的记录数据中，fields dict 的 key 是字段名称，不是字段ID
    field_type_by_name = {}
    field_names = []
    for f in fields:
        fname = f["field_name"]
        ftype = f["type"]
        field_type_by_name[fname] = ftype
        field_names.append(fname)
        type_name = FIELD_TYPE_HANDLERS.get(ftype, ("未知", None))[0]
        print(f"    [{type_name}] {fname} (ID: {f['field_id']})")

    # 4. 获取所有记录并解析
    print("\n[4/5] 获取表格数据...")
    records = get_table_records(app_token, table_id, access_token)
    print(f"  共获取 {len(records)} 条记录")

    # 解析每条记录
    rows = []
    for record in records:
        row = {}
        fields_data = record.get("fields", {})

        for fname in field_names:
            ftype = field_type_by_name[fname]
            raw_value = fields_data.get(fname)
            row[fname] = convert_field_value(raw_value, ftype)

        rows.append(row)

    df = pd.DataFrame(rows)

    # 按字段原始顺序排列列
    existing_cols = [c for c in field_names if c in df.columns]
    remaining_cols = [c for c in df.columns if c not in field_names]
    df = df[existing_cols + remaining_cols]

    print(f"  数据解析完成: {df.shape[0]} 行 × {df.shape[1]} 列")

    # 5. 导出到Excel
    print("\n[5/5] 导出到Excel...")
    output_dir = os.path.dirname(os.path.abspath(__file__))
    # 使用多维表格名称作为文件名，去除 emoji、非法文件名字符
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', file_name)
    safe_name = safe_name.encode('ascii', 'ignore').decode('ascii').strip()
    safe_name = re.sub(r'_+', '_', safe_name).strip('_')
    if not safe_name:
        safe_name = "飞书表格数据"
    output_path = os.path.join(output_dir, f"{safe_name}.xlsx")

    try:
        df.to_excel(output_path, index=False, sheet_name="Sheet1")
        print(f"  成功导出: {output_path}")
    except PermissionError:
        print(f"  错误: 文件 '{output_path}' 已被锁定（可能Excel已打开）。")
        print(f"  请先关闭该文件，然后重新运行脚本。")
        raise

    print("\n" + "=" * 60)
    print("完成！")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
