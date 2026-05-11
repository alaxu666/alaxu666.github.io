import requests
import json

# 配置信息
TENANT_ACCESS_TOKEN = "t-g10458eGZTFLTB4JJYW3I3FT3W4GSOK6CBCOZSWF"
BITABLE_APP_TOKEN = "VMSRbxAWGaElLosVgsbcp3wGnkc"
APP_ID = "cli_a977c868ecf95bc4"

# API 地址
url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/collaborators"

# 请求头
headers = {
    "Authorization": f"Bearer {TENANT_ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

# 请求体：添加应用为可编辑协作者
payload = {
    "member_id": APP_ID,
    "member_type": "app",  # 关键：指定类型为应用
    "permitted_role": "editor"  # 可编辑权限
}

# 发送请求
response = requests.post(url, headers=headers, data=json.dumps(payload))
result = response.json()

if result.get("code") == 0:
    print("✅ 应用已成功添加为多维表格协作者！")
else:
    print(f"❌ 添加失败: {result}")