import requests

APP_TOKEN = "DWRubGKJyaOxp1sF5FXcig5NnSd"
TABLE_ID = "tbl1G062INm1DObt"
# 使用你已有的token（先获取一次）
token = "t-g10459cxYMQJZBOEXK72T2LPYQQA2Q3MLUH5VPKK"

url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields"
headers = {"Authorization": f"Bearer {token}"}
resp = requests.get(url, headers=headers)
data = resp.json()
print("字段名称\t字段ID\t\t类型")
for f in data["data"]["items"]:
    print(f"{f['field_name']}\t{f['field_id']}\t{f['type']}")