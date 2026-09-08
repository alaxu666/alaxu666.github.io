import requests

URL = "https://ddeyfjkvtekhrbjxoout.supabase.co"
KEY = "sb_publishable_neQqtQmrO3f5Z6IjrkiQ5A_P6ZH_X8q"  # 你的 publishable key

headers = {
    "apikey": KEY,
    "Authorization": f"Bearer {KEY}",
    "Content-Type": "application/json",
}

# 插入一条成员
data = {
    "姓名": "张三",
    "岗位": "工程师",
    "手机号": "13800138000",
    "邮箱": "zhangsan@example.com",
    "英文名": "San Zhang",
    "UID": "U123",
    "部门": "研发部"
}
res = requests.post(f"{URL}/rest/v1/成员", headers=headers, json=data)
print(res.status_code, res.json())

# 查询所有成员
res = requests.get(f"{URL}/rest/v1/成员?select=*", headers=headers)
print(res.json())