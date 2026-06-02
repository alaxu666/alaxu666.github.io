import subprocess
import time
from datetime import datetime, timedelta

bat_files = {
    "pkr": r"C:\Users\uxuxs004\OneDrive - Yanfeng\桌面\py\PKR数据爬取py.bat",
    "caigou": r"C:\Users\uxuxs004\OneDrive - Yanfeng\桌面\py\采购审批.bat",
    "kucun": r"C:\Users\uxuxs004\OneDrive - Yanfeng\桌面\py\库存统计.bat",
    "feishu": r"C:\Users\uxuxs004\OneDrive - Yanfeng\桌面\py\飞书项目信息表更新.bat",
}

# 记录每个任务上一次执行的日期（避免同一分钟内重复执行）
last_run = {
    "pkr": None,
    "caigou": None,
    "kucun": None,
    "feishu": None,
}

def run_bat(path, task_name):
    now = datetime.now()
    # 避免在一分钟内重复执行（因为循环可能多次触发）
    if last_run[task_name] == now.date() and now.minute == last_run.get(task_name+"_min"):
        return
    try:
        print(f"[{now}] 执行 {task_name}: {path}")
        subprocess.run(path, shell=True, check=True)
        print(f"[{now}] 完成 {task_name}")
        last_run[task_name] = now.date()
        last_run[task_name+"_min"] = now.minute
    except Exception as e:
        print(f"[{now}] 失败 {task_name}: {e}")

def should_run_now(task_name, target_time, target_weekday=None):
    now = datetime.now()
    # 时间匹配（忽略秒）
    if now.strftime("%H:%M") != target_time:
        return False
    # 如果指定了星期，检查星期几（weekday: 0=周一, 1=周二...6=周日）
    if target_weekday is not None and now.weekday() != target_weekday:
        return False
    return True

print("标准库调度器已启动，按 Ctrl+C 停止...")
while True:
    now = datetime.now()

    # 每天 10:00 执行 PKR
    if should_run_now("pkr", "10:00"):
        run_bat(bat_files["pkr"], "pkr")

    # 每天 16:00 执行采购
    if should_run_now("caigou", "16:00"):
        run_bat(bat_files["caigou"], "caigou")

    # 周二 14:00 执行库存（weekday=1 表示周二）
    if should_run_now("kucun", "14:00", target_weekday=1):
        run_bat(bat_files["kucun"], "kucun")

    # 周一 09:00 执行飞书（weekday=0 表示周一）
    if should_run_now("feishu", "09:00", target_weekday=0):
        run_bat(bat_files["feishu"], "feishu")

    time.sleep(30)  # 每30秒检查一次