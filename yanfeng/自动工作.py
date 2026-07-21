import subprocess
import sys
import time
from datetime import datetime

# 要执行的 .py 文件路径
py_files = {
    "pkr": r"C:\XSR\githubPage\yanfeng\PKR数据爬取.py",
    "caigou": r"C:\XSR\githubPage\yanfeng\采购自动确认.py",
    "kucun": r"C:\XSR\githubPage\yanfeng\库存统计.py",
    "feishu": r"C:\XSR\githubPage\yanfeng\飞书项目信息表更新.py",
    "paopao": r"C:\XSR\githubPage\yanfeng\泡泡图数据统计.py",
}

# 记录每个任务上一次执行的日期（避免同一分钟内重复执行）
last_run = {
    "pkr": None,
    "caigou": None,
    "kucun": None,
    "feishu": None,
    "paopao": None,
}


def run_py(path, task_name):
    now = datetime.now()
    # 避免在一分钟内重复执行（因为循环可能多次触发）
    if last_run[task_name] == now.date() and now.minute == last_run.get(task_name + "_min"):
        return
    try:
        print(f"[{now}] 执行 {task_name}: {path}")
        subprocess.run([sys.executable, path], check=True)
        print(f"[{now}] 完成 {task_name}")
        last_run[task_name] = now.date()
        last_run[task_name + "_min"] = now.minute
    except Exception as e:
        print(f"[{now}] 失败 {task_name}: {e}")


def should_run_now(task_name, target_time, target_weekday=None, target_day_of_month=None):
    now = datetime.now()
    # 时间匹配（忽略秒）
    if now.strftime("%H:%M") != target_time:
        return False
    # 如果指定了星期，检查星期几（weekday: 0=周一, 1=周二...6=周日）
    if target_weekday is not None and now.weekday() != target_weekday:
        return False
    # 如果指定了每月起始日，检查日期是否 >= 起始日
    if target_day_of_month is not None and now.day < target_day_of_month:
        return False
    return True


print("标准库调度器已启动，按 Ctrl+C 停止...")
while True:
    now = datetime.now()

    # 每天 10:00 执行 PKR数据爬取.py
    if should_run_now("pkr", "10:00"):
        run_py(py_files["pkr"], "pkr")

    # 每天 15:40 执行 采购自动确认.py
    if should_run_now("caigou", "15:40"):
        run_py(py_files["caigou"], "caigou")

    # 周二 14:00 执行 库存统计.py（weekday=1 表示周二）
    if should_run_now("kucun", "14:00", target_weekday=1):
        run_py(py_files["kucun"], "kucun")

    # 周一 09:00 执行 飞书项目信息表更新.py（weekday=0 表示周一）
    if should_run_now("feishu", "09:00", target_weekday=0):
        run_py(py_files["feishu"], "feishu")

    # 每月25日起每天 09:30 执行 泡泡图数据统计.py（target_day_of_month=25 表示25日开始）
    if should_run_now("paopao", "09:30", target_day_of_month=25):
        run_py(py_files["paopao"], "paopao")

    time.sleep(30)  # 每30秒检查一次
