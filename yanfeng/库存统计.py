import pandas as pd
import re
import json
from datetime import datetime
import subprocess
import os
import sys

# ================== 原有数据读取和处理（不变） ==================
dfLY = pd.read_excel("C:/XSR/githubPage/yanfeng/库存数据/库存统计.xlsx", sheet_name="本周领用报表")
dfRK = pd.read_excel("C:/XSR/githubPage/yanfeng/库存数据/库存统计.xlsx", sheet_name="本周入库报表")
dfH = pd.read_excel("C:/XSR/githubPage/yanfeng/库存数据/库存统计.xlsx", sheet_name="huang")
dfKC = pd.read_excel("C:/XSR/githubPage/yanfeng/库存数据/本周库存统计.xlsx", sheet_name="2026")

# 从dfH中提取数据
gl_row = dfH[dfH["类别"] == "合计"]
if len(gl_row) > 0:
    GLTP_raw = gl_row.iloc[0]["隔离金额"]
    LYTP_raw = gl_row.iloc[0]["领用金额"]
    GLTP = float(GLTP_raw) if GLTP_raw not in (None, '', 'nan') else 0.0
    LYTP = float(LYTP_raw) if LYTP_raw not in (None, '', 'nan') else 0.0
else:
    GLTP = 0.0
    LYTP = 0.0

# 筛选dfRK数据并计算退库总价
dfRK_filtered = dfRK[dfRK["类别"].isin(["002-结余", "004-备件", "010-保税区"])]
RKTP = dfRK_filtered["总价"].sum()
RKTP = float(RKTP) if not pd.isna(RKTP) else 0.0

# ========== 自动计算上周（要更新的周） ==========
today = datetime.now()
_, current_week, _ = today.isocalendar()
last_week_num = current_week - 1
if last_week_num == 0:
    last_week_num = 52
target_week = f"CW{last_week_num:02d}"
print(f"本周是{current_week}周，将数据写入{target_week}")

# 确保 dfKC 中有目标周的行
if not (dfKC["周数"] == target_week).any():
    new_row = pd.DataFrame({'周数': [target_week], '领用': [0], '隔离': [0], '退库': [0], '2025参考': [0], '每周需复用': [0]})
    dfKC = pd.concat([dfKC, new_row], ignore_index=True)

# 更新dfKC中的数据
dfKC.loc[dfKC["周数"] == target_week, "领用"] = LYTP
dfKC.loc[dfKC["周数"] == target_week, "隔离"] = GLTP
dfKC.loc[dfKC["周数"] == target_week, "退库"] = RKTP

# 在统计分析前，先对dfLY和dfRK进行数据筛选
dfLY = dfLY[dfLY["类别"].isin(["002-结余", "004-备件", "010-保税区"])]
dfRK = dfRK[dfRK["类别"].isin(["002-结余", "004-备件", "010-保税区"])]

# 处理dfLY表格 - 项目统计
dfXM_ZJ = pd.DataFrame(columns=["项目名", "领用总价"])
lXMBH = dfLY["项目编号名称"].drop_duplicates().tolist()
for xmbh in lXMBH:
    total_price = dfLY[dfLY["项目编号名称"] == xmbh]["总价"].sum()
    new_row = pd.DataFrame({"项目名": [xmbh], "领用总价": [total_price]})
    dfXM_ZJ = pd.concat([dfXM_ZJ, new_row], ignore_index=True)

# 处理dfLY表格 - 隔离人统计
dfRM_ZJ = pd.DataFrame(columns=["隔离人", "隔离总价"])
lGLR = dfLY["隔离人"].drop_duplicates().tolist()
for glr in lGLR:
    total_price = dfLY[dfLY["隔离人"] == glr]["总价"].sum()
    new_row = pd.DataFrame({"隔离人": [glr], "隔离总价": [total_price]})
    dfRM_ZJ = pd.concat([dfRM_ZJ, new_row], ignore_index=True)

# 处理dfRK表格 - 退库统计
dfTK_ZJ = pd.DataFrame(columns=["项目名", "退库总价"])
lXMTK = dfRK["项目编码名称"].drop_duplicates().tolist()
for xmtk in lXMTK:
    total_price = dfRK[dfRK["项目编码名称"] == xmtk]["总价"].sum()
    new_row = pd.DataFrame({"项目名": [xmtk], "退库总价": [total_price]})
    dfTK_ZJ = pd.concat([dfTK_ZJ, new_row], ignore_index=True)

# 读取花名册并筛选部门为 PD 和 Manufacturing 的记录
script_dir = os.path.dirname(os.path.abspath(__file__))
roster_path = os.path.join(script_dir, "DATA", "装备中心花名册.xlsx")
if os.path.exists(roster_path):
    try:
        dfCY = pd.read_excel(roster_path, sheet_name="花名册")
        if "部门" in dfCY.columns:
            dfCY = dfCY[dfCY["部门"].isin(["PD", "Manufacturing"])].copy()
        else:
            print(f"警告：花名册中未找到 '部门' 列，保留原始 dfCY")
        if "姓名" not in dfCY.columns:
            print(f"警告：花名册中未找到 '姓名' 列，dfCY 过滤后仍会保留原始数据")
    except Exception as e:
        print(f"读取花名册失败：{e}")
        dfCY = pd.DataFrame(columns=["姓名", "部门"])
else:
    print(f"未找到花名册文件：{roster_path}")
    dfCY = pd.DataFrame(columns=["姓名", "部门"])

# 删除 dfRM_ZJ 中“隔离人”在 dfCY“姓名”中出现的数据
if not dfRM_ZJ.empty and "隔离人" in dfRM_ZJ.columns and "姓名" in dfCY.columns:
    cy_names = dfCY["姓名"].dropna().astype(str).tolist()
    dfRM_ZJ = dfRM_ZJ[~dfRM_ZJ["隔离人"].astype(str).isin(cy_names)].copy()

# 过滤函数：只保留项目名以6位数字开头的数据
def starts_with_6_digits(text):
    if pd.isna(text) or not isinstance(text, str):
        return False
    return bool(re.match(r'^\d{6}', text))

# 安全过滤
if not dfXM_ZJ.empty:
    mask_xm = dfXM_ZJ["项目名"].apply(starts_with_6_digits)
    dfXM_ZJ = dfXM_ZJ.loc[mask_xm]
if not dfTK_ZJ.empty:
    mask_tk = dfTK_ZJ["项目名"].apply(starts_with_6_digits)
    dfTK_ZJ = dfTK_ZJ.loc[mask_tk]

# 排序
if not dfXM_ZJ.empty and '领用总价' in dfXM_ZJ.columns:
    dfXM_ZJ = dfXM_ZJ.sort_values(by="领用总价", ascending=False)
if not dfRM_ZJ.empty and '隔离总价' in dfRM_ZJ.columns:
    dfRM_ZJ = dfRM_ZJ.sort_values(by="隔离总价", ascending=False)
if not dfTK_ZJ.empty and '退库总价' in dfTK_ZJ.columns:
    dfTK_ZJ = dfTK_ZJ.sort_values(by="退库总价", ascending=False)
else:
    print("警告：dfTK_ZJ 无有效数据，跳过排序")

# 保存到Excel文件
output_file = "C:/XSR/githubPage/yanfeng/库存数据/本周库存统计.xlsx"
try:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        dfKC.to_excel(writer, sheet_name="2026", index=False)
        dfXM_ZJ.to_excel(writer, sheet_name="dfXM_ZJ", index=False)
        dfRM_ZJ.to_excel(writer, sheet_name="dfRM_ZJ", index=False)
        dfTK_ZJ.to_excel(writer, sheet_name="dfTK_ZJ", index=False)
    print(f"数据已保存到: {output_file}")
except PermissionError:
    output_file = "C:/XSR/githubPage/yanfeng/库存数据/本周库存统计_新.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        dfKC.to_excel(writer, sheet_name="2026", index=False)
        dfXM_ZJ.to_excel(writer, sheet_name="dfXM_ZJ", index=False)
        dfRM_ZJ.to_excel(writer, sheet_name="dfRM_ZJ", index=False)
        dfTK_ZJ.to_excel(writer, sheet_name="dfTK_ZJ", index=False)
    print(f"数据已保存到: {output_file}")

print("库存统计数据更新完成！")
print(f"领用金额: {LYTP}")
print(f"隔离金额: {GLTP}")
print(f"退库总价: {RKTP}")
print(f"项目统计: {len(dfXM_ZJ)} 个项目")
print(f"隔离人统计: {len(dfRM_ZJ)} 个隔离人")
print(f"退库统计: {len(dfTK_ZJ)} 个项目")

# ================== 专业图表数据准备（固定1-52周，横坐标全部显示） ==================
all_weeks = [f"CW{i:02d}" for i in range(1, 53)]
print(f"主图表将显示周数范围: {all_weeks[0]} ~ {all_weeks[-1]}")

week_data = {w: {'2025参考': 0.0, '领用': 0.0, '隔离': 0.0, '每周需复用': 0.0} for w in all_weeks}
for _, row in dfKC.iterrows():
    week_raw = str(row.get('周数', '')).strip().upper()
    if week_raw in week_data:
        for col in ['2025参考', '领用', '隔离', '每周需复用']:
            val = row.get(col, 0)
            try:
                week_data[week_raw][col] = float(val) if not pd.isna(val) else 0.0
            except:
                week_data[week_raw][col] = 0.0

weeks = all_weeks
data_2025 = [week_data[w]['2025参考'] for w in weeks]
data_ly = [week_data[w]['领用'] for w in weeks]
data_gl = [week_data[w]['隔离'] for w in weeks]
data_fy = [week_data[w]['每周需复用'] for w in weeks]

print(f"数据提取完成，共 {len(weeks)} 周，最后10周领用数据预览: {data_ly[-10:]}")

# 横向条形图数据准备
TOP_N = 15
dfXM_top = dfXM_ZJ.head(TOP_N).copy() if not dfXM_ZJ.empty else pd.DataFrame(columns=['项目名', '领用总价'])
dfRM_top = dfRM_ZJ.head(TOP_N).copy() if not dfRM_ZJ.empty else pd.DataFrame(columns=['隔离人', '隔离总价'])
dfTK_top = dfTK_ZJ.head(TOP_N).copy() if not dfTK_ZJ.empty else pd.DataFrame(columns=['项目名', '退库总价'])

project_names = dfXM_top['项目名'].tolist() if not dfXM_top.empty else []
project_ly_vals = dfXM_top['领用总价'].tolist() if not dfXM_top.empty else []
isolator_names = dfRM_top['隔离人'].tolist() if not dfRM_top.empty else []
isolator_vals = dfRM_top['隔离总价'].tolist() if not dfRM_top.empty else []
return_project_names = dfTK_top['项目名'].tolist() if not dfTK_top.empty else []
return_vals = dfTK_top['退库总价'].tolist() if not dfTK_top.empty else []

project_ly_vals = [float(x) for x in project_ly_vals]
isolator_vals = [float(x) for x in isolator_vals]
return_vals = [float(x) for x in return_vals]

project_names_json = json.dumps(project_names, ensure_ascii=False)
project_ly_json = json.dumps(project_ly_vals)
isolator_names_json = json.dumps(isolator_names, ensure_ascii=False)
isolator_vals_json = json.dumps(isolator_vals)
return_names_json = json.dumps(return_project_names, ensure_ascii=False)
return_vals_json = json.dumps(return_vals)

# ================== 生成专业库存图表（横坐标全部显示） ==================
professional_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>专业库存统计图表</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        /* 关键：A3横向布局 */
        @page {{
            size: A3 landscape;
            margin: 0.5cm;
        }}
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 5mm;
            background: #f5f5f5;
        }}
        .container {{
            width: 1200px;
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 10px;
            box-sizing: border-box;
            border-radius: 8px;
        }}
        canvas {{
            width: 1200px !important;
            height: auto !important;
            display: block;
            max-width: 1200px;
        }}
        @media print {{
            .container {{
                width: auto !important;
                max-width: none;
                margin: 0;
            }}
            canvas {{
                width: 1200px !important;
                height: auto !important;
                max-width: none;
            }}
        }}
        h1 {{
            text-align: center;
            font-size: 24px;
            margin: 10px 0;
        }}
        .legend {{
            display: flex;
            justify-content: center;
            gap: 20px;
            margin: 20px 0;
            flex-wrap: wrap;
        }}
        .legend-item {{
            display: flex;
            align-items: center;
        }}
        .legend-color {{
            width: 20px;
            height: 20px;
            margin-right: 8px;
            border-radius: 3px;
        }}
        .color-2025 {{ background: #A6D396; }}
        .color-ly {{ background: #20678A; }}
        .color-gl {{ background: #F26E29; }}
        .color-fy {{ background: #0F9ED5; }}
        .chart-container {{
            width: 1200px;
            height: 560px;
            margin: 20px auto;
        }}
        .horizontal-chart-container {{
            width: 1200px;
            height: 600px;
            margin: 30px auto;
        }}
        h2 {{
            margin-top: 40px;
            color: #156082;
            border-left: 5px solid #156082;
            padding-left: 15px;
            font-size: 20px;
        }}
        .data-info {{
            background: #f8f9fa;
            padding: 15px;
            margin: 20px 0;
            border-radius: 5px;
            font-size: 14px;
        }}
    </style>
</head>
<body>
<div class="container">
    <h1>专业库存统计图表</h1>
    <div class="legend">
        <div class="legend-item"><div class="legend-color color-2025"></div><span>2025参考 (区域图)</span></div>
        <div class="legend-item"><div class="legend-color color-ly"></div><span>领用 (柱形图-下)</span></div>
        <div class="legend-item"><div class="legend-color color-gl"></div><span>隔离 (柱形图-上)</span></div>
        <div class="legend-item"><div class="legend-color color-fy"></div><span>每周需复用 (线型图)</span></div>
    </div>
    <div class="chart-container"><canvas id="mainChart" width="1200" height="560" style="width:1200px;height:560px"></canvas></div>
    <h2>📊 项目 - 库存领用</h2>
    <div class="horizontal-chart-container"><canvas id="barChartProject" width="1200" height="600" style="width:1200px;height:600px"></canvas></div>
    <h2>👤 隔离人 - 库存隔离</h2>
    <div class="horizontal-chart-container"><canvas id="barChartIsolator" width="1200" height="600" style="width:1200px;height:600px"></canvas></div>
    <h2>📦 项目 - 退库</h2>
    <div class="horizontal-chart-container"><canvas id="barChartReturn" width="1200" height="600" style="width:1200px;height:600px"></canvas></div>
    <div class="data-info">
        <h3>图表说明</h3>
        <ul><li>区域图 (绿色)：2025年每周金额趋势</li><li>柱形图 (蓝色+橙色)：领用（下）+ 隔离（上）</li><li>线型图 (蓝色)：每周需复用金额</li><li>横向条形图：项目领用、隔离人隔离、项目退库 TOP15</li></ul>
        <p>数据更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>
</div>
<script>
    const weekData = {json.dumps(weeks)};
    const data2025 = {json.dumps(data_2025)};
    const dataLY = {json.dumps(data_ly)};
    const dataGL = {json.dumps(data_gl)};
    const dataFY = {json.dumps(data_fy)};
    const projectNames = {project_names_json};
    const projectValues = {project_ly_json};
    const isolatorNames = {isolator_names_json};
    const isolatorValues = {isolator_vals_json};
    const returnNames = {return_names_json};
    const returnValues = {return_vals_json};

    new Chart(document.getElementById('mainChart'), {{
        type:'bar', data:{{ labels:weekData, datasets:[
            {{ type:'line', label:'2025参考', data:data2025, backgroundColor:'rgba(166,211,150,0.3)', borderColor:'#A6D396', borderWidth:2, fill:true, tension:0.1, order:3 }},
            {{ type:'bar', label:'领用', data:dataLY, backgroundColor:'#20678A', stack:'stack0', order:2 }},
            {{ type:'bar', label:'隔离', data:dataGL, backgroundColor:'#F26E29', stack:'stack0', order:2 }},
            {{ type:'line', label:'每周需复用', data:dataFY, borderColor:'#0F9ED5', borderWidth:3, pointBackgroundColor:'#0F9ED5', pointBorderColor:'#fff', pointRadius:4, fill:false, tension:0.2, order:1 }}
        ] }},
        options:{{ responsive:false, maintainAspectRatio:false, devicePixelRatio:1, scales:{{ x:{{ stacked:true, title:{{ display:true, text:'周数' }}, ticks:{{ maxRotation:90, minRotation:45, autoSkip:false }} }}, y:{{ beginAtZero:true, title:{{ display:true, text:'金额 (元)' }}, ticks:{{ callback:v=>v.toLocaleString() }} }} }}, plugins:{{ tooltip:{{ callbacks:{{ label:ctx=>ctx.dataset.label+': '+ctx.parsed.y.toLocaleString() }} }} }} }}
    }});
    new Chart(document.getElementById('barChartProject'), {{ type:'bar', data:{{ labels:projectNames, datasets:[{{ label:'领用总价', data:projectValues, backgroundColor:'#156082', barThickness:24, maxBarThickness:24 }}] }}, options:{{ indexAxis:'y', responsive:false, maintainAspectRatio:false, devicePixelRatio:1, scales:{{ x:{{ title:{{ display:true, text:'领用总价 (元)' }}, ticks:{{ callback:v=>v.toLocaleString() }} }} }} }} }});
    new Chart(document.getElementById('barChartIsolator'), {{ type:'bar', data:{{ labels:isolatorNames, datasets:[{{ label:'隔离总价', data:isolatorValues, backgroundColor:'#156082', barThickness:24, maxBarThickness:24 }}] }}, options:{{ indexAxis:'y', responsive:false, maintainAspectRatio:false, devicePixelRatio:1, scales:{{ x:{{ title:{{ display:true, text:'隔离总价 (元)' }}, ticks:{{ callback:v=>v.toLocaleString() }} }} }} }} }});
    new Chart(document.getElementById('barChartReturn'), {{ type:'bar', data:{{ labels:returnNames, datasets:[{{ label:'退库总价', data:returnValues, backgroundColor:'#156082', barThickness:24, maxBarThickness:24 }}] }}, options:{{ indexAxis:'y', responsive:false, maintainAspectRatio:false, devicePixelRatio:1, scales:{{ x:{{ title:{{ display:true, text:'退库总价 (元)' }}, ticks:{{ callback:v=>v.toLocaleString() }} }} }} }} }});
</script>
</body>
</html>"""
professional_html_file = "C:/XSR/githubPage/yanfeng/库存数据/专业库存图表.html"
with open(professional_html_file, 'w', encoding='utf-8') as f:
    f.write(professional_html)
print(f"专业库存图表已生成: {professional_html_file}")

# ================== 将专业图表转换为PDF（横向纸张） ==================
def html_to_pdf(html_path, pdf_path):
    browser_paths = [
        "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
        "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    ]
    browser_exe = None
    for path in browser_paths:
        if os.path.exists(path):
            browser_exe = path
            break
    if not browser_exe:
        print(f"未找到 Edge 或 Chrome 浏览器，无法将 {html_path} 转换为 PDF。")
        return False

    cmd = [
        browser_exe, "--headless", "--disable-gpu",
        "--landscape",
        "--window-size=1366,768",
        "--force-device-scale-factor=1",
        "--virtual-time-budget=15000",
        f"--print-to-pdf={pdf_path}",
        html_path
    ]
    try:
        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60,
            encoding='utf-8',
            errors='ignore',
            startupinfo=startupinfo
        )
        if result.returncode == 0 and os.path.exists(pdf_path):
            print(f"PDF已生成: {pdf_path}")
            return True
        else:
            print(f"转换失败: {html_path} -> {pdf_path}, 错误: {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        print(f"转换超时（60秒）: {html_path}")
        return False
    except Exception as e:
        print(f"转换异常: {e}")
        return False

# 只转换专业图表为PDF
html_path = professional_html_file
pdf_path = html_path.replace('.html', '.pdf')
html_to_pdf(html_path, pdf_path)

# ================== 发送邮件（专业图表PDF作为附件） ==================
print("\n正在打开 Outlook 邮件窗口...")
attachment_file = pdf_path
try:
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "liang.cao@yanfeng.com"
    mail.Subject = "库存统计报告 - 专业图表"
    mail.HTMLBody = (
        "请查收附件中的库存统计图表 PDF 文件。<br><br>"
        "<a href=\"https://shidingtech.cn/yanfeng/%E5%BA%93%E5%AD%98%E6%95%B0%E6%8D%AE/%E4%B8%93%E4%B8%9A%E5%BA%93%E5%AD%98%E5%9B%BE%E8%A1%A8.html\">点击此处</a>，查看网页版图表"
    )
    mail.Attachments.Add(attachment_file)
    mail.Display()
    print("Outlook 邮件窗口已打开，请确认收件人和附件后点击发送。")
except ImportError:
    print("未安装 pywin32 库，请执行：pip install pywin32")
    print(f"请手动发送邮件，附件位置：{attachment_file}")
except Exception as e:
    print(f"打开 Outlook 失败：{e}")
    print(f"请手动发送邮件，附件位置：{attachment_file}")