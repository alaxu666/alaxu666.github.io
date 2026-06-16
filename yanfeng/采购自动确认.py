import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from datetime import datetime, timedelta


def is_driver_alive(driver):
    try:
        _ = driver.current_window_handle
        return True
    except WebDriverException:
        return False
    except Exception:
        return False


def Login(driver, max_retries=3):
    """
    自动登录延锋BPM系统，遇到超时或ERR_TIMED_OUT错误自动重新加载。
    """
    wait = WebDriverWait(driver, 20)
    url = "https://onebpm.yanfeng.com/wui/index.html#/main/portal/portal-1-1?menuIds=0,1&_key=a3r00x"

    def has_todo_page():
        try:
            driver.find_element(By.XPATH, "//span[@title='待办事宜']")
            return True
        except Exception:
            return False

    for attempt in range(1, max_retries + 1):
        try:
            print(f"\n===== 登录尝试 {attempt}/{max_retries} =====")
            driver.get(url)

            # 检查是否跳转到错误页（如 ERR_TIMED_OUT）
            try:
                body_text = driver.find_element(By.TAG_NAME, 'body').text
                if "ERR_TIMED_OUT" in body_text or "错误" in body_text:
                    print("⚠️ 检测到错误页面，重新加载...")
                    driver.refresh()
                    time.sleep(3)
                    continue
            except:
                pass  # 忽略获取body失败的情况

            if has_todo_page():
                print("✅ 登录成功，已进入待办事宜页面（自动检测）。")
                return
            
            # 1. 等待邮箱输入框出现
            try:
                email_present = wait.until(EC.presence_of_element_located((By.NAME, "loginfmt")))
                print("检测到邮箱输入框，等待 8 秒观察是否自动登录...")
            except TimeoutException:
                print("未检测到邮箱输入框，可能已登录或无需输入邮箱。")
                email_present = None

            if has_todo_page():
                print("✅ 登录成功，已进入待办事宜页面（自动检测）。")
                return
            
            # 2. 如果邮箱框出现，等待自动跳转
            if email_present:
                try:
                    WebDriverWait(driver, 8).until_not(
                        EC.presence_of_element_located((By.NAME, "loginfmt"))
                    )
                    print("邮箱输入框已自动消失，可能已完成 SSO 登录。")
                    email_present = False
                except TimeoutException:
                    print("邮箱输入框持续存在，需要手动输入邮箱。")
                    email_present = True

            if has_todo_page():
                print("✅ 登录成功，已进入待办事宜页面（自动检测）。")
                return
            
            # 3. 手动输入邮箱（如需要）
            if email_present is True:
                try:
                    email_input = wait.until(EC.element_to_be_clickable((By.NAME, "loginfmt")))
                    email_input.clear()
                    email_input.send_keys("shirong.xu@yanfeng.com")
                    print("已输入邮箱。")

                    submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
                    submit_btn.click()
                    print("已提交邮箱。")
                except TimeoutException:
                    print("⚠️ 邮箱输入或提交超时，重新加载...")
                    driver.refresh()
                    time.sleep(3)
                    continue

            if has_todo_page():
                print("✅ 登录成功，已进入待办事宜页面（自动检测）。")
                return
            
            # 4. 等待密码框
            try:
                time.sleep(2)
                pwd_input = wait.until(EC.presence_of_element_located((By.NAME, "passwd")))
                print("检测到密码输入框。")

                # 等待密码框可能因 SSO 自动消失
                try:
                    WebDriverWait(driver, 3).until_not(
                        EC.presence_of_element_located((By.NAME, "passwd"))
                    )
                    print("密码框已自动消失，可能已完成登录。")
                except TimeoutException:
                    # 密码框持续存在，输入密码
                    try:
                        pwd_input = driver.find_element(By.NAME, "passwd")
                        pwd_input.clear()
                        pwd_input.send_keys("2045O2O5-7")
                        print("已输入密码。")

                        submit_btn = driver.find_element(By.XPATH, "//input[@type='submit']")
                        submit_btn.click()
                        print("已提交密码。")
                    except TimeoutException:
                        print("⚠️ 密码输入或提交超时，重新加载...")
                        driver.refresh()
                        time.sleep(3)
                        continue
            except TimeoutException:
                print("未检测到密码输入框，可能已直接登录。")

            if has_todo_page():
                print("✅ 登录成功，已进入待办事宜页面（自动检测）。")
                return

            # 5. 最终确认登录成功
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//span[@title='待办事宜']")))
                print("✅ 登录成功，已进入待办事宜页面。")
                return  # 登录成功，退出函数
            except TimeoutException:
                print("⚠️ 未能确认进入待办事宜页面，重新加载...")
                driver.refresh()
                time.sleep(3)
                continue

        except TimeoutException:
            print(f"⚠️ 登录流程超时，重新加载并重试...")
            driver.refresh()
            time.sleep(3)
            continue

    print("❌ 已用尽所有重试次数，登录失败。")



def AutomaticSupplierSelection(driver):
    """
    在新打开的采购申请详情页中，自动判断选择最佳供应商，完成后点击“批准”按钮。
    """
    wait = WebDriverWait(driver, 20)
    SFTJ = {}          # ← 必须在此处初始化，供循环外使用

    # 等待详情页下拉框加载完成
    try:
        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "td.detail_1_3_21.etype_3 div.ant-select")
        ))
        print("  详情页已加载，下拉框就绪")
    except TimeoutException:
        print("  [警告] 下拉框加载超时，尝试继续...")

    df_SupplierInfo = pd.DataFrame(columns=[
        "序号", "供应商代码", "供应商名称", "是否推荐", "第一轮报价", "不含税最终报价"
    ])

    try:
        table = wait.until(EC.presence_of_element_located((By.ID, "oTable0")))
        print("  [自动选择供应商] 找到表格 oTable0")
    except TimeoutException:
        print("  [错误] 未能在20秒内找到表格 oTable0")
        return df_SupplierInfo

    rows = table.find_elements(By.XPATH, ".//tr[@data-rowindex]")
    gys = []
    for row in rows:
        try:
            gys.append(int(row.get_attribute("data-rowindex")))
        except:
            continue
    print(f"  [自动选择供应商] 找到 {len(gys)} 个供应商")

    for row_idx in gys:
        try:
            current_row = table.find_element(By.XPATH, f".//tr[@data-rowindex='{row_idx}']")
        except:
            continue

        def get_cell_text(class_name):
            try:
                td = current_row.find_element(By.CSS_SELECTOR, f"td.{class_name}")
                span = td.find_element(By.TAG_NAME, "span")
                return span.text.strip()
            except:
                return ""

        xuhao = get_cell_text("detail_1_3_1.etype_22")
        gys_code = get_cell_text("detail_1_3_2.etype_3")
        gys_name = get_cell_text("detail_1_3_3.etype_3")
        recommend = get_cell_text("detail_1_3_4.etype_3")
        first_price = get_cell_text("detail_1_3_5.etype_3")
        final_price = get_cell_text("detail_1_3_6.etype_3")
        SFTJ[row_idx] = [xuhao, gys_code, gys_name, recommend, first_price, final_price]

        # ----- 下拉框选择 -----
        def select_dropdown_option(td_class, option_text):
            try:
                driver.find_element(By.TAG_NAME, 'body').click()
                time.sleep(0.3)
                td = current_row.find_element(By.CSS_SELECTOR, f"td.{td_class}")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", td)
                dropdown = td.find_element(By.CSS_SELECTOR, "div.ant-select")
                dropdown.click()
                time.sleep(0.5)
                dropdown_container = wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, ".ant-select-dropdown:not(.ant-select-dropdown-hidden)")
                    )
                )
                option = dropdown_container.find_element(
                    By.XPATH,
                    f".//li[@unselectable='unselectable' and normalize-space()='{option_text}']"
                )
                option.click()
                time.sleep(0.3)
                print(f"    已选择下拉选项: {option_text}")
            except Exception as e:
                print(f"    下拉选项失败 ({td_class}): {repr(e)}")

        select_dropdown_option("detail_1_3_21.etype_3", "是")
        if final_price == "0.00":
            select_dropdown_option("detail_1_3_22.etype_3", "否")
        else:
            select_dropdown_option("detail_1_3_22.etype_3", "是")

        # ----- 输入框填充 -----
        def fill_input(td_class, value):
            try:
                td = current_row.find_element(By.CSS_SELECTOR, f"td.{td_class}")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", td)
                input_elem = td.find_element(By.TAG_NAME, "input")
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"td.{td_class} input")))
                input_elem.clear()
                input_elem.send_keys(value)
                print(f"    已填充 {td_class}: {value}")
            except Exception as e:
                print(f"    输入框填充失败: {repr(e)}")

        input_value = "一般" if recommend == "N" else "良好"
        fill_input("detail_1_3_23.etype_3", input_value)
        fill_input("detail_1_3_24.etype_3", input_value)

    # ========== 点击“批准”按钮 ==========
    try:
        print("  正在查找“批准”按钮...")
        approve_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='批准' or contains(@title,'批准')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", approve_btn)
        time.sleep(0.5)
        approve_btn.click()
        print("  ✅ 已点击“批准”按钮。")
        time.sleep(2)
    except TimeoutException:
        print("  未找到批准按钮，可能无需审批。")
    except WebDriverException as e:
        print(f"  点击批准按钮时 Selenium 异常: {repr(e)}")
    except Exception as e:
        print(f"  点击批准按钮异常: {repr(e)}")

    # ----- 构建结果 DataFrame -----
    rows_data = []
    for idx in sorted(gys):
        if idx in SFTJ and len(SFTJ[idx]) >= 6:
            rows_data.append({
                "序号": SFTJ[idx][0],
                "供应商代码": SFTJ[idx][1],
                "供应商名称": SFTJ[idx][2],
                "是否推荐": SFTJ[idx][3],
                "第一轮报价": SFTJ[idx][4],
                "不含税最终报价": SFTJ[idx][5]
            })
    df_SupplierInfo = pd.DataFrame(rows_data)
    print("  供应商选择结果：")
    print(df_SupplierInfo.to_string(index=False))
    return df_SupplierInfo


def PurchaseApplicationApproval(driver):
    """
    自动寻找BPM中未审批的采购申请，排除徐时荣后，逐个打开并执行供应商自动选择与审批。
    """
    wait = WebDriverWait(driver, 15)
    df = pd.DataFrame(columns=["流程标题", "流程名称", "创建人", "创建日期", "未操作者"])

    # 1. 点击“待办事宜”
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@title='待办事宜']"))).click()
        print("已点击“待办事宜”")
        time.sleep(2)
    except TimeoutException:
        print("错误：找不到“待办事宜”菜单")
        return

    # 2. 设置每页100条
    try:
        size_selector = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.ant-select-selection.ant-select-selection--single")))
        size_selector.click()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//li[contains(text(), '100条/页')]"))).click()
        print("已切换为 100条/页")
        time.sleep(2)
    except TimeoutException:
        print("警告：未能设置每页100条")

    # 3. 提取表格数据
    try:
        tbody = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.ant-table-tbody")))
    except TimeoutException:
        print("错误：未找到表格主体")
        return

    rows = tbody.find_elements(By.TAG_NAME, "tr")
    print(f"找到 {len(rows)} 行待办流程")

    for tr in rows:
        cells = tr.find_elements(By.CSS_SELECTOR, "td.autoWrap")
        if len(cells) < 6:
            continue
        col2 = cells[1].get_attribute("stsdata") or ""
        def extract_text(td):
            spans = td.find_elements(By.TAG_NAME, "span")
            return spans[-1].text.strip() if spans else td.text.strip()
        col3 = extract_text(cells[2])
        col4 = extract_text(cells[3])
        col5 = extract_text(cells[4])
        try:
            spans_6 = cells[5].find_elements(By.TAG_NAME, "span")
            col6 = spans_6[2].text.strip() if len(spans_6) >= 3 else ""
        except:
            col6 = ""
        df.loc[len(df)] = [col2, col3, col4, col5, col6]

    # 4. 删除“徐时荣”
    df = df[df["创建人"].str.strip() != "徐时荣"]
    print("\n" + "="*60)
    print("筛选后的采购申请（已排除徐时荣）：")
    print("="*60)
    if df.empty:
        print("没有符合条件的记录。")
        return
    else:
        print(df.to_string(index=False))
    print("="*60)

    # 5. 遍历流程，点击并处理
    for idx, row in df.iterrows():
        title_value = row["流程标题"]
        print(f"\n>>> 处理流程: {title_value}")

        if not is_driver_alive(driver):
            print("  Selenium 会话已失效，停止处理流程。")
            return

        # 记录当前窗口句柄和窗口总数
        main_window = driver.current_window_handle
        original_window_count = len(driver.window_handles)
        prev_url = driver.current_url

        # ---------- 查找并点击流程标题 ----------
        clicked = False
        click_element = None

        # 处理标题中的特殊字符（单引号转义）
        safe_title = title_value.replace("'", "\\'")

        # 尝试1：a 标签文本精确匹配
        try:
            link = driver.find_element(By.XPATH, f"//a[normalize-space()='{safe_title}']")
            click_element = link
            clicked = True
            print("  点击方式：a 标签文本精确匹配")
        except:
            pass

        # 尝试2：a 标签文本包含匹配
        if not clicked:
            try:
                link = driver.find_element(By.XPATH, f"//a[contains(text(), '{safe_title}')]")
                click_element = link
                clicked = True
                print("  点击方式：a 标签文本包含匹配")
            except:
                pass

        # 尝试3：td[stsdata] 内的第一个 a 标签
        if not clicked:
            try:
                td = driver.find_element(By.CSS_SELECTOR, f"td[stsdata*='{title_value}']")
                link = td.find_element(By.TAG_NAME, "a")
                click_element = link
                clicked = True
                print("  点击方式：td 内的 a 标签")
            except:
                pass

        # 尝试4：兜底点击 td 本身
        if not clicked:
            try:
                td = driver.find_element(By.CSS_SELECTOR, f"td[stsdata*='{title_value}']")
                click_element = td
                clicked = True
                print("  点击方式：td stsdata 属性（兜底）")
            except:
                pass

        if not clicked:
            print("  无法找到任何可点击元素，跳过")
            continue

        # 使用 JS 点击（确保触发事件）
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", click_element)
        driver.execute_script("arguments[0].click();", click_element)
        time.sleep(1)

        # ---------- 等待新窗口或页面跳转 ----------
        new_window = False
        try:
            WebDriverWait(driver, 30).until(
                lambda d: len(d.window_handles) > original_window_count or d.current_url != prev_url
            )
            if not is_driver_alive(driver):
                print("  Selenium 会话已失效，跳过该流程")
                return
            if len(driver.window_handles) > original_window_count:
                new_window = True
                # 切换到新窗口（最后一个新增窗口）
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        break
                print("  已跳转到详情页（新窗口）")
            else:
                print("  当前页面已跳转到详情页（同页)")
        except TimeoutException:
            print("  等待跳转超时，跳过该流程")
            continue
        except WebDriverException:
            print("  Selenium 会话已失效，跳过该流程")
            return

        # ---------- 执行自动审批 ----------
        AutomaticSupplierSelection(driver)

        # ---------- 等待审批结果，最多重试3次 ----------
        if new_window:
            # 新窗口模式：审批成功后窗口会自动关闭
            max_retries = 3
            approved = False
            for attempt in range(1, max_retries + 1):
                print(f"  等待审批结果... (第 {attempt}/{max_retries} 次，最多等10秒)")
                # 每隔2秒检查一次窗口是否已关闭，累计等10秒
                for i in range(5):
                    time.sleep(2)
                    # 检查审批窗口是否已关闭（窗口数量减少）
                    if len(driver.window_handles) <= original_window_count:
                        approved = True
                        print("  ✅ 审批窗口已关闭，审批成功！")
                        break
                if approved:
                    break
                # 检查当前是否还在审批页面（可能窗口没关但页面变了）
                try:
                    current_handles = driver.window_handles
                    if len(current_handles) <= original_window_count:
                        approved = True
                        print("  ✅ 审批窗口已关闭，审批成功！")
                        break
                except:
                    pass
                if attempt < max_retries:
                    print(f"  第 {attempt} 次等待超时，继续等待...")
                else:
                    # 3次都没等到窗口关闭，刷新页面重新审批
                    print(f"  ⚠️ 等待 {max_retries} 次后仍未检测到审批成功，刷新页面重新审批...")
                    try:
                        driver.refresh()
                        time.sleep(3)
                        # 重新执行审批操作
                        AutomaticSupplierSelection(driver)
                        # 再等一次审批结果
                        for i in range(5):
                            time.sleep(2)
                            if len(driver.window_handles) <= original_window_count:
                                approved = True
                                print("  ✅ 刷新后审批成功！")
                                break
                        if not approved:
                            print("  ❌ 刷新后仍未审批成功，跳过此流程")
                    except Exception as e:
                        print(f"  ❌ 刷新重试失败: {repr(e)}")

            # 切换回主窗口
            try:
                if driver.window_handles:
                    driver.switch_to.window(main_window)
            except WebDriverException as e:
                print(f"  无法切换回主窗口：{repr(e)}")
                return
        else:
            # 同页跳转模式：审批成功后页面会跳转或出现成功提示
            max_retries = 3
            approved = False
            for attempt in range(1, max_retries + 1):
                print(f"  等待审批结果... (第 {attempt}/{max_retries} 次，最多等10秒)")
                time.sleep(10)
                # 检查是否已返回列表页（URL变化或出现列表元素）
                try:
                    driver.find_element(By.CSS_SELECTOR, "tbody.ant-table-tbody")
                    current_url = driver.current_url
                    if current_url != prev_url or "portal" in current_url.lower():
                        approved = True
                        print("  ✅ 已返回列表页，审批成功！")
                        break
                except:
                    pass
                if attempt < max_retries:
                    print(f"  第 {attempt} 次等待超时，继续等待...")
                else:
                    # 3次都没等到，刷新页面重新审批
                    print(f"  ⚠️ 等待 {max_retries} 次后仍未检测到审批成功，刷新页面重新审批...")
                    try:
                        driver.back()
                        time.sleep(2)
                        driver.refresh()
                        time.sleep(3)
                        # 重新点击流程并审批
                        try:
                            driver.find_element(By.XPATH, f"//a[contains(text(), '{safe_title}')]").click()
                            time.sleep(3)
                            AutomaticSupplierSelection(driver)
                            for i in range(5):
                                time.sleep(2)
                                driver.find_element(By.CSS_SELECTOR, "tbody.ant-table-tbody")
                                if driver.current_url != prev_url:
                                    approved = True
                                    print("  ✅ 刷新后审批成功！")
                                    break
                            if not approved:
                                print("  ❌ 刷新后仍未审批成功，跳过此流程")
                        except Exception as e:
                            print(f"  ❌ 刷新重试失败: {repr(e)}")
                    except Exception as e:
                        print(f"  ❌ 返回并刷新失败: {repr(e)}")

            # 确保回到列表页
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.ant-table-tbody"))
                )
            except:
                try:
                    driver.back()
                    time.sleep(2)
                except:
                    pass

        # 完全回到主窗口后，稍作停顿再处理下一个
        time.sleep(2)

    print("\n所有流程处理完毕。")

        # (前面的流程循环代码保持不变……)

    # ========== 循环结束后，更新采购申请清单.csv ==========
    print("\n正在更新采购申请清单...")
    csv_path = r"C:\XSR\githubPage\yanfeng\采购自动确认\采购申请清单.csv"

    # 计算上周一的日期
    today = datetime.now().date()
    # 上周一 = 今天 - 7天 - (今天星期几 - 1) ？ 更稳健：先找本周一，再减7天
    this_monday = today - timedelta(days=today.weekday())  # 本周一
    last_monday = this_monday - timedelta(days=7)          # 上周一

    try:
        # 尝试读取现有 CSV
        old_df = pd.read_csv(csv_path)
        # 将创建日期列转换为 datetime
        old_df['创建日期'] = pd.to_datetime(old_df['创建日期']).dt.date
        # 保留创建日期 >= 上周一的记录
        old_df = old_df[old_df['创建日期'] >= last_monday]
        print(f"原CSV保留上周一({last_monday})及以后的数据，剩余 {len(old_df)} 行。")
    except FileNotFoundError:
        old_df = pd.DataFrame(columns=df.columns)
        print("未找到原CSV文件，将新建。")

    # 本次新数据（df_PurchaseRequisition 已排除徐时荣）
    new_candidates = df.copy()
    # 标准化创建日期列类型
    new_candidates['创建日期'] = pd.to_datetime(new_candidates['创建日期']).dt.date

    # 合并数据并去重（基于流程标题和创建日期）
    combined = pd.concat([old_df, new_candidates], ignore_index=True)
    combined.drop_duplicates(subset=['流程标题', '创建日期'], keep='last', inplace=True)

    # 覆盖保存回 CSV
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    combined.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"采购申请清单已更新，共 {len(combined)} 条记录。")



def main():
    from selenium.webdriver.edge.options import Options   # Edge 专用选项
    from selenium.webdriver.edge.service import Service  # 指定驱动日志路径
    options = Options()
    # 可选：保持登录状态（第一次运行时需注释掉，等成功登录后再启用）
    # options.add_argument("--user-data-dir=C:/Users/uxuxs004/AppData/Local/Microsoft/Edge/User Data")
    
    # 关闭浏览器启动时的冗余日志输出
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--log-level=3')
    options.add_argument('--disable-logging')

    # 使用 Edge 驱动（确保 msedgedriver.exe 在 PATH 中或已指定路径）
    service = Service(log_path=os.devnull)
    driver = webdriver.Edge(service=service, options=options)

    try:
        Login(driver)
        time.sleep(3)
        PurchaseApplicationApproval(driver)
        print("\n操作完成，浏览器将在1秒后关闭...")
        time.sleep(1)
    except Exception as e:
        print(f"程序出错: {e}")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()