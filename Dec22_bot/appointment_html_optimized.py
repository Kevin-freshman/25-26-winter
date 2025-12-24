from playwright.sync_api import sync_playwright
import pandas as pd
import time
import os
import re
from datetime import datetime

# ---------------- 配置信息 ----------------
URL = "https://emsvip.linkedlife.cn/"
COMPANY = "xm-lf"
USERNAME = "前台"
PASSWORD = "123"
EXCEL_PATH = "appointments.xlsx"

# ---------------- 工具函数 ----------------

def parse_date_time(raw_time_str):
    """解析时间字符串"""
    if not raw_time_str:
        return "", ""
    try:
        parts = raw_time_str.split("-")[0].strip()
        dt = datetime.strptime(parts, "%Y/%m/%d %H:%M")
        date_str = dt.strftime("%m月%d日").lstrip("0").replace("月0", "月")
        time_str = dt.strftime("%H:%M")
        return date_str, time_str
    except Exception as e:
        return raw_time_str, ""

def get_next_index():
    """获取 Excel 下一个序号"""
    if not os.path.exists(EXCEL_PATH): return 1
    try:
        df = pd.read_excel(EXCEL_PATH)
        if "序号" in df.columns and not df.empty:
            return int(df["序号"].max()) + 1
        return 1
    except:
        return 1

def save_to_excel(raw_data: dict):
    """保存到 Excel"""
    date_str, time_str = parse_date_time(raw_data.get("预约时间", ""))
    
    # 构建数据行
    new_row = {
        "序号": get_next_index(),
        "上门日期": date_str,
        "具体时间": time_str,
        "顾客姓名": raw_data.get("姓名", ""),
        "病历号/会员卡号": raw_data.get("会员号", ""),
        "来源渠道": raw_data.get("客户来源", "")
    }

    df_new = pd.DataFrame([new_row])

    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        # 确保列类型一致，防止报错
        if "病历号/会员卡号" in df_old.columns:
            df_old["病历号/会员卡号"] = df_old["病历号/会员卡号"].astype(str)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new
        cols = ["序号", "上门日期", "具体时间", "顾客姓名", "病历号/会员卡号", "来源渠道"]
        # 简单的列存在性检查
        valid_cols = [c for c in cols if c in df.columns]
        df = df[valid_cols]
    
    df.to_excel(EXCEL_PATH, index=False)
    print(f"✅ [写入成功] 序号: {new_row['序号']} | 姓名: {new_row['顾客姓名']}")

def already_exists(member_id: str, date_check: str) -> bool:
    """防止重复录入"""
    if not os.path.exists(EXCEL_PATH): return False
    try:
        df = pd.read_excel(EXCEL_PATH)
        if "病历号/会员卡号" in df.columns and "上门日期" in df.columns:
            # 统一转为字符串比较，防止 excel 里的数字和读取的字符串不匹配
            df["病历号/会员卡号"] = df["病历号/会员卡号"].astype(str)
            filtered = df[df["病历号/会员卡号"] == str(member_id)]
            if not filtered.empty:
                if date_check in filtered["上门日期"].values:
                    return True
    except Exception as e:
        print(f"查重读取失败: {e}")
    return False

# ---------------- 页面行为 ----------------

def login(page):
    print("正在登录...")
    try:
        page.goto(URL, wait_until="domcontentloaded", timeout=30000)
    except Exception as e:
        print(f"⚠️ 首次连接超时，正在重试... ({e})")
        page.goto(URL, wait_until="domcontentloaded", timeout=30000)

    try:
        page.locator("input[type='text']").nth(0).fill(COMPANY)
        page.locator("input[type='text']").nth(1).fill(USERNAME)
        page.locator("input[type='password']").fill(PASSWORD)
        page.get_by_role("button", name="登 录").click()
        
        # 等待左侧菜单加载
        page.wait_for_selector("text=预约", timeout=30000)
        print("登录成功。")
    except Exception as e:
        print(f"登录过程出错: {e}")

def goto_appointment_center(page):
    print("正在跳转到预约中心...")
    try:
        # 1. 点击一级菜单 "预约"
        menu_btn = page.locator("li").filter(has_text="预约").first
        menu_btn.click()
        time.sleep(1)
        
        # 2. 点击二级菜单 "预约中心"
        sub_menu_btn = page.locator("li").filter(has_text="预约中心").first
        if sub_menu_btn.is_visible():
            sub_menu_btn.click()
        else:
            sub_menu_btn.click(force=True)
        
        time.sleep(5)

        # 3. 切换视图
        view_tab = page.locator("div").filter(has_text="预约视图").last
        if view_tab.is_visible():
            print("正在切换到【预约视图】(日历模式)...")
            view_tab.click()
        
        # 4. 等待加载
        page.wait_for_selector(".appointment-block-container, .fc-view-container", timeout=15000)
        print("日历视图加载完成。")
        
    except Exception as e:
        print(f"跳转导航警告: {e}")
        print("尝试继续执行...")

def extract_detail_from_modal(page) -> dict:
    """提取弹窗数据"""
    data = {}
    # 等待弹窗内容
    page.locator(".ant-modal-content").first.wait_for(timeout=5000)
    modal_text = page.locator(".ant-modal-body").inner_text()
    
    try:
        # 头部信息
        header = page.locator(".header-info").inner_text()
        # 提取会员号
        id_match = re.search(r"\d{6,}", header)
        data["会员号"] = id_match.group(0) if id_match else ""
        
        # 提取姓名
        lines = header.split('\n')
        name_candidate = lines[0].strip()
        data["姓名"] = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]', '', name_candidate)
    except:
        data["姓名"] = "未知"
        data["会员号"] = ""

    # 提取键值对
    for line in modal_text.split('\n'):
        if "：" in line:
            parts = line.split("：", 1)
            if len(parts) == 2:
                data[parts[0].strip()] = parts[1].strip()
            
    return data

def process_appointments(page):
    blue_card_selector = "div.appointment-block-container.blue"
    
    print("\n>>> 开始执行滚动扫描 <<<")
    
    # ---------------- 新增：滚动加载逻辑 ----------------
    # 尝试在页面中心位置进行滚轮滚动，触发懒加载
    # 很多日历是在 div 内部滚动的，所以我们将鼠标移动到屏幕中间
    try:
        page.mouse.move(x=500, y=500)
        
        # 循环滚动几次，确保底部卡片加载出来
        # range(5) 表示滚动 5 次，每次滚动 2000 像素，你可以根据数据量调整次数
        for i in range(1, 6):
            print(f"正在向下滚动 ({i}/5)...")
            page.mouse.wheel(delta_x=0, delta_y=2000)
            time.sleep(2) # 每次滚动后等待 2 秒让数据渲染
            
        print("滚动完成，等待 DOM 稳定...")
        time.sleep(2)
        
    except Exception as e:
        print(f"滚动过程出现小问题（不影响主流程）: {e}")

    # ---------------- 扫描与处理 ----------------
    
    print("正在扫描【蓝色/已到店】卡片...")
    
    # 再次等待，确保滚动后的元素已就位
    try:
        page.wait_for_selector(blue_card_selector, timeout=5000)
    except:
        pass # 超时也没关系，依靠下面的 count 判断

    cards = page.locator(blue_card_selector)
    count = cards.count()
    print(f"--> 发现 {count} 个蓝色卡片待处理。")

    if count == 0:
        print("⚠️ 依然未检测到蓝色卡片。请检查：\n1. 页面上是否真的有蓝色卡片？\n2. 是否需要手动筛选日期？")
        return

    # 遍历处理
    for i in range(count):
        # 注意：在 Playwright 中，当你操作完第一个元素，页面DOM可能刷新，
        # 所以每次循环最好重新获取一下列表的引用，或者使用 .nth(i) 这种动态定位
        
        card = cards.nth(i)
        
        # 确保卡片在可视区域（Playwright 点击前会自动滚动，但显示出来更安全）
        try:
            card.scroll_into_view_if_needed()
        except:
            pass

        try:
            # 获取名字日志
            try:
                card_name = card.locator(".user-name").inner_text().strip()
            except:
                card_name = f"第 {i+1} 个卡片"

            print(f"[{i+1}/{count}] 处理: {card_name}")
            
            # 点击卡片
            card.click()
            
            # 提取详情
            raw_data = extract_detail_from_modal(page)
            
            # 查重逻辑
            date_check, _ = parse_date_time(raw_data.get("预约时间", ""))
            
            if already_exists(raw_data.get("会员号"), date_check):
                print(f"   -> 跳过 (Excel中已存在)")
            else:
                save_to_excel(raw_data)
            
            # 关闭弹窗
            page.keyboard.press("Escape")
            time.sleep(1) # 等待弹窗完全关闭
            
        except Exception as e:
            print(f"   -> 处理出错: {e}")
            # 出错后尝试按 ESC 复位，防止阻挡下一个
            page.keyboard.press("Escape")
            time.sleep(1)

# ---------------- 主程序 ----------------

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False, 
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()
        page.set_default_timeout(30000)

        login(page)
        goto_appointment_center(page)
        process_appointments(page)
        
        print("\n所有任务完成，程序将在 5 秒后关闭...")
        time.sleep(5)
        browser.close()

if __name__ == "__main__":
    main()