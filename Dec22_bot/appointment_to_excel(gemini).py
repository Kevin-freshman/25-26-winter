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
    """
    解析时间字符串，例如 "2025/12/22 10:30 - 11:15"
    返回: ("12月22日", "10:30")
    """
    if not raw_time_str:
        return "", ""
    try:
        # 取第一部分 "2025/12/22 10:30"
        parts = raw_time_str.split("-")[0].strip()
        dt = datetime.strptime(parts, "%Y/%m/%d %H:%M")
        
        # 格式化为 Excel 需要的格式
        date_str = dt.strftime("%m月%d日").lstrip("0").replace("月0", "月") # 12月22日
        time_str = dt.strftime("%H:%M") # 10:30
        return date_str, time_str
    except Exception as e:
        print(f"时间解析失败: {raw_time_str}, 错误: {e}")
        return raw_time_str, ""

def get_next_index():
    """获取 Excel 下一个序号"""
    if not os.path.exists(EXCEL_PATH):
        return 1
    try:
        df = pd.read_excel(EXCEL_PATH)
        if "序号" in df.columns and not df.empty:
            # 获取最大序号并 +1
            return int(df["序号"].max()) + 1
        return 1
    except:
        return 1

def save_to_excel(raw_data: dict):
    """
    将抓取的数据转换为用户指定的 Excel 格式
    目标列: [序号, 上门日期, 具体时间, 顾客姓名, 病历号/会员卡号, 来源渠道]
    """
    # 1. 解析日期和时间
    date_str, time_str = parse_date_time(raw_data.get("预约时间", ""))
    
    # 2. 构建新的一行数据
    new_row = {
        "序号": get_next_index(),
        "上门日期": date_str,
        "具体时间": time_str,
        "顾客姓名": raw_data.get("姓名", ""),
        "病历号/会员卡号": raw_data.get("会员号", ""),
        "来源渠道": raw_data.get("客户来源", "")
    }

    df_new = pd.DataFrame([new_row])

    # 3. 读取或创建 Excel
    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        # 确保卡号是字符串，防止变成科学计数法
        if "病历号/会员卡号" in df_old.columns:
            df_old["病历号/会员卡号"] = df_old["病历号/会员卡号"].astype(str)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new
        # 设置列顺序
        cols = ["序号", "上门日期", "具体时间", "顾客姓名", "病历号/会员卡号", "来源渠道"]
        df = df[cols]
    
    df.to_excel(EXCEL_PATH, index=False)
    print(f"✅ [写入成功] 序号: {new_row['序号']} | 姓名: {new_row['顾客姓名']}")

def is_blue_card(rgb_string: str) -> bool:
    """颜色判断逻辑：蓝色才处理"""
    if not rgb_string: return False
    match = re.search(r"rgb\((\d+),\s*(\d+),\s*(\d+)\)", rgb_string)
    if not match: return False
    r, g, b = map(int, match.groups())
    # 蓝色判定：B > R 且 B > 200 (排除黄色/米色)
    if b > r and b > 200: return True
    if b > 220 and r < 230: return True
    return False

def already_exists(member_id: str, date_check: str) -> bool:
    """防止重复录入"""
    if not os.path.exists(EXCEL_PATH): return False
    df = pd.read_excel(EXCEL_PATH)
    # 简单去重逻辑：如果同一个会员号，且日期（上门日期）也一样，则视为重复
    # 注意：这里 date_check 传入的是 "12月22日" 这种格式
    if "病历号/会员卡号" in df.columns and "上门日期" in df.columns:
        # 筛选会员号
        filtered = df[df["病历号/会员卡号"].astype(str) == str(member_id)]
        if not filtered.empty:
            # 检查日期是否也存在
            if date_check in filtered["上门日期"].values:
                return True
    return False

# ---------------- 页面行为 ----------------

def login(page):
    print("正在登录...")
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    
    page.locator("input[type='text']").nth(0).fill(COMPANY)
    page.locator("input[type='text']").nth(1).fill(USERNAME)
    page.locator("input[type='password']").fill(PASSWORD)

    page.get_by_role("button", name="登 录").click()
    
    # --- 修复点 1: 不再等待特定 URL，而是等待菜单栏出现 ---
    print("等待跳转到首页...")
    try:
        # 等待左侧菜单的“预约”二字出现，最长等 30 秒
        page.wait_for_selector("text=预约", timeout=30000)
        print("登录成功，检测到菜单栏。")
    except:
        print("⚠️ 登录后未检测到菜单，可能需要人工干预或网络太慢。")

def goto_appointment_center(page):
    print("正在跳转到预约中心...")
    
    # --- 修复点 2: 增强点击稳定性 ---
    try:
        # 1. 点击一级菜单 "预约"
        # 使用模糊匹配，防止因为图标或空格导致匹配失败
        menu_btn = page.locator("li").filter(has_text="预约").first
        menu_btn.click()
        time.sleep(1) # 等待子菜单动画展开
        
        # 2. 点击二级菜单 "预约中心"
        sub_menu_btn = page.locator("li").filter(has_text="预约中心").first
        
        # 如果不可见，强制点击
        if sub_menu_btn.is_visible():
            sub_menu_btn.click()
        else:
            print("尝试强制点击预约中心...")
            sub_menu_btn.click(force=True)

        # 3. 等待日历视图加载
        page.wait_for_selector(".fc-view-container", timeout=20000)
        time.sleep(3) # 额外等待数据渲染
        print("已进入预约视图。")
        
    except Exception as e:
        print(f"跳转导航失败: {e}")
        # 截图保存现场，方便调试
        page.screenshot(path="error_nav.png")

def extract_detail_from_modal(page) -> dict:
    """提取数据"""
    data = {}
    
    # 等待弹窗
    page.locator(".ant-modal-content").first.wait_for(timeout=5000)
    modal_text = page.locator(".ant-modal-body").inner_text()
    
    # 提取姓名 (尝试从头部获取)
    try:
        header = page.locator(".header-info").inner_text()
        # 假设第一行是名字，或者是除去数字的部分
        lines = header.split('\n')
        name_candidate = lines[0].strip()
        # 简单清洗，去掉性别符号等
        data["姓名"] = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]', '', name_candidate)
        
        # 提取会员号
        id_match = re.search(r"\d{6,}", header)
        data["会员号"] = id_match.group(0) if id_match else ""
    except:
        data["姓名"] = "未知"
        data["会员号"] = ""

    # 提取列表项
    for line in modal_text.split('\n'):
        if "：" in line:
            key, val = line.split("：", 1)
            data[key.strip()] = val.strip()
            
    return data

def process_appointments(page):
    # 查找日历卡片
    card_selector = "a.fc-day-grid-event"
    try:
        page.wait_for_selector(card_selector, timeout=10000)
    except:
        print("当前视图无预约。")
        return

    cards = page.locator(card_selector)
    count = cards.count()
    print(f"检测到 {count} 个预约卡片。")

    for i in range(count):
        card = cards.nth(i)
        
        # 颜色判断
        bg_color = card.evaluate("el => window.getComputedStyle(el).backgroundColor")
        if not is_blue_card(bg_color):
            continue
        
        # 点击处理
        try:
            print(f"处理第 {i+1} 个卡片...")
            card.click()
            
            # 提取
            raw_data = extract_detail_from_modal(page)
            
            # 预处理日期以便查重
            date_check, _ = parse_date_time(raw_data.get("预约时间", ""))
            
            # 查重与保存
            if already_exists(raw_data.get("会员号"), date_check):
                print(f"   -> 跳过: {raw_data.get('姓名')} (已存在)")
            else:
                save_to_excel(raw_data)
            
            # 关闭弹窗
            page.keyboard.press("Escape")
            time.sleep(0.5)
            
        except Exception as e:
            print(f"   -> 处理出错: {e}")
            page.keyboard.press("Escape")

# ---------------- 主程序 ----------------

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        login(page)
        goto_appointment_center(page)
        process_appointments(page)
        
        print("任务完成，3秒后退出...")
        time.sleep(3)
        browser.close()

if __name__ == "__main__":
    main()