# -*- coding: utf-8 -*-
"""
配置文件 - 根据实际环境修改以下参数
"""

# ---------- 路径配置 ----------
# Excel文件绝对路径（E列为Article ID）
EXCEL_PATH = r"C:\Users\E1547105\Desktop\China-Guardian\modle\KBAdemo.xlsx"

# PDF下载保存目录
DOWNLOAD_DIR = r"C:\Users\E1547105\Desktop\China-Guardian\downloads"

# Chrome浏览器驱动路径（如果已加入PATH，可设为None）
CHROME_DRIVER_PATH = None  # 例如：r"D:\chromedriver.exe"

# 目标网页
TARGET_URL = "https://sms.emerson.com/kba"

# ---------- 等待时间（秒） ----------
WAIT_AFTER_OPEN = 8          # 打开网页后等待
WAIT_AFTER_USERNAME = 7      # 输入用户名后等待
WAIT_AFTER_PASSWORD = 7       # 输入密码后等待
WAIT_AFTER_SEARCH = 10        # 搜索后等待结果加载
WAIT_AFTER_OPEN_DETAIL = 7    # 打开详情页后等待
WAIT_AFTER_DOWNLOAD = 6       # 点击下载后等待文件完成

# ---------- 预定义元素定位器（CSS选择器或XPath）----------
# 注意：这些仅为示例，必须使用浏览器开发者工具替换为实际值
# 如果自动定位失败，程序会弹出交互窗口让用户手动指定并自动保存到selectors.json
LOGIN_USERNAME_SELECTOR = "#username"        # 用户名输入框
LOGIN_PASSWORD_SELECTOR = "#password"        # 密码输入框
LOGIN_BUTTON_SELECTOR = "button[type='submit']"  # 登录按钮

SEARCH_INPUT_SELECTOR = "input[placeholder='All kba by numbers']"  # 搜索输入框
FIRST_RESULT_SELECTOR = "table tbody tr:first-child"  # 第一条搜索结果行
DOWNLOAD_ICON_SELECTOR = ".download-icon"  # 下载图标（右上角）

# ---------- 其他配置 ----------
# 文件下载超时（秒）
DOWNLOAD_TIMEOUT = 30
# 手动定位时鼠标点击超时（秒）
MANUAL_LOCATE_TIMEOUT = 10