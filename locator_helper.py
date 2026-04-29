# -*- coding: utf-8 -*-
"""
智能元素定位器 - 自动定位失败时提示用户手动点击元素，并记录选择器到JSON文件
"""
import os
import json
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pyautogui
import tkinter as tk
from tkinter import simpledialog, messagebox

# 选择器存储文件
SELECTORS_FILE = "selectors.json"

class LocatorHelper:
    def __init__(self, driver):
        self.driver = driver
        self.selectors = self._load_selectors()

    def _load_selectors(self):
        """加载已保存的选择器（JSON文件）"""
        if os.path.exists(SELECTORS_FILE):
            with open(SELECTORS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    def _save_selectors(self):
        """保存选择器到JSON文件"""
        with open(SELECTORS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.selectors, f, indent=2, ensure_ascii=False)

    def _generate_xpath(self, element):
        """为给定的WebElement生成一个相对XPath（基于id、name、class和位置）"""
        # 尝试使用id
        element_id = element.get_attribute('id')
        if element_id:
            return f"//*[@id='{element_id}']"
        # 尝试使用name
        name = element.get_attribute('name')
        if name:
            return f"//*[@name='{name}']"
        # 使用class组合（取第一个非空class）
        classes = element.get_attribute('class')
        if classes:
            first_class = classes.split()[0]
            return f"//*[contains(@class, '{first_class}')]"
        # 最后使用标签+父级路径（简化，避免过长）
        tag = element.tag_name
        parent = element.find_element(By.XPATH, "..")
        parent_xpath = self._generate_xpath(parent) if parent != self.driver.find_element(By.TAG_NAME, 'body') else "//body"
        siblings = parent.find_elements(By.TAG_NAME, tag)
        if len(siblings) == 1:
            return f"{parent_xpath}/{tag}"
        else:
            index = siblings.index(element) + 1
            return f"{parent_xpath}/{tag}[{index}]"

    def _manual_locate(self, element_name, timeout=10):
        """
        手动定位：弹出提示，让用户将鼠标移动到目标元素上并按下Ctrl+Shift+L
        返回定位到的WebElement及其生成的选择器
        """
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        msg = (f"无法自动定位元素「{element_name}」\n"
               f"请将鼠标移动到目标元素上，然后按下 Ctrl+Shift+L\n"
               f"超时时间：{timeout}秒")
        messagebox.showinfo("手动定位", msg)
        
        # 等待用户按下热键 Ctrl+Shift+L
        start_time = time.time()
        while time.time() - start_time < timeout:
            # 使用pyautogui检测组合键（简单轮询键盘状态，不完美，但够用）
            # 更可靠的方式是使用keyboard库，但需要额外安装；这里用pyautogui检测按键状态
            # 由于pyautogui不能直接检测热键，改为监听鼠标点击或简单等待用户确认
            # 简化：弹出一个对话框，用户点击“已定位”按钮后，记录鼠标位置元素
            # 避免复杂的热键监听，采用对话框+鼠标坐标方式
            break  # 跳出循环，改用更简单的对话框
        
        # 更简单可靠的方案：弹出对话框，用户点击“确定”后，程序获取当前鼠标位置的元素
        confirm = messagebox.askokcancel("手动定位", 
                                         "请将鼠标指向目标元素，然后点击「确定」")
        if confirm:
            # 获取鼠标坐标
            x, y = pyautogui.position()
            # 通过JavaScript获取该坐标下的元素
            js = "return document.elementFromPoint(arguments[0], arguments[1]);"
            element = self.driver.execute_script(js, x, y)
            if element:
                # 生成选择器
                selector = self._generate_xpath(element)
                self.selectors[element_name] = selector
                self._save_selectors()
                messagebox.showinfo("成功", f"已记住元素「{element_name}」\n选择器：{selector}")
                return element, selector
            else:
                messagebox.showerror("错误", "未找到元素，请重试")
                return self._manual_locate(element_name, timeout)
        else:
            raise Exception(f"用户取消手动定位元素「{element_name}」")

    def find_element(self, element_name, default_selector, timeout=10, by=By.CSS_SELECTOR):
        """
        查找元素，支持自动重试和手动回退
        :param element_name: 元素逻辑名（用于保存选择器）
        :param default_selector: 默认选择器字符串
        :param timeout: 等待超时秒数
        :param by: 选择器类型（默认CSS_SELECTOR，也可用By.XPATH）
        :return: WebElement对象
        """
        # 优先使用已保存的选择器
        selector = self.selectors.get(element_name, default_selector)
        try:
            # 尝试等待元素可见
            wait = WebDriverWait(self.driver, timeout)
            if by == By.CSS_SELECTOR:
                element = wait.until(EC.visibility_of_element_located((by, selector)))
            else:
                element = wait.until(EC.visibility_of_element_located((by, selector)))
            return element
        except TimeoutException:
            print(f"[警告] 自动定位元素「{element_name}」失败，尝试手动定位...")
            # 手动定位
            element, new_selector = self._manual_locate(element_name, timeout=10)
            # 更新选择器类型（手动定位返回的是XPath）
            # 以后查找时使用XPath
            self.selectors[element_name] = new_selector
            self._save_selectors()
            return element