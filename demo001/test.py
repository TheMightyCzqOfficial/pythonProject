import json
import random

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from readExcel import read, write_excel
import time
import urllib  # 网络访问
from configController import JsonCon
import logging
import os


class WebEntity:

    def __init__(self, url):
        self.main_url = url
        self.global_webdriver = None
        self.chrome_options = webdriver.ChromeOptions()
        # 调试模式 进入360浏览器文件夹运行cmd 输入.\360se.exe --remote-debugging-port=9222
        # chrome_options.debugger_address = "localhost:9222"
        self.chrome_options.add_experimental_option('detach', True)
        # self.chrome_options.binary_location = JsonCon().config["360SE_path"]
        ser = Service()
        ser.path = r'C:\Users\13458\Desktop\部门小助手\chromedriver.exe'
        self.chrome_options.add_argument(r'--lang=zh-CN')
        self.global_webdriver = webdriver.Chrome(options=self.chrome_options,service=ser)
        self.global_webdriver.get(str(self.main_url))

    def run(self):
        time.sleep(10)
        username = self.global_webdriver.find_elements(By.XPATH, "//input[@id='userName']")[0]
        print(username)
        username.send_keys("jzyw-pz07")

    def run1(self):
        self.global_webdriver.minimize_window()
        time.sleep(1)
        print(self.global_webdriver.get_window_size())
        print(self.global_webdriver.get_window_rect())
        print(self.global_webdriver.get_screenshot_as_png())
        self.global_webdriver.maximize_window()
        time.sleep(1)
        print(self.global_webdriver.get_window_size())
        print(self.global_webdriver.get_window_rect())
        print(self.global_webdriver.get_screenshot_as_png())
        self.global_webdriver.quit()

    def run2(self):
        # print(self.global_webdriver.page_source)
        # a = self.global_webdriver.find_element(By.XPATH, "//span[text()='登录']")
        # a.click()
        a = self.global_webdriver.find_element(By.XPATH, "//*[contains(text(),'中文')]")

        print(a.text)
        # a.click()


if __name__ == '__main__':
    wb = WebEntity('https://fssce.powerchina.cn:9081/')
    wb.run2()
