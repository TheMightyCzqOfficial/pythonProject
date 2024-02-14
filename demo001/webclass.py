import json
import random

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from readExcel import read, write_excel
import time
import urllib  # 网络访问
from configController import JsonCon


class WebEntity:

    def __init__(self, url):
        self.main_url = url
        self.global_webdriver = None
        self.chrome_options = webdriver.ChromeOptions()
        # 调试模式 进入360浏览器文件夹运行cmd 输入.\360se.exe --remote-debugging-port=9222
        # chrome_options.debugger_address = "localhost:9222"
        self.chrome_options.add_experimental_option('detach', True)
        self.chrome_options.binary_location = JsonCon().config["360SE_path"]
        self.chrome_options.add_argument(r'--lang=zh-CN')
        self.global_webdriver = webdriver.Chrome(options=self.chrome_options)
        self.global_webdriver.get(str(self.main_url))

    # def set_message(self, message, status, socketio):
    #     self.message_list.append([time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), message, status])
    #     log = {'messageList': self.message_list}
    #     socketio.emit("messageFormPython", json.dumps(log, ensure_ascii=False))
    #
    # def set_log(self, log, status, socketio):
    #     self.log_list.append([time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), log, status])
    #     log = {'log': self.log_list}
    #     socketio.emit("logAndPercentage", json.dumps(log, ensure_ascii=False))
    #
    # def set_name_list(self, name, description, status):
    #     self.name_list.append([name, description, status])
    #
    # def set_percentage(self, percentage, percentage_type, socketio):
    #     log = {"percentage": {'percentage': percentage, 'percentageType': percentage_type}}
    #     socketio.emit("logAndPercentage", json.dumps(log, ensure_ascii=False))
