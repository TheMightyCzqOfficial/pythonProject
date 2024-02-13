import json

from selenium import webdriver
from readExcel import read, write_excel
import time


class WebEntity:
    global_webdriver = None
    excel_data = []
    name_list = []
    no_result_list = []
    log_list = []
    message_list = []
    current_num = 0
    name_department = []
    running_status = {'is_resetting': False, 'is_logining': False, 'message': False, 'isLogin': False}
    chrome_options = webdriver.ChromeOptions()
    # 调试模式 进入360浏览器文件夹运行cmd 输入.\360se.exe --remote-debugging-port=9222
    # chrome_options.debugger_address = "localhost:9222"
    chrome_options.add_experimental_option('detach', True)
    chrome_options.binary_location = r"C:\Users\Administrator\AppData\Roaming\360se6\Application\360se.exe"
    chrome_options.add_argument(r'--lang=zh-CN')

    def __init__(self):
        pass

    def return_to_parent(self, num):
        print("return_to_parent")
        for i in range(num):
            self.global_webdriver.switch_to.parent_frame()

    def open_url(self, url):
        self.global_webdriver = webdriver.Chrome(options=self.chrome_options)
        print("open_url" + str(url))
        # 'http://10.27.32.6:8083/business-ui/login'
        self.global_webdriver.get(str(url))

    def read_excel(self, file_path):
        print("read_excel")
        self.name_list = read(file_path)

    def set_message(self, message, status, socketio):
        self.message_list.append([time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), message, status])
        log = {'messageList': self.message_list}
        socketio.emit("messageFormPython", json.dumps(log, ensure_ascii=False))

    def set_log(self, log, status, socketio):
        self.log_list.append([time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), log, status])
        log = {'log': self.log_list}
        socketio.emit("logAndPercentage", json.dumps(log, ensure_ascii=False))

    def set_name_list(self, name, description, status):
        self.name_list.append([name, description, status])

    def set_percentage(self, percentage, percentage_type, socketio):
        log = {"percentage": {'percentage': percentage, 'percentageType': percentage_type}}
        socketio.emit("logAndPercentage", json.dumps(log, ensure_ascii=False))
