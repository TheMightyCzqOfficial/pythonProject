import datetime

import webclass
from configController import JsonCon
from selenium.webdriver.common.by import By
from readExcel import read, write_excel
import time
import urllib  # 网络访问


def print_log(step_dict, methodName, step_num):
    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d %H:%M:%S")
    print(date_str + " 当前正在运行方法：" + methodName + " 步骤：" + str(step_num) + " 步骤名称：" + step_dict[
        "stepName"] + " 当前步骤等待时长：" + str(step_dict["waitSecond"]) + "秒")


class MethodCon:

    def __init__(self):
        self.jsonCon = JsonCon()
        self.running_method_list = []
        self.webdriver_list = []

    def init_method(self, methodId):
        if self.get_methodIndex_by_id(methodId) == -1:
            return False
        now = datetime.datetime.now()
        date_str = now.strftime("%Y-%m-%d %H:%M:%S")
        method_dict = self.jsonCon.get_method(methodId)
        self.running_method_list.append(
            {"methodId": methodId, "methodName": method_dict["methodName"], "isPause": True, "currentStep": '',
             "startTime": date_str,
             "status": '等待开始'})
        self.webdriver_list.append(webclass.WebEntity(method_dict['url'], methodId))
        return True
        # init方法打开一个浏览器实例写入url网址，然后开始执行方法

    def is_pause(self, methodId):
        if self.running_method_list[self.get_methodIndex_by_id(methodId)]["isPause"]:
            time.sleep(1)
            self.is_pause(methodId)
        else:
            pass

    def pause_by_id(self, methodId):
        self.running_method_list[self.get_methodIndex_by_id(methodId)]["isPause"] = not \
            self.running_method_list[self.get_methodIndex_by_id(methodId)]["isPause"]

    def get_running_method(self):
        return self.running_method_list

    def get_webdriver_by_id(self, methodId):
        for driver in self.webdriver_list:
            if driver.methodId == methodId:
                return driver.global_webdriver

    def get_methodIndex_by_id(self, methodId):
        num = 0
        for item in self.running_method_list:
            if item.methodId == methodId:
                return num
            num += 1
        return -1

    def return_pre(self, methodId):
        self.get_webdriver_by_id(methodId).switch_to.parent_frame()

    def enter_frame(self, methodId, iframe):  # 传入iframe Xpath
        self.get_webdriver_by_id(methodId).switch_to.frame(iframe)

    def exit_frame(self, methodId):
        self.get_webdriver_by_id(methodId).switch_to.parent_frame()

    def find_element_from_list(self, methodId, xpath, num):
        return self.get_webdriver_by_id(methodId).find_elements(By.XPATH, xpath)[num]

    def find_and_input(self, methodId, text, xpath, num, waitSecond):
        input_text = self.get_webdriver_by_id(methodId).find_elements(By.XPATH, xpath)[num]
        input_text.clear()
        input_text.send_keys(text)
        time.sleep(waitSecond)

    def find_and_click(self, methodId, xpath, num, waitSecond):
        input_text = self.get_webdriver_by_id(methodId).find_elements(By.XPATH, xpath)[num]
        input_text.click()
        time.sleep(waitSecond)

    def find_and_click_by_text(self, methodId, xpath, text, waitSecond):
        input_text = self.get_webdriver_by_id(methodId).find_element(By.XPATH, xpath + "[text()='" + text + "']")
        input_text.click()
        time.sleep(waitSecond)

    def execute_login(self, loginMethodId, methodId, username, password):
        from special_method import pass_verify
        method_dict = self.jsonCon.get_method(methodId)
        step_data_dict = method_dict["stepData"]
        stepCount = method_dict["stepCount"]
        for i in range(1, stepCount + 1):  # 从1开始的range要加一
            step_dict = step_data_dict["step" + str(i)]
            print_log(step_dict, method_dict["methodName"], i)
            if step_dict["action"] == "input":
                if "username" in step_dict:
                    self.find_and_input(loginMethodId, username, step_dict["xpath"], step_dict["xpathIndex"],
                                        step_dict["waitSecond"])
                if "password" in step_dict:
                    self.find_and_input(loginMethodId, password, step_dict["xpath"], step_dict["xpathIndex"],
                                        step_dict["waitSecond"])
            if step_dict["action"] == "click":
                self.find_and_click(loginMethodId, step_dict["xpath"], step_dict["xpathIndex"],
                                    step_dict["waitSecond"])
            if step_dict["action"] == "pass_verify":
                pass_verify(self.get_webdriver_by_id(loginMethodId))
                time.sleep(step_dict["waitSecond"])

    def execute_method(self, methodId, executeParam):
        # executeParam:循环步骤start end 、每个步骤输入的参数，后期看情况升级衔接执行下一个方法或者中间执行某些特定方法
        if not self.init_method(methodId):
            return "相同方法执行中。不允许重复创建"
        index = self.get_methodIndex_by_id(methodId)
        method_dict = self.jsonCon.get_method(methodId)
        step_data_dict = method_dict["stepData"]
        stepCount = method_dict["stepCount"]
        for i in range(1, stepCount + 1):  # 从1开始的range要加一
            self.is_pause(methodId)
            step_dict = step_data_dict["step" + str(i)]
            self.running_method_list[index]["currentStep"] = step_dict["stepName"]
            print_log(step_dict, method_dict["methodName"], i)
            if step_dict["action"] == "input":
                self.find_and_input(methodId, step_dict["inputText"], step_dict["xpath"], step_dict["xpathIndex"],
                                    step_dict["waitSecond"])
            if step_dict["action"] == "click":
                if len(step_dict["locateText"]) > 0:  # // a[text() = "新闻"]
                    self.find_and_click_by_text(methodId, step_dict["xpath"], step_dict["xpathIndex"],
                                                step_dict["waitSecond"])
                self.find_and_click(methodId, step_dict["xpath"], step_dict["xpathIndex"],
                                    step_dict["waitSecond"])
            if step_dict["action"] == "enter_frame":
                self.exit_frame(methodId)
            if step_dict["action"] == "exit_frame":
                self.enter_frame(methodId, step_dict["xpath"])
        del self.running_method_list[index]


if __name__ == '__main__':
    mc = MethodCon("https://fssce.powerchina.cn:9081/", "loginCWGX")
    mc.execute_login("loginCWGX", "chenzq-fjdj", "chen2000624.")
