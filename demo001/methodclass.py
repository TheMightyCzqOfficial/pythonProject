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
        self.running_driver_dict = {}
        self.login_list = []
        self.webdriver_dict = {}

    def init_driver(self, pk):
        # 浏览器实例以账号PK为主键的dict 同时生成每个pk中运行的方法（单线程运行，不允许重复运行）
        userData = self.jsonCon.get_user_config()[pk]
        if userData["isLogin"]:
            return False
        methodId = ""
        url = ""
        if userData["type"] == '1':
            url = self.jsonCon.config["cwgxUrl"]
            methodId = "loginCWGX"
        elif userData["type"] == '2':
            url = self.jsonCon.config["ygekUrl"]
            methodId = "loginYGEK"
        now = datetime.datetime.now()
        date_str = now.strftime("%Y-%m-%d %H:%M:%S")
        self.login_list.append(pk)
        self.running_driver_dict[pk] = {"methodList": [], "isPause": False, "loginTime": date_str}
        self.webdriver_dict[pk] = webclass.WebEntity(url).global_webdriver
        # self.webdriver_dict[pk].minimize_window()
        if self.run_login(pk, methodId, userData["username"], userData["password"]):
            self.jsonCon.change_login_status(pk)
            self.webdriver_dict[pk].minimize_window()
            return True

        # init方法打开一个浏览器实例写入url网址，然后开始执行方法 
        # {"methodId": "", "methodName": method_dict["methodName"], "isPause": True, 
        # "currentStep": '',"startTime": date_str,"status": '等待开始'} 

    def init_method(self, pk, methodId, param):
        # 运行的status代表运行状态 0-等待运行，1-运行中，2-已运行
        # 运行的step的参数模板文件使用dict stepName:[] 数组长度代表循环的次数
        # step_param={"step2":[1,2,3,4],"step3":[1,2,3,4],"step4":"xxx"} list代表输入的 字符串代表定位的文本
        self.running_driver_dict[pk]["methodList"].append({
            "methodId": methodId, "status": "0", "isLoop": param["isLoop"], "loopStart": param["loopStart"],
            "loopEnd": param["loopEnd"], "stepParam": param["stepParam"], "currentStep": "0","currentParam":""
        })

    def is_pause(self, pk):
        if self.running_driver_dict[pk]["isPause"]:
            time.sleep(1)
            self.is_pause(pk)
        else:
            pass

    def changePause(self, pk):
        self.running_driver_dict[pk]["isPause"] = not self.running_driver_dict[pk]["isPause"]

    def queryDriver(self):
        resList = []
        userData = self.jsonCon.get_user_config()
        for key in self.running_driver_dict.keys():
            username = userData[key]["username"]
            usertype = userData[key]["type"]
            driver_dict = self.running_driver_dict[key]
            loginTime = driver_dict["loginTime"]
            isPause = driver_dict["isPause"]
            res_dict = {"userName": username, "userType": usertype, "loginTime": loginTime,
                        "status": "暂停中" if isPause else "运行中"}
            resList.append(res_dict)
        return resList

    def enter_frame(self, pk, iframe):  # 传入iframe Xpath
        self.webdriver_dict[pk].switch_to.frame(iframe)

    def exit_frame(self, pk):
        self.webdriver_dict[pk].switch_to.parent_frame()

    def refresh(self, pk):
        self.webdriver_dict[pk].refresh()

    def switch_to_newWindow(self, pk):
        # 获取所有窗口的句柄（handle）列表
        window_handles = self.webdriver_dict[pk].window_handles

        # 切换到最后一个窗口（新窗口）
        new_window = window_handles[-1]
        self.webdriver_dict[pk].switch_to.window(new_window)

    def return_to_preWindow(self, pk):
        window_handles = self.webdriver_dict[pk].window_handles
        self.webdriver_dict[pk].close()
        self.webdriver_dict[pk].switch_to.window(window_handles[0])

    def find_element_from_list(self, pk, xpath, num):
        return self.webdriver_dict[pk].find_elements(By.XPATH, xpath)[num]

    def find_and_input(self, pk, text, xpath, num, waitSecond):
        input_text = self.webdriver_dict[pk].find_elements(By.XPATH, xpath)[num]
        input_text.clear()
        input_text.send_keys(text)
        time.sleep(waitSecond)

    def find_and_click(self, pk, xpath, num, waitSecond):
        input_text = self.webdriver_dict[pk].find_elements(By.XPATH, xpath)[num]
        input_text.click()
        time.sleep(waitSecond)

    def find_and_click_by_text(self, pk, xpath, text, waitSecond):
        input_text = self.webdriver_dict[pk].find_element(By.XPATH, xpath + "[text()='" + text + "']")
        input_text.click()
        time.sleep(waitSecond)

    def run_logout(self, pk):
        self.webdriver_dict[pk].quit()
        del self.webdriver_dict[pk]
        self.jsonCon.change_login_status(pk)

    def run_login(self, pk, methodId, username, password):
        from special_method import pass_verify
        method_dict = self.jsonCon.get_method(methodId)
        step_data_dict = method_dict["stepData"]
        stepCount = method_dict["stepCount"]
        for i in range(1, stepCount + 1):  # 从1开始的range要加一
            step_dict = step_data_dict["step" + str(i)]
            print_log(step_dict, method_dict["methodName"], i)
            try:
                if step_dict["action"] == "input":
                    if "username" in step_dict:
                        self.find_and_input(pk, username, step_dict["xpath"], step_dict["xpathIndex"],
                                            step_dict["waitSecond"])
                    if "password" in step_dict:
                        self.find_and_input(pk, password, step_dict["xpath"], step_dict["xpathIndex"],
                                            step_dict["waitSecond"])
                if step_dict["action"] == "click":
                    self.find_and_click(pk, step_dict["xpath"], step_dict["xpathIndex"],
                                        step_dict["waitSecond"])
                if step_dict["action"] == "pass_verify":
                    pass_verify(self.webdriver_dict[pk])
                    time.sleep(step_dict["waitSecond"])
            except BaseException as e:
                return step_dict["stepName"] + "运行异常"
        return True

    def run_method(self, pk, methodId, param):
        # param:循环执行步骤loopStart loopEnd 、每个步骤输入的参数inputDataList[{step1:xxx,step2:xxx}]
        # 后期看情况升级衔接执行下一个方法或者中间执行某些特定方法
        self.running_driver_dict[pk]["methodList"].append(methodId)
        method_dict = self.jsonCon.get_method(methodId)
        step_data_dict = method_dict["stepData"]
        stepCount = method_dict["stepCount"]
        for i in range(1, stepCount + 1):  # 从1开始的range要加一
            self.is_pause(methodId)
            step_dict = step_data_dict["step" + str(i)]
            self.running_method_list[index]["currentStep"] = step_dict["stepName"]
            print_log(step_dict, method_dict["methodName"], i)
            if step_dict["action"] == "input":
                self.find_and_input(pk, step_dict["inputText"], step_dict["xpath"], step_dict["xpathIndex"],
                                    step_dict["waitSecond"])
            if step_dict["action"] == "click":
                if len(step_dict["locateText"]) > 0:  # // a[text() = "新闻"]
                    self.find_and_click_by_text(pk, step_dict["xpath"], step_dict["xpathIndex"],
                                                step_dict["waitSecond"])
                self.find_and_click(pk, step_dict["xpath"], step_dict["xpathIndex"],
                                    step_dict["waitSecond"])
            if step_dict["action"] == "enterFrame":
                self.exit_frame(pk)
            if step_dict["action"] == "exitFrame":
                self.enter_frame(pk, step_dict["xpath"])
        del self.running_method_list[index]

# if __name__ == '__main__':
# mc = MethodCon("https://fssce.powerchina.cn:9081/", "loginCWGX")
# mc.execute_login("loginCWGX", "chenzq-fjdj", "chen2000624.")
