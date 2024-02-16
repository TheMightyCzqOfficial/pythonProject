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
            "methodId": methodId, "status": "0", "isLoop": param["isLoop"], "loopStartNum": param["loopStartNum"],
            "loopEndNum": param["loopEndNum"], "loopTimes": param["loopTimes"], "stepParam": param["stepParam"],
            "currentStep": "0", "currentParam": ""
        })

    def setMethodStatus(self, pk, index, currentStep, currentStepName, currentParam=""):
        self.running_driver_dict[pk]["methodList"][index]["currentStep"] = currentStep
        self.running_driver_dict[pk]["methodList"][index]["currentParam"] = currentParam
        self.running_driver_dict[pk]["methodList"][index]["currentStepName"] = currentStepName
        # {
        #     "methodId": methodId, "status": "0", "isLoop": param["isLoop"], "loopStart": param["loopStart"],
        #     "loopEnd": param["loopEnd"], "stepParam": param["stepParam"], "currentStep": "0","currentParam":""
        # }

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

    def queryMethodList(self, pk):
        return self.running_driver_dict[pk]["methodList"]

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

    def switch_to_preWindow(self, pk):
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

    def find_and_click_by_text(self, pk, xpath, locateText, waitSecond):
        input_text = self.webdriver_dict[pk].find_element(By.XPATH, xpath + "[text()='" + locateText + "']")
        input_text.click()
        time.sleep(waitSecond)

    def runAction(self, step_dict, pk, param=""):
        if step_dict["action"] == "input":
            self.find_and_input(pk, param, step_dict["xpath"], step_dict["xpathIndex"],
                                step_dict["waitSecond"])
        if step_dict["action"] == "click":
            if len(param) > 0:  # // a[text() = "新闻"]
                self.find_and_click_by_text(pk, step_dict["xpath"], param,
                                            step_dict["waitSecond"])
            elif len(param) == 0:
                self.find_and_click(pk, step_dict["xpath"], step_dict["xpathIndex"],
                                    step_dict["waitSecond"])
        if step_dict["action"] == "enterFrame":
            self.exit_frame(pk)
        if step_dict["action"] == "exitFrame":
            self.enter_frame(pk, step_dict["xpath"])
        if step_dict["action"] == "refresh":
            self.refresh(pk)
        if step_dict["action"] == "switchToNewWindow":
            self.switch_to_newWindow(pk)
        if step_dict["action"] == "switchToPreWindow":
            self.switch_to_preWindow(pk)

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
        # "isLoop": param["isLoop"], "loopStart": param["loopStart"],"loopEnd": param["loopEnd"], "stepParam": param[
        # "stepParam"], "currentStep": "0", "currentParam": ""
        # 运行前参数赋值
        self.init_method(pk, methodId, param)
        index = len(self.running_driver_dict[pk]["methodList"]) - 1
        method_dict = self.jsonCon.get_method(methodId)
        step_data_dict = method_dict["stepData"]
        stepCount = method_dict["stepCount"]
        loopStartNum = param["loopStartNum"]
        loopEndNum = param["loopEndNum"]
        loopTimes = param["loopTimes"]  # 循环次数 需要与传入excel行数对应
        stepParam = param["stepParam"]
        isLoop = param["isLoop"]
        # 方法主循环
        # 变更状态，开始运行
        self.running_driver_dict[pk]["methodList"]["status"] = "1"
        for i in range(1, stepCount + 1):  # 从1开始的range要加一
            stepKey = "step" + str(i)
            step_dict = step_data_dict[stepKey]
            # pk, index, currentStep, currentStepName, currentParam=""
            # print_log(step_dict, method_dict["methodName"], i)
            if isLoop and i == loopStartNum:  # 步骤循环执行
                for looptime in range(loopTimes):
                    for loopStep in range(loopStartNum, loopEndNum + 1):
                        self.is_pause(pk)
                        self.setMethodStatus(pk, index, str(i), step_dict["stepName"])
                        stepKey = "step" + str(loopStep)
                        if stepKey in stepParam.keys():
                            currentStepParam = stepParam[stepKey][loopTimes]
                            self.runAction(step_dict, pk, currentStepParam)
                        else:
                            self.runAction(step_dict, pk)
            elif isLoop and loopStartNum < i <= loopEndNum:  # 循环的步骤已执行完毕，跳过
                pass
            else:  # 未在循环体内的步骤（单次执行步骤）
                # step_param={"step2":[1,2,3,4],"step3":[1,2,3,4],"step4":"xxx"} list代表输入的 字符串代表定位的文本
                self.is_pause(pk)
                self.setMethodStatus(pk, index, str(i), step_dict["stepName"])
                if stepKey in stepParam.keys():
                    currentStepParam = stepParam[stepKey]
                    self.runAction(step_dict, pk, currentStepParam)
                else:
                    self.runAction(step_dict, pk)
        # 运行的status代表运行状态 0-等待运行，1-运行中，2-已运行
        self.running_driver_dict[pk]["methodList"]["status"] = "2"

# if __name__ == '__main__':
# mc = MethodCon("https://fssce.powerchina.cn:9081/", "loginCWGX")
# mc.execute_login("loginCWGX", "chenzq-fjdj", "chen2000624.")
