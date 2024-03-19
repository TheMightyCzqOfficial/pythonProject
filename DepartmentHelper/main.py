import datetime
import logging
import os
import sys
import time
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from webclass import WebEntity
import pass_verifycode
from readExcel import readXlsFile, readXlsFile4YG
import json


def isDebug():
    if input("isDebugMode?y/n:") == 'y':
        return ['C:/Users/13458/AppData/Roaming/360se6/Application/360se.exe', 'jzyw-pz07', '1234.asdf']
    else:
        return False


def initLogging():
    if not os.path.exists("log"):
        os.mkdir("log")
    dateStr = datetime.datetime.now().strftime("%Y年%m月%d日 %H点%M分%S秒")
    logging.basicConfig(filename="./log/" + dateStr + "运行日志" + 'log.log',
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s  来自方法：<%(funcName)s>',
                        level=logging.INFO, encoding="utf-8")


# def getWaitTime(option):
#     timeDict = {
#         "high":{
#             "login":5,
#             "switchOrganization":5,
#             "enterMenu":5,
#             "saveBasicInfoWithParent":15,
#             "saveBasicInfo":5,
#             "saveXZZZ":5,
#             "switchToXZZZ":3,
#             "refresh":7,
#
#
#         }
#     }
#     return

class DepartmentHelper:
    def __init__(self):
        initLogging()
        flag = isDebug()
        cwUrl = "https://10.19.18.12:9081/grm/dap/mapp/dccweb/bizcomponent/index.html#/login/ecp-login-zdj"
        ygUrl = "http://10.19.18.12:33021/metamodel/necp/mapp/metamodel/component.ef/bcp/login.html"
        self.errorDict = {}
        self.retryTimes = 0
        self.ygOption = 0
        if not flag:
            try:
                self.SEpath = input("请输入360浏览器的路径(360se.exe)：\n")
                print("请选择需要登录的平台：1=财务共享、2=远光二开：")
                if input() == "1":
                    self.url = cwUrl
                    self.loginType = '1'
                else:
                    self.url = ygUrl
                    self.loginType = '2'
                self.webDriver = WebEntity(self.SEpath).global_webdriver
                self.stepProcess = 0
                self.currentStepProcess = 1
                self.startUp()
            except BaseException as e:
                logging.error(e)
                print("初始化失败！请确认浏览器路径，是否配置环境变量")
                self.__init__()
        else:
            print("请选择需要登录的平台：1=财务共享、2=远光二开：")
            if input() == "1":
                self.url = cwUrl
                self.loginType = '1'
            else:
                self.url = ygUrl
                self.loginType = '2'
            self.webDriver = WebEntity(flag[0]).global_webdriver
            self.stepProcess = 0
            self.currentStepProcess = 1
            self.startUp(flag[1], flag[2])

    def enter_frame(self, iframe):  # 传入iframe Xpath
        self.webDriver.switch_to.frame(iframe)

    def exit_frame(self):
        self.webDriver.switch_to.parent_frame()

    def refresh(self):
        self.webDriver.refresh()
        time.sleep(10)

    def submit(self, xpath, num, waitSecond=0.0):
        self.webDriver.find_elements(By.XPATH, xpath)[num].submit()
        time.sleep(waitSecond)

    def switch_to_newWindow(self):
        # 获取所有窗口的句柄（handle）列表
        window_handles = self.webDriver.window_handles

        # 切换到最后一个窗口（新窗口）
        new_window = window_handles[-1]
        self.webDriver.switch_to.window(new_window)

    def switch_to_preWindow(self):
        window_handles = self.webDriver.window_handles
        self.webDriver.close()
        time.sleep(5)
        self.webDriver.switch_to.window(window_handles[0])

    def find_element_from_list(self, xpath, num):
        return self.webDriver.find_elements(By.XPATH, xpath)[num]

    def find_and_input(self, text, xpath, num, waitSecond=0.0):
        input_text = self.webDriver.find_elements(By.XPATH, xpath)[num]
        input_text.clear()
        input_text.send_keys(text)
        time.sleep(waitSecond)

    def find_and_click(self, xpath, num, waitSecond=0.0):
        input_text = self.webDriver.find_elements(By.XPATH, xpath)[num]
        self.webDriver.execute_script("arguments[0].click()", input_text)
        time.sleep(waitSecond)

    def find_and_click_by_text(self, xpath, locateText, waitSecond=0.0):
        # self.global_webdriver.find_element(By.XPATH, "//*[contains(text(),'中文')]")
        input_text = self.webDriver.find_element(By.XPATH, xpath + "[text()='" + locateText + "']")
        self.webDriver.execute_script("arguments[0].click()", input_text)
        time.sleep(waitSecond)

    def find_and_click_by_location(self, xpath, locateText, locationXpath):
        input_text = self.webDriver.find_element(By.XPATH, xpath + "[contains(text(),'" + locateText + "')]" +
                                                 locationXpath)
        self.webDriver.execute_script("arguments[0].click()", input_text)

    def find_by_location(self, xpath, locateText):
        return self.webDriver.find_elements(By.XPATH, xpath + "[contains(text(),'" + locateText + "')]")

    # aria-level="6“
    def step_progress_bar(self, pauseSecond=0.0):
        count = int(100 * (self.currentStepProcess / self.stepProcess))
        print("\r", end="")
        print("进度: ", "▋" * (count // 2), "{}% ".format(count), end="")
        self.currentStepProcess += 1
        time.sleep(pauseSecond)

    def login(self, userName, passWord):
        isLogin = False
        try:
            if self.loginType == '2':
                isLogin = self.loginYG(userName, passWord)
            else:
                self.webDriver.minimize_window()
                print("输入账号：" + userName)
                self.find_and_input(userName, "//input", 0, 0)
                print("输入密码：" + passWord)
                self.find_and_input(passWord, "//input", 1, 3)
                print("拖图验证开始...")
                self.find_and_click("//div[@class='verify-content']", 0, 0)
                pass_verifycode.login(self.webDriver)
                print("拖图验证成功")
                time.sleep(3)
                print("点击登录按钮")
                self.find_and_click("//button[@class='el-button el-button--primary']", 0, 0)
                self.webDriver.maximize_window()
                time.sleep(8)
                print("验证是否登录成功")
                if len(self.find_by_location("//*", '工作台')) == 0:
                    print("登录失败...")
                else:
                    print("登录成功...")
                    logging.info("财务共享平台用户：%s登录成功" % userName)
                    isLogin = True
            # self.webDriver.set_window_size(width=960, height=900)
        except BaseException as e:
            logging.error(e)
            logging.info("财务共享平台用户：%s登录失败，尝试重新登录..." % userName)
            print("登录失败，尝试重新登录...")
            pass
        if not isLogin:
            self.refresh()
            self.login(input("请输入账号:\n"), input("请输入密码：\n"))

    # 远光二开登录
    def loginYG(self, userName, passWord):
        try:
            self.webDriver.minimize_window()
            time.sleep(10)
            # self.webDriver.minimize_window()
            print("输入账号：" + userName)
            self.find_and_input(userName, "//input[@id='userName']", 0, 0)
            print("输入密码：" + passWord)
            self.find_and_input(passWord, "//input[@id='password']", 0, 0)
            print("点击登录按钮")
            self.find_and_click("//div[@id='loginBtn']", 0, 10)
            print("正在验证登录是否成功")
            self.webDriver.maximize_window()
            self.find_and_click_by_text('//*', '确定', 7)
            self.find_and_click_by_text('//*', '配置视图', 0)
            print("登录成功...")
            logging.info("远光二开平台用户：%s登录成功" % userName)
            if not self.enterYGEKMenu():
                print("进入菜单失败。。。返回。。。")
            return True
            # self.webDriver.set_window_size(width=960, height=900)
        except BaseException as e:
            logging.error(e)
            print("登录失败，尝试重新登录...")
            return False

    def changeOrganization(self, depName):
        try:
            self.webDriver.execute_script("arguments[0].click()",
                                          self.webDriver.find_element(By.XPATH, "//a[contains(text(),'工作台')]"))
            self.refresh()
            print("正在切换组织：" + depName)
            logging.info("正在切换组织：" + depName)
            self.find_and_click("//*[@id='item itemlogo']/div[3]/div[1]/div/div/div[1]/span/i", 0, 1)
            self.find_and_input(depName, "//input", 4, 1)
            self.find_and_click_by_text("//i", depName, 2)
            self.find_and_click("//button[@class='el-button el-button--primary el-button--medium']", 2, 5)
            print("切换成功..")
            logging.info("切换成功")
            return True
        except BaseException as e:
            if self.retryTimes > 3:
                self.retryTimes = 0
                logging.error("重试次数超过3次，自动停止")
                return False
            else:
                self.retryTimes += self.retryTimes
                logging.error(e)
                logging.error(depName + " 切换部门失败，尝试重新切换部门...")
                return self.changeOrganization(depName)

    def enterCWGXMenu(self):
        try:
            print("进入菜单：应用中心----标准管理----设置组织单元")
            self.webDriver.execute_script("arguments[0].click()",
                                          self.webDriver.find_element(By.XPATH, "//a[contains(text(),'应用中心')]"))
            time.sleep(1.5)
            self.webDriver.execute_script("arguments[0].click()",
                                          self.webDriver.find_element(By.XPATH, "//span[contains(text(),'标准管理')]"))
            time.sleep(1.5)
            self.webDriver.execute_script("arguments[0].click()",
                                          self.webDriver.find_element(By.XPATH, "//p[contains(text(),'设置组织单元')]"))
            time.sleep(7)
            self.switch_to_newWindow()
            print("进入菜单成功..")
            return True
        except BaseException as e:
            if self.retryTimes > 3:
                self.retryTimes = 0
                logging.error("重试次数超过3次，自动停止")
                return False
            else:
                self.retryTimes += self.retryTimes
                logging.error(e)
                return self.enterCWGXMenu()

    # 远光二开进入流程管理
    def enterYGEKMenu(self):
        isSuccess = False
        try:
            if self.ygOption == '1':
                self.find_and_click_by_text('//*', '配置视图', 5)
                print("进入菜单：配置视图-->流程管理")
                print("等待菜单加载...")
                self.find_and_click_by_text('//span', '流程管理', 10)
                print("进入菜单成功..")
            elif self.ygOption == '2':
                self.find_and_click_by_text('//*', '配置视图', 5)
                print("进入菜单：配置视图-->流程管理")
                print("等待菜单加载...")
                self.find_and_click_by_text('//span', '打印设置', 5)
                action = ActionChains(self.webDriver)
                action.move_to_element(self.webDriver.find_element(By.XPATH,"//span[text()='固定实体']")).double_click().perform()
                self.find_and_click_by_text('//span', '凭证打印', 0.5)
                self.find_and_click("//a[@id='entitySettingBtn']",0,10)
                print("进入菜单成功..")
            isSuccess = True
        except BaseException as e:
            logging.error(e)
        if not isSuccess:
            if input("进入菜单失败，是否重新尝试？y/n") == "y":
                self.refresh()
                self.enterYGEKMenu()
            else:
                return isSuccess
        return isSuccess

    def runSetProcessByDepNameMain(self):
        processList = readXlsFile4YG(input("请输入流程配置模板文件路径(.xls)："))
        if not processList:
            return False
        if input("确认开始之前，请再次检查流程名称、配置的组织是否正确，数据可贵！若有请输入任意键重新加载文件，输入y开始运行") == "y":
            print("准备开始运行...3")
            time.sleep(1)
            print("准备开始运行...2")
            time.sleep(1)
            print("准备开始运行...1")
            time.sleep(1)
        else:
            return False
        for index, item in enumerate(processList):
            print("当前总进度：%s/%s" % (str(index + 1), str(len(processList))))
            self.setProcessByDepName(item)
        logging.info("文件内容执行完毕，运行情况如下:")
        for key in self.errorDict.keys():
            if self.errorDict[key]["error"] == "运行正常，没有错误" and len(self.errorDict[key]["errorDepName"]) == 0:
                self.errorDict[key] = "运行正常"
        logging.info(json.dumps(self.errorDict, ensure_ascii=False))
        return True

    # 远光二开配置流程主方法
    def setProcessByDepName(self, processDict):
        try:
            self.currentStepProcess = 1
            processType = processDict["type"]
            processName = processDict["processName"]
            self.errorDict[processName] = {"action": processType, "errorDepName": [], "error": ""}
            print("当前配置流程：" + processName)
            logging.info("当前配置流程：" + processName)
            self.stepProcess = 10 + processDict["count"]
            queryInput = self.webDriver.find_element(By.XPATH, "//input[@id='queryText']")
            queryInput.clear()
            queryInput.send_keys(processName, Keys.ENTER)
            self.step_progress_bar(1)
            if len(self.find_by_location("//td", processName)) == 0:
                print("流程未找到")
                logging.error("流程未找到")
                self.errorDict[processName]["error"] = "流程未找到"
                return False
            self.find_and_click_by_text("//td", processName, 0)
            self.step_progress_bar(0)
            enableCount = len(self.find_by_location("//td", "已发布"))  # 保存已发布与未发布数量判断是否已发布
            disableCount = len(self.find_by_location("//td", "未发布"))
            self.find_by_location("//td", "未发布")
            self.webDriver.execute_script("arguments[0].click()", self.find_by_location("//span", "编辑")[0])
            self.step_progress_bar(3)
            self.find_and_click("//*[@id='form1']/div/div/table/tbody/tr[2]/td[2]/span/span", 0)
            self.step_progress_bar(3)
            if self.webDriver.find_element(By.XPATH,
                                           "//*[text()='流程适用组织']/following-sibling::td").text.strip() == "未设置":
                print("该流程未设置使适用组织，有可能为共用流程，请检查")
                logging.error("该流程未设置使适用组织，有可能为共用流程，请检查")
                self.errorDict[processName]["error"] = "该流程未设置使适用组织，有可能为共用流程，请检查"
                self.find_and_click("//a[@id='btnCancel']", 0, 0.5)
                self.find_and_click("//a[@id='btnProcessClose']", 0, 0.5)
                return False
            self.find_and_click("//*[text()='流程适用组织']/following-sibling::td", 0)
            self.step_progress_bar(0.1)
            self.find_and_click("//*[text()='流程适用组织']/following-sibling::td/span/span", 0)
            self.step_progress_bar(7)
            # 匹配部门
            for depName in processDict["depName"]:
                logging.info("正在搜索组织：" + depName)
                self.errorDict[processName]["errorDepName"].append({depName: ""})
                isMatched = True
                if len(self.find_by_location("//span", depName)) == 0:
                    isMatched = False
                loopTime = 0
                firstMatchedRes = ""
                clickList = []
                searchInputText = self.webDriver.find_element(By.XPATH, "//input[@id='unit_filter']")
                searchInputText.clear()
                searchInputText.send_keys(depName)
                while isMatched:
                    searchInputText.send_keys(Keys.ENTER)
                    searchRes = self.webDriver.find_element(By.XPATH,
                                                            "//a[contains(@class,'curSelectedNode')and contains(@id,'UNIT')]")
                    if loopTime == 0:
                        firstMatchedRes = searchRes.text.strip()
                    if loopTime > 0 and firstMatchedRes == searchRes.text.strip():
                        logging.info("循环结束")
                        isMatched = False
                    elif searchRes.text.strip() == depName:
                        logging.info("成功匹配")
                        clickList.append(searchRes)
                    loopTime += 1
                self.step_progress_bar(0)
                if len(clickList) > 1:
                    logging.error("组织名称重复")
                    self.errorDict[processName]["errorDepName"][
                        len(self.errorDict[processName]["errorDepName"]) - 1][depName] = "组织名称重复"
                    continue
                elif len(clickList) == 0:
                    logging.error("组织名称未找到")
                    self.errorDict[processName]["errorDepName"][
                        len(self.errorDict[processName]["errorDepName"]) - 1][depName] = "组织名称未找到"
                    continue
                # 判断是否已选中
                falseCheckBox = self.webDriver.find_elements(By.XPATH,
                                                             "//*[text()='" + depName + "']/parent::a/preceding-sibling::button[@class='chk checkbox_false_full']")
                trueCheckBox = self.webDriver.find_elements(By.XPATH,
                                                            "//*[text()='" + depName + "']/parent::a/preceding-sibling::button[@class='chk checkbox_true_full']")
                if processType == "新增":
                    if len(falseCheckBox) == 1:  # 未选中
                        self.webDriver.execute_script("arguments[0].click()", clickList[0])
                    elif len(trueCheckBox) == 1:  # 已选择
                        pass
                elif processType == "删除":
                    if len(falseCheckBox) == 1:  # 未选中
                        pass
                    elif len(trueCheckBox) == 1:  # 已选择
                        self.webDriver.execute_script("arguments[0].click()", clickList[0])
                del self.errorDict[processName]["errorDepName"][len(self.errorDict[processName]["errorDepName"]) - 1]
            self.find_and_click("//a[@id='ok']", 0)
            self.step_progress_bar(2)
            self.find_and_click("//a[@id='btnSelect']", 0)
            self.step_progress_bar(0.5)
            self.find_and_click("//a[@id='btnProcessSave']", 0)
            self.step_progress_bar(7)
            enableCount1 = len(self.find_by_location("//td", "已发布"))  # 保存已发布与未发布数量判断是否已发布
            disableCount1 = len(self.find_by_location("//td", "未发布"))
            # print("已发布：%d 未发布：%d 修改后---->已发布：%d未发布：%d" % (enableCount, disableCount, enableCount1,disableCount1))
            if disableCount1 - 1 == disableCount and enableCount - 1 == enableCount1:  # 未发布数量+1 已发布数量-1 流程正常需要发布
                self.find_and_click("//a[@id='enableProcessBtn']", 0)
            elif disableCount1 == disableCount and enableCount == enableCount1:
                self.find_and_click("//a[@id='enableProcessBtn']", 0, 3)
                if len(self.find_by_location("//div", "已发布的流程未进行任何编辑")) == 1:
                    self.find_and_click("//a[@class='ZebraDialog_Button0 focusMark']", 0)
            self.step_progress_bar(6)
            self.currentStepProcess = 1
            print("\n" + processName + " 操作完毕")
            self.errorDict[processName]["error"] = "运行正常，没有错误"
            return True
        except BaseException as e:
            logging.error(e)
            print("出错了...")
            self.errorDict[processDict["processName"]]["error"] = "未知错误，请检查日志"
            return False

    def openTree(self):
        # print("检测到特殊符号，自动展开所有部门匹配中...")
        logging.info("特殊符号，自动展开所有部门")
        openList = self.webDriver.find_elements(By.XPATH, "//li[contains(@class,'close')]/i")
        if len(openList) == 0:
            return 0
        for item in openList:
            self.webDriver.execute_script("arguments[0].click()", item)
            time.sleep(2)
        if len(openList) != 0:
            logging.info("检测到子节点继续递归")
            self.openTree()

    def selectDep(self, depName):
        try:
            logging.info(depName + " 搜索部门定位中...")
            if not ('(' in depName or ')' in depName or '（' in depName or '）' in depName
                    or '-' in depName or '[' in depName or ']' in depName or '【' in depName
                    or '】' in depName):
                time.sleep(5)
                self.find_and_input(depName, "//input[@id='dwText']", 0, 1)
                self.webDriver.find_element(By.XPATH, "//input[@id='dwText']").send_keys(Keys.ENTER)
                time.sleep(1)
                self.webDriver.find_element(By.XPATH, "//input[@id='dwText']").send_keys(Keys.ENTER)
                time.sleep(3)
                if len(self.find_by_location("//div", "没有找到相关的基础组织")) == 1:
                    logging.info("未找到部门：" + depName)
                    print("未找到部门：" + depName)
                    return False
            else:
                self.openTree()
            # self.find_and_click_by_text("//a", depName, 1)
            searchList = self.find_by_location("//a", depName)
            if len(searchList) != 1:
                clickList = []
                # print(searchList)
                # print(depName)
                for item in searchList:
                    # print(item.text.strip())
                    if item.text.strip() == depName:
                        clickList.append(item)
                # print(clickList)
                if len(clickList) > 1:
                    return False
                else:
                    self.webDriver.execute_script("arguments[0].click()", clickList[0])
            else:
                self.webDriver.execute_script("arguments[0].click()", searchList[0])
            logging.info("成功定位部门：" + depName)
            return True
        except BaseException as e:
            logging.error(e)
            logging.error("定位部门 " + depName + "失败，正在重试，当前次数：" + str(self.retryTimes))
            if self.retryTimes > 3:
                self.retryTimes = 0
                logging.error("超出尝试最大次数，自动跳过")
                return False
            else:
                self.retryTimes += 1
                self.refresh()
                self.selectDep(depName)

    def runChangeDepartmentMain(self):
        fileName = input("请输入xls文件路径：")
        depList = readXlsFile(fileName)
        logging.info(fileName + " 文件内容：")
        logging.info(json.dumps(depList, ensure_ascii=False))
        if not depList:
            logging.info("重新输入文件")
            return False
        print("请确认文件内容：")
        for item in depList:
            print("组织名称：" + item["organization"] + " 组织编号：" + item["prefixCode"])
            print(item)
        print("循环次数：" + str(len(depList)) + "\n")
        if input("确认开始之前，请再次检查部门编码是否存在数字、组织前缀是否正确，数据可贵！若有请输入n重新加载文件？y/n") == "y":
            print("开始运行，请最小化窗口或不要操作鼠标...3")
            time.sleep(1)
            print("开始运行，请最小化窗口或不要操作鼠标...2")
            time.sleep(1)
            print("开始运行，请最小化窗口或不要操作鼠标...1")
            time.sleep(1)
            logging.info(fileName + " 部门变更开始执行...")
            for index, depDict in enumerate(depList):
                self.retryTimes = 1
                logging.info("当前总进度：%s/%s" % (str(index + 1), str(len(depList))))
                print("当前总进度：%s/%s" % (str(index + 1), str(len(depList))))
                if not self.changeOrganization(depDict["prefixCode"]):
                    return False
                if not self.enterCWGXMenu():
                    return False
                self.errorDict[depDict["organization"]] = []
                if not self.changeDepartment(depDict, depDict["prefixCode"]):
                    logging.error(depDict["organization"] + " 修改部门出错，已自动跳过")
                    logging.error("当前出错部门:")
                    logging.error(json.dumps(self.errorDict, ensure_ascii=False))
            logging.info("文件内容执行完毕，出错部门如下:")
            logging.info(json.dumps(self.errorDict, ensure_ascii=False))
            return True
        else:
            return False

    def stopDepartment(self, depDict):
        depName = depDict["depName"]

    def checkAndFixParentDep(self, parentDepName):
        # 判断是否需要修改上级组织
        try:
            logging.info("判断上级组织是否一致：%s" % parentDepName)
            currentParentDep = self.webDriver.find_elements(By.XPATH, "//span[@class='select2-chosen']")[0]
            self.step_progress_bar(0)
            # print(currentParentDep.text)
            if parentDepName != currentParentDep.text.strip():  # 上级部门不相等
                logging.info("当前上级组织：%s 与 系统内数据：%s 不一致，准备修改上级组织" % (parentDepName, currentParentDep.text.strip()))
                currentParentDep.click()
                time.sleep(15)
                self.webDriver.find_element(By.XPATH, "//*[@id='select2-drop']/div[2]/a").click()
                time.sleep(15)
                searchList = self.webDriver.find_elements(By.XPATH,
                                                          "//div[@class='treeBorder']//a[text()='" + parentDepName + "']", )
                # print(searchList)
                # print(parentDepName)
                if len(searchList) > 1:
                    clickList = []
                    for item in searchList:
                        # print(item.text.strip())
                        if item.text.strip() == parentDepName:
                            clickList.append(item)
                    if len(clickList) > 1:
                        logging.info("修改上级部门定位失败,已自动跳过本组织")
                        print("修改上级部门定位失败,已自动跳过本组织")
                        return False
                    elif len(clickList) == 1:
                        self.webDriver.execute_script("arguments[0].click()", clickList[0])
                elif len(searchList) == 1:
                    self.webDriver.execute_script("arguments[0].click()", searchList[0])

                else:
                    logging.error("修改上级部门定位失败,已自动跳过本组织")
                    print("修改上级部门定位失败,已自动跳过本组织")
                    return False
                self.webDriver.execute_script("arguments[0].click()",
                                              self.find_by_location("//button", '确定')[0])
                time.sleep(1)
            return True
        except Exception as e:
            logging.error(e)
            logging.error("修改上级部门定位失败,已自动跳过本组织")
            print("修改上级部门定位失败,已自动跳过本组织")
            return False

    def changeDepInfo(self, depName, sDepName, depCode):
        try:
            logging.info("修改部门信息：%s %s %s" % (depName, sDepName, depCode))
            saveType = 0
            currentDepName = self.webDriver.find_element(By.XPATH, "//input[@id='compname']").text.strip()
            if currentDepName == depName:
                saveType = 1
            self.find_and_input(depName, "//input[@id='compname']", 0, 0)
            self.step_progress_bar(0)
            self.find_and_input(sDepName, "//input[@id='scompname']", 0, 0)
            self.step_progress_bar(0)
            self.find_and_input(depCode, "//input[@id='compcode']", 0, 0)
            self.step_progress_bar(0)
            return saveType
        except Exception as e:
            logging.error(e)
            logging.error("修修改部门信息失败,已自动跳过本组织")
            print("修改部门信息定位失败,已自动跳过本组织")
            return False

    def changeDepInfo2(self, depName):
        try:
            logging.info("修改行政组织部门信息：%s " % depName)
            action = ActionChains(self.webDriver)
            action.move_to_element(self.webDriver.find_element(By.XPATH, "//*[@id='xtgldxSetTable']/tbody/tr["
                                                                         "7]/td[5]/div/div")).perform()
            self.step_progress_bar(0)
            action.click().perform()
            self.step_progress_bar(0.5)
            for i in range(2):
                action.move_to_element(self.webDriver.find_element(By.XPATH, "//*[@id='xtgldxSetTable']/tbody/tr["
                                                                             "6]/td[5]/div/div")).perform()
                self.step_progress_bar(0)
                action.click().perform()
                self.step_progress_bar(0.5)
                action.key_down(Keys.SHIFT).send_keys(Keys.END).perform()
                time.sleep(0.5)
                action.key_down(Keys.SHIFT).send_keys(Keys.HOME).key_up(Keys.SHIFT)
                self.step_progress_bar(0.05)
                action.send_keys(Keys.BACKSPACE).perform()
                self.step_progress_bar(0.1)
                action.send_keys(depName).perform()
                self.step_progress_bar(0.5)
                if i == 0:
                    for j in range(0, 14):
                        action.send_keys(Keys.ENTER).perform()
                    self.step_progress_bar(0.5)
            return True
        except Exception as e:
            logging.error(e)
            logging.error("修改行政组织部门信息失败,已自动跳过本组织")
            print("修改行政组织部门信息失败,已自动跳过本组织")
            return False

    def saveDepInfo(self, saveType):
        try:
            logging.info("保存部门基本信息")
            if saveType == 1:
                self.webDriver.execute_script("arguments[0].click()",
                                              self.find_by_location("//button", '保存')[0])
                time.sleep(15)
            elif saveType == 0:
                self.webDriver.execute_script("arguments[0].click()",
                                              self.find_by_location("//button", '保存')[0])
                time.sleep(2)
                self.webDriver.execute_script("arguments[0].click()",
                                              self.find_by_location("//button", '确定')[0])
                time.sleep(15)
            self.step_progress_bar(2)
            self.webDriver.execute_script("arguments[0].click()",
                                          self.find_by_location("//button", '保存')[0])
            time.sleep(0.5)
            if len(self.find_by_location('//div', '数据未变动')) == 1:
                self.step_progress_bar(0)
            else:
                print("保存部门基本信息失败,已自动跳过本部门")
                logging.error("保存部门基本信息失败,已自动跳过本组织")
                return False
            return True
        except Exception as e:
            logging.error(e)
            logging.error("保存部门基本信息失败,已自动跳过本组织")
            print("保存部门基本信息失败,已自动跳过本组织")
            return False

    def saveDepInfo1(self):  # 行政组织保存
        try:
            logging.info("保存行政组织基本信息")
            self.find_and_click("//div[@id='baseInfoA']", 0)
            self.step_progress_bar(1)
            self.webDriver.execute_script("arguments[0].click()", self.find_by_location("//button", '确定')[0])
            time.sleep(7)
            # js点击，牛逼！！！！！！！
            self.webDriver.execute_script("arguments[0].click()", self.find_by_location("//button", '保存')[0])
            time.sleep(0.5)
            if len(self.find_by_location('//div', '数据未变动')) != 0:
                self.step_progress_bar(0)
                self.refresh()
            else:
                print("保存部门行政组织基本信息失败,已自动跳过本组织")
                logging.error("保存部门行政组织基本信息失败,已自动跳过本组织")
                return False
            return True
        except Exception as e:
            logging.error(e)
            logging.error("保存部门行政组织基本信息失败,已自动跳过本组织")
            print("保存部门行政组织基本信息失败,已自动跳过本组织")
            return False

    def changeDepartment(self, depDict, prefix):
        currentIndex = 0
        try:
            print("当前变更中的组织：" + depDict["organization"])
            for index in range(0, depDict["count"]):
                self.retryTimes = 1
                self.currentStepProcess = 1
                self.errorDict[depDict["organization"]].append(depDict["depName"][index])
                currentIndex = index
                self.stepProcess = 24
                print("\n" + depDict["organization"] + "--部门变更进度：%d/%d" % (index + 1, depDict["count"]))
                parentDepName = depDict["parentName"][index].strip()
                depName = depDict["changeDepName"][index].strip()
                sDepName = depDict["changeDepName"][index].strip()
                depCode = prefix.strip() + str(depDict["changeDepCode"][index])
                print("原部门名称：%s，新上级部门：%s，新部门名称：%s，新部门编码：%s" % (
                    depDict["depName"][index], parentDepName, depName, depCode))
                logging.info("原部门名称：%s，新上级部门：%s，新部门名称：%s，新部门编码：%s" % (
                    depDict["depName"][index], parentDepName, depName, depCode))
                if not self.selectDep(depDict["depName"][index]):
                    self.refresh()
                    continue
                self.find_and_click("//button[@id='modifyBtn']", 0, 0)
                self.step_progress_bar(5)
                if not self.checkAndFixParentDep(parentDepName):
                    self.refresh()
                    continue
                saveType = self.changeDepInfo(depName, sDepName, depCode)
                if type(saveType) == bool:
                    self.refresh()
                    continue
                if not self.saveDepInfo(saveType):
                    self.refresh()
                    continue
                self.webDriver.execute_script("document.body.style.zoom='95%'")
                self.step_progress_bar(0)
                self.find_and_click("//div[@id='roleAttrInfoA']", 0, 10)
                self.step_progress_bar(0)
                #     self.currentStepProcess = 1
                # ================================修改组织角色属性简称======================================
                if not self.changeDepInfo2(depName):
                    self.refresh()
                    continue
                if not self.saveDepInfo1():
                    self.refresh()
                    continue
                logging.info(
                    "*************************部门：%s，修改成功**************************************" % depDict["depName"][
                        index])
                del self.errorDict[depDict["organization"]][len(self.errorDict[depDict["organization"]]) - 1]
        except BaseException as e:
            logging.error(e)
            logging.error(
                "出错单位：" + depDict["organization"] + "，出错部门：" + depDict["depName"][currentIndex] + "已自动跳过执行下一个部门")
            print("出错单位：" + depDict["organization"] + "，出错部门：" + depDict["depName"][currentIndex] + " 出错步骤：" + str(
                self.currentStepProcess))
            self.switch_to_preWindow()
            return False
        print("\n******************运行完毕************************")
        logging.info(depDict["organization"] + "================>运行完毕")
        self.switch_to_preWindow()
        return True

    def printTemplate(self):
        pass

    # //td[text()='财务主管：']/following-sibling::td[1]
    # //td[text()='记账：']/following-sibling::td[1]
    # //div[text()='文本']/parent::div/parent::td/following-sibling::td[1]
    # //td[text()='默认方案']//td[@class='sortDefault']
    # //td[text()='否']/preceding-sibling::td[1]
    # //a[@id='grid20_filterOkBtn']
    def startUp(self, userName='', passWord=''):
        logging.info("打开网页：" + self.url)
        self.webDriver.get(self.url)
        runningFlag = False
        if userName == '':
            self.login(input("请输入账号:\n"), input("请输入密码：\n"))
        else:
            self.login(userName, passWord)
        while not runningFlag:
            if self.loginType == '1':
                runningFlag = self.runChangeDepartmentMain()
                print("运行完毕,请检查日志最后错误部门报告")
            elif self.loginType == '2':
                self.ygOption = str(input("请输入需要运行的选项(1.流程下发 2.凭证模板设置)："))
                if self.ygOption == '1':
                    runningFlag = self.runSetProcessByDepNameMain()
                elif self.ygOption == '2':
                    pass
        input("运行完毕。。。")
        logging.info("运行完毕，正常退出")
        sys.exit(0)


if __name__ == '__main__':
    d = DepartmentHelper()
    d.startUp()
