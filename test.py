from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from readExcel import read, write_excel
import webclass


def get_department_by_name(name, web_entity, socketio):
    department = ''
    web_entity.set_log("正在查询部门.....\n当前人员： " + str(name), 'info', socketio)
    print("正在查询部门.....\n当前人员： " + str(name))
    iframe = web_entity.global_webdriver.find_element(By.XPATH, "//iframe[@id='ifr']")
    web_entity.global_webdriver.switch_to.frame(iframe)
    iframe1 = web_entity.global_webdriver.find_element(By.XPATH, "//iframe[@id='ifra']")
    web_entity.global_webdriver.switch_to.frame(iframe1)
    web_entity.set_log("正在执行搜索人员： " + str(name), 'info', socketio)
    time.sleep(2)
    web_entity.global_webdriver.find_element(By.XPATH, "//input[@id='fullname']").clear()
    web_entity.global_webdriver.find_element(By.XPATH, "//input[@id='fullname']").send_keys(name)
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#searchbtn").click()
    time.sleep(1)
    tabledata = web_entity.global_webdriver.find_elements(By.XPATH, "//table[@class='l-grid-body-table']/tbody/tr")
    if len(tabledata) == 1:
        web_entity.set_log("当前人员1条结果，正在重置...", 'success', socketio)
        print("当前人员1条结果，正在重置...")
        department = web_entity.global_webdriver.find_element(By.XPATH,"//table[@class='l-grid-body-table']/tbody/tr[1]/td[4]/div/span").text
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.global_webdriver.switch_to.parent_frame()
        return department
    elif len(tabledata) == 0:
        web_entity.set_log("查无记录", 'error', socketio)
        print("查无记录")
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.name_list[web_entity.current_num] = [name, '查无记录', 'error']
        return "查无记录"
    else:
        web_entity.set_log("搜索到多条记录", 'warning', socketio)
        num = input("名字重复，请选择需要操作第几个？(从上往下数)")
        department = web_entity.global_webdriver.find_element(By.XPATH,"//table[@class='l-grid-body-table']/tbody/tr[" + str(num) + "]/td[4]/div/span").text
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.global_webdriver.switch_to.parent_frame()
        return department


def reset_password(name, web_entity, socketio):
    web_entity.set_log("正在重置密码.....\n当前人员： " + str(name), 'info', socketio)
    print("正在重置密码.....\n当前人员： " + str(name))
    iframe = web_entity.global_webdriver.find_element(By.XPATH, "//iframe[@id='ifr']")
    web_entity.global_webdriver.switch_to.frame(iframe)
    iframe1 = web_entity.global_webdriver.find_element(By.XPATH, "//iframe[@id='ifra']")
    web_entity.global_webdriver.switch_to.frame(iframe1)
    web_entity.set_log("正在执行搜索人员： " + str(name), 'info', socketio)
    print("正在搜索人员： " + str(name))
    time.sleep(2)
    web_entity.global_webdriver.find_element(By.XPATH, "//input[@id='fullname']").clear()
    web_entity.global_webdriver.find_element(By.XPATH, "//input[@id='fullname']").send_keys(name)
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#searchbtn").click()
    time.sleep(1)
    tabledata = web_entity.global_webdriver.find_elements(By.XPATH, "//table[@class='l-grid-body-table']/tbody/tr")
    if len(tabledata) == 1:
        web_entity.set_log("当前人员1条结果，正在重置...", 'success', socketio)
        print("当前人员1条结果，正在重置...")
        # web_entity.global_webdriver.find_element(By.XPATH,
        #                            "//table[@class='l-grid-body-table']/tbody/tr[1]/td[1]/div/span").click()
        # web_entity.global_webdriver.find_element(By.XPATH, "//button[@id='resetPassword']").click()
        # time.sleep(1)
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.global_webdriver.switch_to.parent_frame()
        # # html = web_entity.global_webdriver.find_element(By.XPATH, "//*").get_attribute("outerHTML")
        # # print(html)
        # dialog = web_entity.global_webdriver.find_element(By.XPATH, "//iframe[@id='item_window']")
        # web_entity.global_webdriver.switch_to.frame(dialog)
        # web_entity.global_webdriver.find_element(By.XPATH, "//input[@id='confirmbtn']").click()
        # web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.name_list[web_entity.current_num] = [name, '重置密码成功', 'success']
        web_entity.set_log(str(name) + " 重置密码成功!", 'success', socketio)
        # print(str(name) + " 重置密码成功!")
    elif len(tabledata) == 0:
        web_entity.set_log("查无记录", 'error', socketio)
        print("查无记录")
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.global_webdriver.switch_to.parent_frame()
        web_entity.name_list[web_entity.current_num] = [name, '查无记录', 'error']

    else:
        web_entity.set_log("搜索到多条记录还没开发...阻塞咯！！！", 'warning', socketio)
        num = input("名字重复，请选择需要操作第几个？(从上往下数)")
        id = str(
            web_entity.global_webdriver.find_element(By.XPATH, "//table[@class='l-grid-body-table']/tbody/tr[" + str(
                num) + "]/td[2]/div/span").text)
        if input("确定重置" + id) == 1:
            web_entity.global_webdriver.find_element(By.XPATH,
                                                     "//table[@class='l-grid-body-table']/tbody/tr[" + str(
                                                         num) + "]/td[1]/div/span").click()
            web_entity.global_webdriver.find_element(By.XPATH, "//button[@id='resetPassword']").click()
        else:
            return 'error'


def start(username, password, authcode, web_entity):
    if authcode != "":
        web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#loginCode").clear()
        web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#loginCode").send_keys(authcode)
        web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#login_btn1").click()
        return True
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#loginUID").clear()
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#loginPWD").clear()
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#loginUID").send_keys(username)  # wangsw Fjdj_1234
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#loginPWD").send_keys(password)
    web_entity.global_webdriver.find_element(By.CSS_SELECTOR, "#sendSMSCode").click()
    return True


if __name__ == '__main__':
    start("chenzq", "Chenzq123")
