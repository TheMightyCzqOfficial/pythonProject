import os

from flask import Flask, request
from flask_cors import cross_origin
from flask_socketio import SocketIO
from selenium.webdriver.common.by import By

import webclass
import time
from test import start, reset_password,get_department_by_name
from readExcel import read, write_excel, read_file,write_list
import json

app = Flask(__name__)
socketio = SocketIO()
socketio.init_app(app, cors_allowed_origins="*")
web_entity = webclass.WebEntity()
web_entity.global_webdriver = web_entity.global_webdriver
base_url = 'http://10.27.32.6:8083/business-ui/login'

UPLOAD_FOLDER = 'uploads'
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['JSON_AS_ASCII'] = "False"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def set_res(status, msg, data=None):
    return json.dumps({'status': status, 'msg': msg, 'data': data}, ensure_ascii=False)


@cross_origin()
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    file.save(os.path.join(UPLOAD_FOLDER, file.filename))
    file_data = read_file(os.path.join(UPLOAD_FOLDER, file.filename), web_entity)
    return file_data


@app.route('/login', methods=['POST'])
def login():
    login_flag = False
    data = request.get_json()
    print(data)
    try:
        web_entity.running_status['is_logining'] = True
        if web_entity.global_webdriver is None:
            clear_status()
            web_entity.set_message('正在发送验证码...', 'success', socketio)
            web_entity.open_url(base_url)
        login_flag = start(data.get('username'), data.get('password'), data.get('authCode'), web_entity)
    except Exception as e:
        web_entity.set_message('登录异常，请检查4A网站是否正常！', 'error', socketio)
        print('登录异常，请检查4A网站是否正常！')
        print(e)
        return set_res(0, '程序异常，请查看')
    finally:
        if data.get('authCode') == '' and login_flag:
            web_entity.set_message('验证码发送成功', 'success', socketio)
            return set_res(1, 'success')
        elif data.get('authCode') != '' and login_flag:
            web_entity.set_message('登录成功，欢迎你' + str(data.get('username')), 'success', socketio)
            web_entity.running_status['isLogin'] = True
            web_entity.running_status['is_logining'] = False
            return set_res(1, 'success')
        else:
            web_entity.set_message('程序异常，请查看', 'error', socketio)
            return set_res(0, '程序异常，请查看')


@app.route('/reset_password', methods=['POST'])
def reset_password_by_excel():
    if web_entity.running_status['is_resetting']:
        web_entity.set_message('批量重置密码程序正在运行....请勿重复运行', 'warning', socketio)
        return set_res(False, '批量重置密码程序正在运行....请勿重复运行')
    web_entity.no_result_list = []
    web_entity.log_list = []
    web_entity.current_num = 0
    web_entity.name_list = []
    web_entity.running_status['is_resetting'] = True
    data = request.get_json()
    col_name = data.get('colName')
    sheet_name = data.get('sheetName')
    col_num = col_name[len(col_name) - 1]
    for item in web_entity.excel_data[sheet_name]:
        name = item[int(col_num) - 1]
        web_entity.set_name_list(name, '待重置', 'info')
    socketio.emit("logAndPercentage", json.dumps({'totalNameList': web_entity.name_list}, ensure_ascii=False))
    web_entity.set_log('Excel读取完毕', 'success', socketio)
    if len(web_entity.name_list) == 0:
        web_entity.set_message('待重置人员列表为空，请检查！', 'error', socketio)
        return set_res(False, '待重置人员列表为空，请检查！')
    else:
        web_entity.set_message('批量重置密码运行....', 'success', socketio)
        # web_entity.global_webdriver.find_element(By.XPATH, "//div[@id='/user/list']").click()
        # time.sleep(2)
        web_entity.current_num = 0
        try:
            for name_list in web_entity.name_list:
                web_entity.name_list[web_entity.current_num] = [name_list[0], '重置密码成功', 'success']
                socketio.emit("logAndPercentage",
                              json.dumps({'totalNameList': web_entity.name_list}, ensure_ascii=False))
                web_entity.set_log(str(name_list[0]) + " 重置密码成功!", 'success', socketio)
                print(str(name_list[0]) + " 重置密码成功!")
                web_entity.current_num += 1
                percentage = round((web_entity.current_num / len(web_entity.name_list)), 2) * 100
                web_entity.set_percentage(percentage, 'resetPercentage', socketio)
                time.sleep(2)
                # reset_password(name,web_entity, socketio)
        except Exception as e:
            web_entity.set_message('批量重置密码程序运行错误===>' + str(e), 'error', socketio)
            print(e)
            return set_res(False, '')
        finally:
            web_entity.running_status['is_resetting'] = False
            web_entity.set_message('批量重置密码运行....完毕', 'success', socketio)
            return set_res(True, '批量重置密码运行完毕')


@app.route('/get_department', methods=['POST'])
def get_department():
    web_entity.no_result_list = []
    web_entity.log_list = []
    web_entity.current_num = 0
    web_entity.name_list = []
    data = request.get_json()
    col_name = data.get('colName')
    sheet_name = data.get('sheetName')
    col_num = col_name[len(col_name) - 1]
    for item in web_entity.excel_data[sheet_name]:
        name = item[int(col_num) - 1]
        web_entity.set_name_list(name, '待查询', 'info')
    socketio.emit("logAndPercentage", json.dumps({'totalNameList': web_entity.name_list}, ensure_ascii=False))
    web_entity.set_log('Excel读取完毕', 'success', socketio)
    if len(web_entity.name_list) == 0:
        web_entity.set_message('查询列表为空，请检查！', 'error', socketio)
        return set_res(False, '查询列表为空，请检查！')
    else:
        web_entity.global_webdriver.find_element(By.XPATH, "//div[@id='/user/list']").click()
        time.sleep(2)
        web_entity.current_num = 0
        try:
            for name_list in web_entity.name_list:
                socketio.emit("logAndPercentage",
                              json.dumps({'totalNameList': web_entity.name_list}, ensure_ascii=False))
                percentage = round((web_entity.current_num / len(web_entity.name_list)), 2) * 100
                web_entity.set_percentage(percentage, 'resetPercentage', socketio)
                time.sleep(2)
                department = get_department_by_name(str(name_list[0]),web_entity, socketio)
                web_entity.name_department.append([str(name_list[0]),department])
                web_entity.name_list[web_entity.current_num] = [name_list[0], '查询成功', 'success']
                web_entity.set_log(str(name_list[0]) + " 查询成功!"+"部门：" + department, 'success', socketio)
                web_entity.current_num += 1
        except Exception as e:
            web_entity.set_message('程序运行错误===>' + str(e), 'error', socketio)
            print(e)
            return set_res(False, '')
        finally:
            web_entity.set_message('自动生成文件....', 'success', socketio)
            write_list(web_entity.name_department)

            web_entity.running_status['is_resetting'] = False
            web_entity.set_message('运行完毕', 'success', socketio)
            return set_res(True, '运行完毕')


@socketio.on('init')
def show_init():
    print("get init info...")
    socketio.emit("initMessage", json.dumps({'isLogin': web_entity.running_status['isLogin']}, ensure_ascii=False))


@socketio.on('connect')
def connect():
    clear_status()
    global web_entity
    web_entity = webclass.WebEntity()
    web_entity.running_status['message'] = True
    print(web_entity.name_list)
    print(web_entity.log_list)
    print(web_entity.message_list)
    print("connect..")


@socketio.on('disconnect')
def disconnect():
    print("disconnect...")


@socketio.on('clearStatus')
def clear_status():
    for i in list(web_entity.running_status.keys()):
        web_entity.running_status[i] = False
    web_entity.log_list = []
    web_entity.message_list = []
    web_entity.name_list = []
    web_entity.no_result_list = []

    print("clearStatus...")


if __name__ == '__main__':
    socketio.run(app, debug=True, port=8092)
