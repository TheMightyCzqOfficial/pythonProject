import os

from flask import Flask, request, jsonify
from flask_cors import cross_origin
from readExcel import read, write_excel, write4download
import webclass
from configController import JsonCon
from methodclass import MethodCon

app = Flask(__name__)
SUCCESS_STATUS = 111
FAIL_STATUS = 000
UPLOAD_FOLDER = 'uploads'
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['JSON_AS_ASCII'] = "False"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
mc = MethodCon()
JsonCon().resetLoginStatus()


@cross_origin()
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    file.save(os.path.join(UPLOAD_FOLDER, file.filename))
    # file_data = read(os.path.join(UPLOAD_FOLDER, file.filename))
    return set_res(SUCCESS_STATUS, "")


def set_res(status, data):
    return jsonify({'status': str(status), 'data': data})


@app.route('/getUserConfig', methods=['POST'])
def getUserConfig():
    jc = JsonCon()
    return set_res(SUCCESS_STATUS, jc.get_config_by_filename("user"))


@app.route('/setUserConfig', methods=['POST'])
def setUserConfig():
    data = request.get_json()
    jc = JsonCon()
    jc.set_user_by_name(data["pk"], data)
    return set_res(SUCCESS_STATUS, "")


@app.route('/addUserConfig', methods=['POST'])
def addUserConfig():
    data = request.get_json()
    jc = JsonCon()
    jc.set_user_by_name("new", data)
    return set_res(SUCCESS_STATUS, "")


@app.route('/delUserConfig', methods=['POST'])
def delUserConfig():
    data = request.get_json()
    jc = JsonCon()
    jc.del_user_by_name(data["pk"])
    return set_res(SUCCESS_STATUS, "")


@app.route('/getMethodList', methods=['POST'])
def getMethodList():
    jc = JsonCon()
    jc.get_method_list()
    return set_res(SUCCESS_STATUS, jc.get_method_list())


@app.route('/getMethodData', methods=['POST'])
def getMethodData():
    data = request.get_json()
    jc = JsonCon()
    return set_res(SUCCESS_STATUS, jc.get_method(data["methodId"]))


@app.route('/setMethod', methods=['POST'])
def setMethod():
    data = request.get_json()
    jc = JsonCon()
    jc.set_method(data)
    return set_res(SUCCESS_STATUS, "")


@app.route('/addMethod', methods=['POST'])
def addMethod():
    data = request.get_json()
    jc = JsonCon()
    jc.add_method(data)
    return set_res(SUCCESS_STATUS, "")


@app.route('/delMethod', methods=['POST'])
def delMethod():
    data = request.get_json()
    jc = JsonCon()
    jc.del_method(data["methodId"])
    return set_res(SUCCESS_STATUS, "")


@app.route('/downloadMethodExcel', methods=['POST'])
def downloadMethodExcel():
    data = request.get_json()
    jc = JsonCon()
    return set_res(SUCCESS_STATUS, jc.download_input_step_excel(data["methodId"]))


@app.route('/openAndLogin', methods=['POST'])
def openAndLogin():
    data = request.get_json()
    loginFlag = mc.init_driver(data["pk"])
    if loginFlag:
        return set_res(SUCCESS_STATUS, "登录成功")
    else:
        return set_res(FAIL_STATUS, "登录失败，请检查是否开启VPN，是否重复登录，是否登录异常")


@app.route('/closeAndLogout', methods=['POST'])
def closeAndLogout():
    data = request.get_json()
    mc.run_logout(data["pk"])
    return set_res(SUCCESS_STATUS, "已退出登录")


@app.route('/controlWindow', methods=['POST'])
def controlWindow():
    data = request.get_json()
    mc.controlWindow(data["pk"], data["action"])
    return set_res(SUCCESS_STATUS, "")


@app.route('/getWebList', methods=['POST'])
def getWebList():
    return set_res(FAIL_STATUS, mc.queryDriver())


@app.route('/runMethod', methods=['POST'])
def runMethod():
    data = request.get_json()
    # methodParam = {"isLoop": data["isLoop"],"loopStratNum":data["loopStratNum"],"loopEndNum":data["loopEndNum"]}
    # mc.run_method(data["pk"],data["methodId"])
    return set_res(SUCCESS_STATUS, "已退出登录")


if __name__ == '__main__':
    app.run(port=8092, debug=True)
