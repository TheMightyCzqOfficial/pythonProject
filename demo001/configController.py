import json
import os
import datetime


# 检查文件路径
def path_check():
    if not os.path.exists("jsonfile"):
        os.mkdir("jsonfile")
    if not os.path.exists("jsonfile/method_config.json"):
        open("jsonfile/method_config.json", "w").close()
        update_config_file("jsonfile/method_config.json", {})

    if not os.path.exists("jsonfile/method_list.json"):
        open("jsonfile/method_list.json", "w").close()
        update_config_file("jsonfile/method_list.json", {})

    if not os.path.exists("jsonfile/user.json"):
        open("jsonfile/user.json", "w").close()
        update_config_file("jsonfile/user.json", {})

    if not os.path.exists("jsonfile/config.json"):
        open("jsonfile/config.json", "w").close()
        update_config_file("jsonfile/config.json", {
            "user": "jsonfile/user.json",
            "method": "jsonfile/method_config.json"
        })


# 读取文件
def read_config_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        config_data = json.load(f)
    return config_data


# 写入文件
def update_config_file(file_path, new_config):
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(new_config, f, indent=4)


def has_key(data, key):
    if key in data:
        return True
    else:
        return False


class JsonCon:

    def __init__(self):
        path_check()
        self.config = read_config_file("jsonfile/config.json")

    def save_by_filename(self, filename, data):
        update_config_file(self.config[filename], data)

    def get_config_by_filename(self, filename):
        data = read_config_file(self.config[filename])
        return data

    def get_user_config(self):
        data = read_config_file(self.config["user"])
        return data

    def get_user_by_name(self, name):
        data = read_config_file(self.config["user"])
        return data[name]

    def del_user_by_name(self, pk):
        userdata = read_config_file(self.config["user"])
        del userdata[pk]
        self.save_by_filename("user", userdata)

    def set_user_by_name(self, pk, data):
        data = {"username": data["username"], "password": data["password"], "type": data["type"]}
        userdata = read_config_file(self.config["user"])
        if not has_key(userdata, pk):
            now = datetime.datetime.now()
            date_str = now.strftime("%Y%m%d%H%M%S")
            userdata["U" + date_str] = data
        else:
            userdata[pk] = data
        self.save_by_filename("user", userdata)

    def get_method_list(self):
        return read_config_file(self.config["method_list"])

    def get_method(self, methodId):
        method_list = read_config_file(self.config["method_list"])
        method_data = read_config_file(self.config["method"])
        try:
            for item in method_list:
                if item["methodId"] == methodId:
                    method_data[methodId].update(item)
                    return method_data[methodId]
        except AttributeError as e:
            return e

    def add_method(self, data):
        try:
            now = datetime.datetime.now()
            date_str = now.strftime("%Y-%m-%d %H:%M:%S")
            date_serial = now.strftime("%Y%m%d%H%M%S")
            methodName = data["methodName"]
            methodId = "M" + date_serial
            methodType = data["methodType"]
            userType = data["userType"]
            stepCount = data["stepCount"]
            createDate = date_str
            editable = data["editable"]
            stepData = data["stepData"]
            url = data["url"]
            method_list = read_config_file(self.config["method_list"])
            method_list.append(
                {"methodName": methodName, "methodId": methodId, "url": url, "methodType": methodType,
                 "userType": userType,
                 "stepCount": stepCount, "createDate": createDate, "editable": editable})
            method_data = read_config_file(self.config["method"])
            method_data[methodId] = {"stepData": stepData}
            self.save_by_filename("method_list", method_list)
            self.save_by_filename("method", method_data)
        except AttributeError as e:
            return e

    def del_method(self, methodId):
        try:
            method_list = read_config_file(self.config["method_list"])
            for index, item in enumerate(method_list):
                if item["methodId"] == methodId:
                    del method_list[index]
            method_data = read_config_file(self.config["method"])
            del method_data[methodId]
            self.save_by_filename("method_list", method_list)
            self.save_by_filename("method", method_data)
        except AttributeError as e:
            return e

    def set_method(self, data):
        try:
            now = datetime.datetime.now()
            date_str = now.strftime("%Y-%m-%d %H:%M:%S")
            createDate = date_str
            methodId = data["methodId"]
            methodName = data["methodName"]
            methodType = data["methodType"]
            userType = data["userType"]
            stepCount = data["stepCount"]
            editable = data["editable"]
            stepData = data["stepData"]
            url = data["url"]
            method_list = read_config_file(self.config["method_list"])
            for index, item in enumerate(method_list):
                if methodId == item["methodId"]:
                    method_list[index] = {"methodName": methodName, "url": url, "methodId": methodId,
                                          "methodType": methodType,
                                          "userType": userType,
                                          "stepCount": stepCount, "editable": editable, "createDate": createDate}
            method_data = read_config_file(self.config["method"])
            method_data[methodId] = {"stepData": stepData}
            self.save_by_filename("method_list", method_list)
            self.save_by_filename("method", method_data)
        except AttributeError as e:
            return e

    def download_input_step_excel(self, methodId):
        from readExcel import write4download
        try:
            step_data = self.get_method(methodId)['stepData']
            input_step_list = []
            for key in step_data.keys():
                if step_data[key]['action'] == 'input':
                    input_step_list.append(step_data[key]['stepName'])
                if step_data[key]['action'] == 'click':
                    input_step_list.append(step_data[key]['stepName'] + "(下拉框选项)")
            return write4download(input_step_list)
        except AttributeError as e:
            return e


if __name__ == '__main__':
    jsco = JsonCon()
    jsco.download_input_step_excel("M20240202091212")
    # jsco.set_user_by_name("none", {"username": "jzyw-pz07", "password": "1234.asdf"})
