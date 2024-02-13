from flask import Flask, request, jsonify
from test import start, reset_password
import webclass

app = Flask(__name__)
web_entity = webclass.WebEntity()
driver_global = web_entity.global_webdriver
base_url = 'http://10.27.32.6:8083/business-ui/login'


def set_res(status, msg):
    return jsonify({'status': status, 'msg': msg})


@app.route('/login', methods=['POST'])
def login():
    if driver_global.current_url != base_url:
        web_entity.open_url(base_url)
    data = request.get_json()
    print(data)
    login_flag = False
    try:
        login_flag = start(data.get('username'), data.get('password'), data.get('authCode'), driver_global)
    except Exception as e:
        print(e)
    finally:
        if data.get('authCode') == '' and login_flag:
            return set_res(1, 'success')
        elif data.get('authCode') != '' and login_flag:
            return set_res(1, 'success')
        else:
            return set_res(0, '程序异常，请查看')


@app.route('/reset_password', methods=['POST'])
def reset_password():
    data = request.get_json()
    try:
        driver_global.read_excel(data.get('filePath'))
        reset_password(driver_global.name_list, driver_global)
        print(data)
    except Exception as e:
        print(e)
    finally:
        return "True"


if __name__ == '__main__':
    app.run(port=8092, debug=True)
