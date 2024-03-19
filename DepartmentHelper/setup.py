from cx_Freeze import setup, Executable
import datetime

setup(
    name="部门操作小帮手",
    version="1.0",
    description="最新更新时间为："+datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    author="czq",
    executables=[Executable("main.py")]
)