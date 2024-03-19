from selenium import webdriver


class WebEntity:

    def __init__(self, SEpath):
        self.global_webdriver = None
        self.chrome_options = webdriver.ChromeOptions()
        # 调试模式 进入360浏览器文件夹运行cmd 输入.\360se.exe --remote-debugging-port=9222
        # chrome_options.debugger_address = "localhost:9222"
        self.chrome_options.add_experimental_option('detach', True)
        self.chrome_options.binary_location = SEpath
        self.chrome_options.add_argument(r'--lang=zh-CN')
        self.global_webdriver = webdriver.Chrome(options=self.chrome_options)
        # self.global_webdriver.set_window_size(width=1500, height=1000)
