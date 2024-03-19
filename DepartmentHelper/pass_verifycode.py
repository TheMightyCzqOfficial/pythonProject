import os
import datetime
import cv2  # opencv库
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
import urllib  # 网络访问


def get_canny(img):
    blurred = cv2.GaussianBlur(img, (5, 5), 0, 0)
    canny = cv2.Canny(blurred, 0, 100)  # 轮廓
    # cv2.imshow("Output", canny)
    # cv2.waitKey(0)
    return canny


def get_offset(verify_img_name, code_img_name):
    verify_img = cv2.imread(verify_img_name)
    code_img = cv2.imread(code_img_name)
    get_canny(verify_img)
    get_canny(code_img)
    image = get_canny(verify_img)
    template = get_canny(code_img)
    result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
    (minVal, maxVal, minLoc, maxLoc) = cv2.minMaxLoc(result)
    (startX, startY) = maxLoc
    endX = startX + template.shape[1]
    endY = startY + template.shape[0]
    cv2.rectangle(image, (startX, startY), (endX, endY), (255, 255, 255), 3)
    # cv2.imshow("Output", image)
    # cv2.waitKey(0)
    print("滑块验证距离======>" + str(startX))
    # print(startX + (endX - startX) / 2)
    return startX


def move_verify(global_webdriver, verify_img_name, code_img_name):
    smallImage = global_webdriver.find_element(By.XPATH, '//div[@class="verify-move-block verify-icon '
                                                         'iconfont icon-right"]')
    ActionChains(global_webdriver).click_and_hold(smallImage).perform()
    ActionChains(global_webdriver).move_by_offset(xoffset=get_offset(verify_img_name, code_img_name),
                                                  yoffset=0).perform()
    ActionChains(global_webdriver).release().perform()
    os.remove(verify_img_name)
    os.remove(code_img_name)


def login(global_webdriver):
    now = datetime.datetime.now()
    date_str = now.strftime("%Y%m%d%H%M%S")
    img_list = global_webdriver.find_elements(By.XPATH, "//img")
    verify_img = img_list[2].get_attribute("src")
    verify_img_name = date_str + "verify_img.png"
    code_img_name = date_str + "code_img.png"
    urllib.request.urlretrieve(verify_img, verify_img_name)
    code_img = img_list[3].get_attribute("src")
    urllib.request.urlretrieve(code_img, code_img_name)
    move_verify(global_webdriver, verify_img_name, code_img_name)
