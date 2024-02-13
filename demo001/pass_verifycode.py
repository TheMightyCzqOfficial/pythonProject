import cv2  # opencv库
import json
import random

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
import time
import urllib  # 网络访问


def get_canny(img):
    blurred = cv2.GaussianBlur(img, (5, 5), 0, 0)
    canny = cv2.Canny(blurred, 0, 100)  # 轮廓
    # cv2.imshow("Output", canny)
    # cv2.waitKey(0)
    return canny


def get_offset():
    verify_img = cv2.imread('verify_img.png')
    code_img = cv2.imread('code_img.png')
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


def move_verify(global_webdriver):
    smallImage = global_webdriver.find_element(By.XPATH, '//div[@class="verify-move-block verify-icon '
                                                         'iconfont icon-right"]')
    ActionChains(global_webdriver).click_and_hold(smallImage).perform()
    ActionChains(global_webdriver).move_by_offset(xoffset=get_offset(), yoffset=0).perform()
    ActionChains(global_webdriver).release().perform()


def login(global_webdriver):
    img_list = global_webdriver.find_elements(By.XPATH, "//img")
    verify_img = img_list[2].get_attribute("src")
    urllib.request.urlretrieve(verify_img, 'verify_img.png')
    code_img = img_list[3].get_attribute("src")
    urllib.request.urlretrieve(code_img, 'code_img.png')
    move_verify(global_webdriver)
