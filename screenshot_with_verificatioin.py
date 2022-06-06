from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import time
import os
from verification_killer_functions import ver_killer
from Get_excel import read_cols


def screenshots_with_ver(names, urls, result_address):
    """
    """
    while urls:
        url = urls.pop()
        name = names.pop()

        chromedriver = r"D:\Work\Py_Project\Project_Screenshot\chromedriver.exe"
        os.environ["webdriver.chrome.driver"] = chromedriver

        chrome_options = Options()
        chrome_options.add_argument('headless')
        driver = webdriver.Chrome(chromedriver, options=chrome_options)

        picture_url = url
        driver.get(picture_url)
        # print(dir(driver))
        time.sleep(1)

        # 接下来是全屏的关键，用js获取页面的宽高，如果有其他需要用js的部分也可以用这个方法
        width = 900
        height = 2250

        driver.set_window_size(width, height)

        ver_killer(driver)  # 新增功能

        file_name = result_address + name + ".png"
        driver.save_screenshot(file_name)
        print("%s：截图完毕" % name)
        driver.close()

    print("已全部截图")
