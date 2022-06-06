from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import time
import os

from Get_excel import read_cols


def screenshots(names, urls, result_address):
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

        width = 900
        height = 2250

        driver.set_window_size(width, height)

        file_name = result_address + "\\" + name + ".png"
        driver.save_screenshot(file_name)
        print("%s：截图完毕" % name)
        driver.close()

    print("已全部截图")


if __name__ == "__main__":
    xlsx_path = r"D:\Work\系统\\618\项目导入KOC名单 _618.xlsx"
    name_list = read_cols(xlsx_path, "A")
    url_list = read_cols(xlsx_path, "C")

    name_list.pop(0)
    url_list.pop(0)

    result_address = r"D:\Work\系统\\618\截图\\"

    screenshots(name_list, url_list, result_address)
