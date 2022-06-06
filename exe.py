import tkinter as tk
from tkinter import filedialog
from Get_excel import read_cols
from screenshot_tool import screenshots
from screenshot_with_verificatioin import screenshots_with_ver


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    Folderpath = filedialog.askdirectory()      #获得选择好的文件夹
    print('输出截图的地址:', Folderpath)
    Filepath = filedialog.askopenfilename()      #获得选择好的文件
    print('需要截图的名单:', Filepath)

    xlsx_path = Filepath
    name_list = read_cols(xlsx_path, "A")
    url_list = read_cols(xlsx_path, "B")

    print(url_list)
    print(name_list)

    name_list.pop(0)  # 去掉列标题
    url_list.pop(0)  # 去掉列标题

    result_address = Folderpath

    screenshots(name_list, url_list, result_address)
