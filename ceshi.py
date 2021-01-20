# import win32gui
# from PyQt5.QtWidgets import QApplication
# from PyQt5.QtGui import *
# import win32gui
# import sys
# hwnd_title = {}
# def get_all_hwnd(hwnd, mouse):
#     if (win32gui.IsWindow(hwnd) and
#             win32gui.IsWindowEnabled(hwnd) and
#             win32gui.IsWindowVisible(hwnd)):
#         hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})
#
#
# win32gui.EnumWindows(get_all_hwnd, 0)
#
# for h, t in hwnd_title.items():
#     if t:
#         print(h, t)
#         if t == '雷电模拟器':
#             left, top, right, bottom = win32gui.GetWindowRect(h)
#             print(left, top, right, bottom)
#
# hwnd = win32gui.FindWindow(None, '雷电模拟器')
# app = QApplication(sys.argv)
# screen = QApplication.primaryScreen()
# img = screen.grabWindow(hwnd).toImage()
# img.save("screenshot.jpg")
import openpyxl


device_num='123'
print("序列号为"+device_num)


try:
    book = openpyxl.load_workbook('养户摄像头对应表.xlsx')
    sheet = book.worksheets[0]
    devices_column_index=1
    for column in sheet.columns:

        # print(column[0].value)
        if (column[0].value == '序列号'):
            column_num=devices_column_index
            for i in range(len(column)):
                if (column[i].value == None):
                    devices_num = i
                    print("已录入设备数目为" + str(devices_num - 1))
                    break
        devices_column_index +=1
    sheet.cell(row=devices_num+1, column=column_num, value=device_num)
    book.save('养户摄像头对应表.xlsx')
    book.close()
except:
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")
    print("录入设备序列号失败，把excel关闭，不然读写不了")


