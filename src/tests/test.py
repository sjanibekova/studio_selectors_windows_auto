from pywinauto import Desktop, Application
import os
from UIDesktop import  *
windows = Desktop(backend="uia").windows()
from main import main
import pywinauto
result = main()
print(result)
print(dir(result))
#
# ''' тестировать функционал 1С'''
# target_element = result
#

#
# '''
# TODO:
#  1)Тестирование правильности нахождения селекторов windows - окон
#  2) 1С
#
#  4) Windows;
#   3) Java apps
# '''
#
# # # 1C Testing
#

# # Test case - TaskBar
# app = pywinauto.Application(backend='uia')
# app.connect(**{'class_name': 'Shell_TrayWnd', 'title': 'Панель задач'})
# lTempObject = app
# print(lTempObject.window(**{'class_name': 'Shell_TrayWnd', 'title': 'Панель задач'}).dump_tree())
#
