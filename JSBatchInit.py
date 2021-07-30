#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 6/25/21 2:43 PM
# Author: Sun
# Software: PyCharm
# Description:
#
#
# ========================================
import os
import sys
import tkinter

from Controller.JSETLController import JSETLController
from View.JSETLWindow import JSETLWindow

if __name__ == '__main__':
    # 第一次创建路径

    my_path = os.path.dirname(os.path.realpath(sys.argv[0]))
    my_path = os.path.join(my_path, 'ETLFile')

    data = os.path.join(my_path, 'data')
    if os.path.exists(data):
        print('data文件夹已存在')
    else:
        print('data文件夹不存在')
        os.makedirs(data)

    heads = os.path.join(my_path, 'heads')
    if os.path.exists(heads):
        print('heads文件夹已存在')
    else:
        print('heads文件夹不存在')
        os.makedirs(heads)

    result = os.path.join(my_path, 'result')
    if os.path.exists(result):
        print('result文件夹已存在')
    else:
        print('result文件夹不存在')
        os.makedirs(result)

    # # print(my_path)
    # data = '%s/AutoETL_Data' % (my_path)
    # heads = '%s/AutoETL_Data' % (my_path)
    # result = '%s/AutoETL_Data' % (my_path)

    # if not os.path.exists(data):
    #     os.makedirs(data)
    # print(os.path.exists(heads))
    # if not os.path.exists(heads):
    #     os.makedirs(heads)
    # if not os.path.exists(result):
    #     os.makedirs(result)

    controller = JSETLController()
    etlview = JSETLWindow(tkinter.Tk(), controller)
    tkinter.mainloop()