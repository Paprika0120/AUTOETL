#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 3/15/21 10:42 PM
# Author: Sun
# Software: PyCharm
# Description:
#
#
# ========================================
import tkinter
import os
import datetime
from tkinter import messagebox
import pandas as pd
from Tools.JSExcelHandler import JSExcelHandler
# from View.JSVerifyWindow import JSVerifyWindow


class JSVerifyController:
    # def __init__(self):
    #     # self.verifywindow = JSVerifyWindow()
    #     # self.verifywindow.delegate = self

    def verifyColumn(self, sourthpath, idlist, nulllist, datelist):
        # df = pd.read_excel(r'/Users/sun/Desktop/AUTOEtlTest/data/test.xls')
        df = pd.read_excel(sourthpath, dtype = 'str')
        # 先去重
        df.drop_duplicates()
        print(df.columns)

        # 身份证 ID 校验
        def idCardVerification(row):
            if JSExcelHandler.checkIsNull(row):
                if not JSExcelHandler.checkIDNumber(row):
                    return 'false'
                else:
                    return ''
            else:
                return 'false'

        # 互斥校验 TODO
        def multualExclusion(row, *args):
            # 字符处理
            pre = args[0]
            cur = args[1]
            # for foo in args:

        # 空值校验
        def nullVefify(row):
            if not JSExcelHandler.checkIsNull(row):
                return 'false'
            else:
                return ''
        # 统一时间格式
        def dateformated(string):
            # 先去除所有空格
            string = string.replace(' ', '')
            # 枚举尽可能多出现的情况
            for fmt in ["%Y/%m/%d", "%Y/%m/%d%H:%M:%S", "%Y/%m",
                        "%Y-%m-%d", "%Y-%m-%d%H:%M:%S", "%Y-%m",
                        "%Y%m%d", "%Y%m%d%H:%M:%S", "%Y%m",
                        "%Y.%m.%d", "%Y.%m", "%Y.%m.%d%H:%M:%S",
                        "%Y年%m月%d日", "%Y年", "%Y年%m月"]:
                try:
                    temp = datetime.datetime.strptime(string, fmt).date()
                    temp = temp.strftime("%Y/%m/%d")
                    return temp
                except ValueError:
                    continue

        for word in idlist:
            temp = word + '身份证校验'
            word = word.strip()
            if word in df.columns:
                df[temp] = df[word].apply(idCardVerification)

        for word in nulllist:
            temp = word + '空值校验'
            word = word.strip()
            if word in df.columns:
                df[temp] = df[word].apply(nullVefify)

        for word in datelist:
            if len(datelist) == 0:
                break
            # temp = word + '时间统一'
            word = word.strip()
            if word in df.columns:
                df[word] = df[word].apply(dateformated)
        # print(df)
        filename = os.path.split(sourthpath)[-1]
        dir = os.path.split(sourthpath)[0]
        newobjectpath = dir +'/' + filename.split('.', 1)[0] + '校验结果.xlsx'
        df.to_excel(newobjectpath)
        messagebox.showinfo("message", "校验完成\n输出路径为:%s" % (newobjectpath))


if __name__ == '__main__':
    # GUI 测试
    # controller = JSVerifyController()
    # tkinter.mainloop()

    print(JSExcelHandler.checkIDNumber('51052119810109281X'))
    print(JSExcelHandler.checkIDNumber('61052419911020004X'))










