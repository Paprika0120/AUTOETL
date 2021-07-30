#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 3/26/21 12:31 PM
# Author: Sun
# Software: PyCharm
# Description:
#
#
# ========================================
from Tools.JSExcelHandler import JSExcelHandler
import pandas as pd

if __name__ == '__main__':
    list = JSExcelHandler.getPathFromRootFolder(r'C:\Users\sjj-001\Desktop\2021年处理数据源集合\未处理\2020年市级收支月报')
    reslist = []
    for path in list:
        filename, suffix = JSExcelHandler.SplitPathReturnNameAndSuffix(path)
        df = pd.read_excel(path, sheet_name='YB01', skiprows=3)
        df = df.fillna('')
        sum = []
        for index , row in df.iterrows():
            str = row[1
            list = [row[0], str, row[2], filename]
            sum.append(list)
        for index , row in df.iterrows():
            str = row[4]
            list = [row[3], str, row[5], filename]
            sum.append(list)
        ttt = pd.DataFrame(sum, columns=['科目编码', '科目名称', '金额', '来源'])
        reslist.append(ttt)


    resdf = pd.concat(reslist)
    resdf.reset_index(drop=True, inplace=True)
    resdf.to_excel(r'C:\Users\sjj-001\Desktop\ttt.xls')


