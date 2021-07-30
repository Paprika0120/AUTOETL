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
    list = JSExcelHandler.getPathFromRootFolder(r'C:\Users\sjj-001\Desktop\AUTOETLTEST\data')
    reslist = []
    for path in list:
        df = pd.read_excel(path, sheet_name='YB01',skiprows=3)
        df = df.fillna('')
        sum = []
        for index , row in df.iterrows():
            str = row[1]
            list = [str, row[2], str.count(' ')]
            sum.append(list)
        for index , row in df.iterrows():
            str = row[4]
            list = [str, row[5], str.count(' ')]
            sum.append(list)
        ttt = pd.DataFrame(sum, columns=['科目名称', '数值', 'count'])
        reslist.append(ttt)


    resdf = pd.concat(reslist)
    resdf.reset_index(drop=True, inplace=True)

    prestr = None
    preprestr = ''
    precount = 0
    for index, row in resdf.iterrows():
        print(index)
        count = row[2]
        if count == 0:
            prestr = row[0]
            precount = count
            continue
        else:
            if count > precount:
                preprestr = prestr
                str = prestr + '/' + row[0].replace(' ', '')
                resdf.iloc[index, 0] = str
                prestr = str
                precount = count
                # print(prestr)
            else:
                if count - precount < 0:
                    str = row[0].replace(' ', '')
                    resdf.iloc[index, 0] = str
                    prestr = str
                    preprestr = prestr
                    precount = count
                    continue
                else:
                    if count == precount:
                        str = preprestr + '/' + row[0].replace(' ', '')
                        resdf.iloc[index, 0] = str
        resdf.to_excel(r'C:\Users\sjj-001\Desktop\ttt.xls')

    # print(resdf)
    print("执行完成")

