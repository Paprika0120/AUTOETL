#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 2020/12/29 17:56
# Author: Sun
# Software: PyCharm
# Description:
#
#
# ========================================
from collections import Counter
from typing import List, Union

text = ''

class JSCellModel:
    # 层级 cell
    def __init__(self, cellrange=None, value='', level = 0, child=[]):
        # 子范式
        self.child = child
        # cell 的范围
        self.cellrange = cellrange
        # 当前范式的值
        self.value = value
        self.level = level


if __name__ == '__main__':

    values = []
    dummy = JSCellModel()
    dummy.value = ''
    level1 = JSCellModel()
    level1.value = 'level1'
    dummy.child = [level1]
    level21 = JSCellModel()
    level21.value = 'level21'
    level22 = JSCellModel()
    level22.value = 'level22'
    level1.child = [level21, level22]
    level31 = JSCellModel()
    level31.value = 'leve31'
    level32 = JSCellModel()
    level32.value = 'level32'
    level21.child = [level31, level32]
    level22.child = [level31, level32]

    # 树后序遍历
    def travel(nodes, list, combinestr=''):
        combinestr = "{}/{}".format(combinestr, nodes.value)
        flag = True
        for node in nodes.child:
            flag = flag & (len(node.child) == 0)
        if flag:
            for nextlevel in nodes.child:
                temp = "{}/{}".format(combinestr, nextlevel.value)
                list.append(temp)
            return
        else:
            for node in nodes.child:
                travel(node, list, combinestr)
    list = []
    travel(dummy, list, dummy.value)
    print(list)









