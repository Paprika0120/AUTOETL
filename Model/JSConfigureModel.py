#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 2020/12/8 11:06
# Author: Sun
# Software: PyCharm
# Description:
#
#
# ========================================

class JSConfigureModel:

    def __init__(self, vars):

        # 合表后存放结果的路径
        self.storepath = vars['storepath']
        # 抽取数据文件路径
        self.datapath = vars['datapath']
        # head 的存放路径
        self.headspath = vars['headspath']
        # # 从表头的第几行开始计算,因为前几行可能有标题的情况
        self.validrange = vars['validrange']
        # # 是否保留表头
        # self.reserveHead = vars['reserveHead']
        # # 关键字
        # self.keywords = vars['keywords']
