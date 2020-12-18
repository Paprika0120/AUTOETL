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

    hello = []

    def __init__(self, vars):
        # 关键字
        self.keywords = vars['keywords']
        # 合表后存放结果的路径
        self.storepath = vars['storepath']
        # 抽取数据文件路径
        self.filepath = vars['filespath']
        # head 的存放路径
        self.headspath = vars['headspath']
        # 是否启用配置文件
        # self.configureAvailable = configureAvailable