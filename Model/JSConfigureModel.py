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
        # 关键字
        self.keywords = vars['keywords']
        # 表头起始行
        self.startrow = vars['startrow']
        # 表头起始列
        self.endrow = vars['endrow']
        # 合表后存放的路径
        self.storepath = vars['storepath']
        # 是否启用配置文件
        # self.configureAvailable = configureAvailable