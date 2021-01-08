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
        # 合表后存放结果的路径
        self.storepath = vars['storepath']
        # 抽取数据文件路径
        self.datapath = vars['datapath']
        # head 的存放路径
        self.headspath = vars['headspath']
        # 从表头的第几行开始计算,因为前几行可能有标题的情况
        self.startrow = int(vars['startrow'])
        # 如果 startrow > 0 是否保留表头
        self.startrow = int(vars['startrow'])
        # 是否启用配置文件
        # self.configureAvailable = configureAvailable