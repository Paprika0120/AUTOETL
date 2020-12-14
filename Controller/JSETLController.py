#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 2020/02/8 11:06
# Author: Sun
# Software: PyCharm
# Description:
# 负责整理 excel 抽取, 配置获取等操作的调配
#
# ========================================
from Delegate.ConfigureDelegate import ConfigureDelegate
from Model.JSConfigureModel import JSConfigureModel
from View.JSConfigureView import JSConfigureView


class JSETLController(ConfigureDelegate):

    def __init__(self, configurefile):
        self.configureV = JSConfigureView(self, configurefile)
        self.configureV.readConfigureFile()


    # 可以增加 model 的属性
    # def readConfigureFile(self, filepath):
    #     model = self.configureV.reloadView()
    #     return model

    def handleSheetHead(self, filepath):
        # 启用配置文件的情况
        if self.configureM.configureAvailable == 'M':
            self.extractSheetHeadWithConfigureFile(self.configureV.configuremodel, filepath)
        # 自动搜寻表头
        else:
            self.autoExtractSheetHead(filepath)


    # 自动搜索表头的位置
    def autoExtractSheetHead(self, filepath):
        print("开始搜索表头")


    # 根据配置文件抽取表头
    def extractSheetHeadWithConfigureFile(self,configure, filepath):
        # 这里要考虑有的不规则表格没有表头的情况,要特殊处理
        # 目前考虑是单独走异常标记,收工处理
        startrow = configure.startrow
        endrow = configure.endrow



if __name__ == '__main__':
    controller = JSETLController('/Users/sun/Desktop/Development/Python_Demos/AutoETL/test.txt')


