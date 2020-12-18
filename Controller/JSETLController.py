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
import os
from Model.JSConfigureModel import JSConfigureModel

from Controller.JSExcelController import JSExcelController
from Delegate.ConfigureDelegate import ConfigureDelegate
from View.JSConfigureView import JSConfigureView
from Tools.JSExcelHandler import JSExcelHandler
import pandas as pd
import numpy as np


class JSETLController(ConfigureDelegate):

    def __init__(self, configurefile):
        self.configureV = JSConfigureView(configurefile)
        self.configureV.readConfigureFile()

    # 可以增加 model 的属性
    # def readConfigureFile(self, filepath):
    #     model = self.configureV.reloadView()
    #     return model
    def handleSheetHead(self, configuremodel):
        # 启用配置文件的情况
        if self.configureM.configureAvailable == 'M':
            self.extractSheetHeadWithConfigureFile(configuremodel, configuremodel)
        # 自动搜寻表头
        else:
            self.autoExtractSheetHead(configuremodel)

    # 自动搜索表头的位置
    def autoExtractSheetHead(self, configuremodel):
        storepath = configuremodel.storepath
        # 所有需要合并的表格数据
        exceldatalist = JSExcelHandler.getPathFromRootFolder(storepath)
        excelhandler = JSExcelController()
        excelhandler.getPathFromRootFolder()
        print("开始搜索表头")

    # 根据配置文件抽取表头 -- 根据表头进行比对来抽取数据
    def readHeadsThenCompareAndStore(self, configuremodel):
        # 存放表头的路径
        headspath = configuremodel.headspath
        # 搜索所有 标准 heads 的 excel 文件
        headslist = JSExcelHandler.getPathFromRootFolder(headspath)
        headmaps = []
        for headpath in headslist:
            readOpenXls, sheetnames, workpath = JSExcelHandler.OpenXls(headpath)
            for name in sheetnames:
                # 按 sheet name 获取 workbook 中的 sheet
                rSheet = readOpenXls.sheet_by_name(name)
                map = {}
                # 之前已经有 map 要进行比对是否是同一个
                if len(headmaps) > 0:
                    print('hello world')
                else:
                    # 表头的列数
                    map['colnumber'] = rSheet.ncols
                    headmaps.append(map)

        # 根据路径读取后,存为 [map1, map2, map3]形式
        excelHandler = JSExcelController()
        excelHandler.getPathFromRootFolder()

if __name__ == '__main__':
    # path = os.path.abspath('..') + "/test.txt"  # 表示当前所处的文件夹上一级文件夹的绝对路径
    # controller = JSETLController(path)
    pathtest = '/Users/sun/Desktop/test'
    excellist = JSExcelHandler.getPathFromRootFolder(pathtest)

    df1 = pd.read_excel(excellist[0])
    df2 = pd.read_excel(excellist[1])
    print(df1)
    print(df2)


