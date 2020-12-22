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


class JSETLController(ConfigureDelegate):

    # def __init__(self, configurefile):
    #
    #     self.configureV = JSConfigureView(configurefile)
    #     self.configureV.readConfigureFile()

    # def handleSheetHead(self, configuremodel):
    #     # 启用配置文件的情况
    #     if self.configureM.configureAvailable == 'M':
    #         self.extractSheetHeadWithConfigureFile(configuremodel, configuremodel)
    #     # 自动搜寻表头
    #     else:
    #         self.autoExtractSheetHead(configuremodel)

    # 自动根据数据提取表头 -- 延后做 TODO
    def autoExtractSheetHead(self, configuremodel):
        storepath = configuremodel.storepath
        # 所有需要合并的表格数据
        exceldatalist = JSExcelHandler.getPathFromRootFolder(storepath)
        excelhandler = JSExcelController()
        excelhandler.getPathFromRootFolder()
        print("开始搜索表头")

    # 读取数据表格进行比对后做合并操作 -- 根据表头进行比对来抽取数据
    def ReadDataThenCompareAndExtract(self, configuremodel):
        # 存放表头的路径
        datapath = configuremodel.datapath
        # 读取所有数据文件
        datafilelist = JSExcelHandler.getPathFromRootFolder(datapath)
        # 类型是 map
        headsmaplist = self.readStandardHeadFromFolder(configuremodel)
        for path in datafilelist:
            for df in headsmaplist.values():
                print('hello')
                # 获取模板表头的行数,用于数据表中获取表头范围
                totalrows = len(df.index)
                # 根据标准表头获取数据里的表头进行比对
                datadf = pd.read_excel(path, nrows=totalrows)
                if datadf.equals(df):
                    print('is the same')

    # 获取标准的表头的文件,建立标准表头格式的 maplist
    def readStandardHeadFromFolder(self, configuremodel):
        headspath = configuremodel.headspath
        excellist = JSExcelHandler.getPathFromRootFolder(headspath)
        dfmaplist = {}
        # 从标准头路径中读取标准表头的格式,为比对做准备
        for index, excelpath in enumerate(excellist):
            newdf = pd.read_excel(str(excelpath))
            # '/Users/sun/Desktop/heads/head3.xlsx'
            if len(dfmaplist) > 0:
                for olddf in dfmaplist.values():
                    if newdf.equals(olddf) is True:
                        break
            else:
                # 以路径来作为唯一标识,因为路径是唯一的
                dfmaplist[excelpath] = newdf
        return dfmaplist


if __name__ == '__main__':
    # 配置文件路径
    path = os.path.abspath('..') + "/configureFile.txt"  # 表示当前所处的文件夹上一级文件夹的绝对路径
    controller = JSETLController()

    View = JSConfigureView(path)
    # view 层读取配置文件
    View.readConfigureFile()

    # Controller 调度执行读取模板表头文件 | 根据数据文件自动识别表头
    controller = JSETLController()
    # 根据标准表头进行比对和抽取合并数据
    controller.ReadDataThenCompareAndExtract(View.configuremodel)




