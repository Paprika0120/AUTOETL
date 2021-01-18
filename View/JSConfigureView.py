#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 2020/02/8 11:05
# Author: Sun
# Software: PyCharm
# Description:
# 负责读取配置文件,生成参数 model
#
# ========================================
from Delegate.ConfigureDelegate import ConfigureDelegate
from Model.JSConfigureModel import JSConfigureModel
from Tools.JSExcelHandler import  JSExcelHandler


class JSConfigureView(ConfigureDelegate):

    def __init__(self, filepath=''):
        self.filepath = filepath
        self.configuremodel = None

    def readConfigureFile(self):
        # 根据路径读取配置
        if(self.filepath is None):
            assert "配置文件路径不能为空"
        else:
            varsmap = self.__readFile(self.filepath)
            self.configuremodel = JSConfigureModel(varsmap)
            print("读取配置文件完毕")

    def __readFile(self,filepath):
        lines = JSExcelHandler.readtxt(filepath)
        varmap = {}
        for index, line in enumerate(lines):
            if index == 0:
                # 抽取文件的路径
                varmap['datapath'] = line.strip()
            elif index == 1:
                # 读取 heads 的路径
                varmap['headspath'] = line.strip()
            elif index == 2:
                # 最终合表结果存放路径
                varmap['storepath'] = line.strip()
            elif index == 3:
                # 进行比对的起始行
                varmap['validrange'] = self.transformStringToMap(line.strip())
            elif index == 4:
                # 进行比对的起始行
                varmap['reserveHead'] = line.strip()
            else:
                varmap['keywords'] = line.strip().split(',')
        return varmap

    # 将对应的 表头:有效范围 分开
    def transformStringToMap(self, text=''):
        print(text)
        map = {}
        maplist = text.split(',')
        for item in maplist:
            temp = item.split(':')
            map[temp[0]] = temp[1]
        return map

# if __name__ == '__main__':
#     print("hello world")
#     view = JSConfigureView()
#     str = '账套 1.xlsx:2,数据 6.xlsx:2'
#     view.transformStringToMap(str)











