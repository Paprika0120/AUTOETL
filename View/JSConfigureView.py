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


class JSConfigureView(ConfigureDelegate):

    def __init__(self, filepath):
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
        f = open(filepath)  # 返回一个文件对象
        lines = f.readlines()  # 调用文件的 readline()方法
        varmap = {}
        for index, line in enumerate(lines):
            if index == 0:
                varmap['startrow'] = line.strip()
            elif index == 1:
                varmap['endrow'] = line.strip()
            elif index == 2:
                varmap['keywords'] = line.strip().split(',')
            else:
                varmap['storepath'] = line.strip()
        f.close()
        return varmap

if __name__ == '__main__':
    view = JSConfigureView("/Users/sun/Desktop/Development/Python_Demos/AutoETL/test.txt")
    view.readConfigureFile()







