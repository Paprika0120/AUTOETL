import shutil
import time
import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import sys
import os

from Model.JSConfigureModel import JSConfigureModel
from Tools.JSExcelHandler import JSExcelHandler


class JSETLWindow:
    def __init__(self):
        self.delegate = None
        # 主窗口
        self.root = tkinter.Tk()
        self.root.title("合表工具")

        self.datalabel = tkinter.Entry(self.root, width=80)
        self.headslabel = tkinter.Entry(self.root, width=80)
        self.resultlabel = tkinter.Entry(self.root, width=80)
        self.validrangelabel = tkinter.Entry(self.root, width=80)

        # 创建根据数据抽取表头的按钮
        self.headExtractButton = tkinter.Button(self.root, command=self.headExtractButtonHandler, text="抽取表头")
        # 创建根据标准表头进行合并数据的按钮
        self.ETLButton = tkinter.Button(self.root, command=self.ETLButtonHandler, text="执行合并")

    def headExtractButtonHandler(self):
        # 先验证路径的正确性
        varmap = self.__readFile()
        configuremodel = JSConfigureModel(varmap)
        if self.delegate:
            self.delegate.autoExtractSheetHead(configuremodel)

    def ETLButtonHandler(self):
        # 先验证路径的正确性
        varmap = self.__readFile()
        configuremodel = JSConfigureModel(varmap)
        self.delegate.ReadDataThenCompareAndExtract(configuremodel)

    def gui_arrang(self):
        self.datalabel.pack()
        self.headslabel.pack()
        self.resultlabel.pack()
        self.validrangelabel.pack()
        self.headExtractButton.pack()
        self.ETLButton.pack()

    def __readFile(self):
        varmap = {}
        varmap['datapath'] = self.datalabel.get().strip()
        varmap['headspath'] = self.headslabel.get().strip()
        varmap['storepath'] = self.resultlabel.get().strip()
        varmap['validrange'] = self.transformStringToMap(self.validrangelabel.get().strip())
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



