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
import shutil
import time
import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import pandas as pd
import numpy as np
from openpyxl import load_workbook

from Delegate.ConfigureDelegate import ConfigureDelegate
from Model.JSCellModel import JSCellModel
from Tools.JSExcelHandler import JSExcelHandler
from View.JSConfigureView import JSConfigureView
from View.JSETLWindow import JSETLWindow


class JSETLController(ConfigureDelegate):
    def __init__(self):
        self.etlwindow = JSETLWindow()
        # self.etlwindow.gui_arrang()
        self.etlwindow.delegate = self

    # 合并两个表中差异字段 TODO
    def mergeDataFrame(self):
        print("开始合并数据")

    # 自动根据数据提取表头
    def autoExtractSheetHead(self, configuremodel):
        print("开始抽取模板表头")
        datapath = configuremodel.datapath
        # 所有需要合并的表格数据
        exceldatalist = JSExcelHandler.getPathFromRootFolder(datapath)
        dflist = []
        for path in exceldatalist:
            try:
                readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(path)
                for sheetname in sheetnames:
                    rSheet = readOpenXlsx.sheet_by_name(sheetname)
                    mergecells = rSheet.merged_cells
                    # 默认第一行为表头
                    # 没有合并单元格的情况
                    if len(mergecells) == 0:
                        headrow = 1
                    else:
                        levelmap = self.createLevelMap(mergecells, rSheet)
                        # 根据合并单元格的位置获取表头的范围
                        # 注意 这里和降范式的 最后行判断是不一样的,因为这个是整个表头提取,而降范式只要层级的根节点的 level 作为行数来提取数据
                        maxlevel = max(levelmap.keys())
                        lowestnodes = levelmap[maxlevel]
                        rlow, rhign, clow, chigh = cellrange = lowestnodes[0].cellrange
                        headrow = rhign + 1
                    curdf = pd.read_excel(path, sheetname, nrows=headrow, header=None)
                    # 每个进行遍历对比,完全不同则添加到 list 中
                    flag = True
                    for olddf in dflist:
                        if curdf.equals(olddf):
                            flag = False
                            pass
                    # 第一次
                    if len(dflist) == 0:
                        flag = True
                    if flag:
                        dflist.append(curdf)
                        filename, suffix = JSExcelHandler.SplitPathReturnNameAndSuffix(path)
                        # 这里要保存为 xlsx 为了兼容合并单元格的功能
                        objectpath = "{}/{}_{}.{}".format(configuremodel.headspath, filename, sheetname, 'xlsx')
                        self.restoreHead(path, sheetname, objectpath, curdf, headrow)
            except Exception as e:
                print(str(e))
                JSExcelHandler.errorlog("自动根据数据提取表头,原数据文件有问题-{}".format(path))
                pass
        print("抽取模板表头完成")

    # 根据 merge cell 定位多范式表格的位置, 以最小range 作为标准, 如果没有 merge cell 则默认第一行为表头
    # 应该将叶子节点的数据向根节点拼接,最后降范式的值从根节点层按列 + 1的顺序得出
    def lowerDimensionOfTitle(self, path, startrow=0):
        readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(path)
        rSheet = readOpenXlsx.sheet_by_name(sheetnames[0])
        valuelist = []
        keywords= []
        lastrow = 0
        # 表格的列边界
        ecol = rSheet.ncols
        resultCells = rSheet.merged_cells
        # mergeCellCount = len(mergedCells)
        mergedCells = []
        # 这里要加从第几行开始降范式,默认是从 0 行开始, 默认是从 0 行开始
        if startrow > 0:
            for index, cell in enumerate(resultCells):
                crlow, crhign, cclow, cchigh = cell
                if crlow >= startrow:
                    mergedCells.append(cell)
        else:
            mergedCells = resultCells
        mergeCellCount = len(mergedCells)
        # 没有 mergeCell 的情况, 默认第一行为表头
        if mergeCellCount == 0:
            for colindex in range(0, ecol):
                value = rSheet.cell_value(0, colindex)
                valuelist.append(value)
                keywords.append(value)
            # print("判断--" + rSheet.name + "--的合并单元格已结束")
            # 参数返回：范式标识（False为一范式，True为多范式）、表头值、数据行坐标
            return valuelist, lastrow, keywords
        else:
            levelmap = self.createLevelMap(mergedCells, rSheet)
            maxlevel = max(levelmap.keys())
            dummy = self.createLevelTree(levelmap, rSheet)
            # TODO 根据层级来生成最终的 valuelist
            self.travel(dummy, valuelist, keywords, maxlevel)
            print(keywords)
            lastrow = max(levelmap.keys()) + 1
            return valuelist, lastrow, keywords

    def createLevelMap(self, mergedCells, rSheet):
        # 用于标记表头的最后一行位置
        levellist = []
        setlist = set()
        levelmap = {}
        # 最终降范式后的值
        valuelist = []
        # 这里是如果是 merge cell 的情况, 暂时定 mergecell 范围 > 6 的时候为标题的情况, 所以也要筛除这种情况 TODO
        # 这里是从上往下遍历的 mergecell,所以 mergecells 中 从左到右 row 依次增加,可以利用这点简化计算
        for index, cell in enumerate(mergedCells):
            # 这里要根据 cell 的row 进行分级
            # 根据 sheet 中合并单元格的属性转换为树的节点
            Cmodel = self.createCellModelAccordingRange(cell, rSheet)
            # 去除空值, x 范围一个以上
            if Cmodel.value is not None and Cmodel.value != '':
                levellist.append(Cmodel)
                # 为了找出所有的层级
                setlist.add(Cmodel.level)

        for value in setlist:
            levelmap[value] = []
        for index, cellmodel in enumerate(levellist):
            levelmap[cellmodel.level].append(cellmodel)
        return levelmap

    # 根据sheet 和 遍历的 levelmap 创建节点树 TODO 是降范式的关键步骤
    def createLevelTree(self, levelmap, rSheet):
        keys = levelmap.keys()
        # 最底层的层级
        lastlevel = max(keys)
        minlevel = min(keys)
        # 创建节点树
        # 伪头结点,方便计算
        dummy = JSCellModel()
        dummy.child = levelmap[minlevel]
        # 按层级遍历,将子节点的 x 轴范围数据父节点 x 轴范围的情况进行连接
        for curlevel in range(minlevel, lastlevel):
            currentcells = levelmap[curlevel]
            # 每层按照起始 col 进行排序 TODO
            for i, curcell in enumerate(currentcells):
                crlow, crhign, cclow, cchigh = curcell.cellrange
                nextlevel = curlevel + 1
                nextlevelcells = levelmap[nextlevel]
                for ncell in nextlevelcells:
                    nrlow, nrhign, nclow, nchigh = ncell.cellrange
                    if nclow >= cclow and nchigh <= cchigh:
                        curcell.child.append(ncell)
                        # 判断当前是不是最后一层, 且 range > 0的情况
                        if nchigh - nclow > 0 and nrhign - 1 == lastlevel:
                            for i in range(nclow, nchigh):
                                cellrange = lastlevel + 1, lastlevel + 2, i, i + 1
                                # print(rSheet.cell_value(lastlevel + 1, i))
                                cmodel = self.createCellModelAccordingRange(cellrange, rSheet)
                                ncell.child.append(cmodel)

                # 出现跨层的情况, 如一个单元格 占两层,但是不是每一层都有子节点的情况要单独处理
                # 注意 这里 层级取得单元格的 rlow, 这里要和 rhigh - 1 进行对比
                if len(curcell.child) == 0 and crhign - 1 == lastlevel and cchigh - cclow > 0:
                    for i in range(cclow, cchigh):
                        cellrange = lastlevel + 1, lastlevel + 2, i, i + 1
                        print(rSheet.cell_value(lastlevel + 1, i))
                        cmodel = self.createCellModelAccordingRange(cellrange, rSheet)
                        curcell.child.append(cmodel)
                # 每一层级的 node 进行从左到右的排序
                if len(curcell.child) > 1:
                    childs = curcell.child
                    childs = sorted(childs, key=lambda node: node.cellrange[2])
                    curcell.child = childs
        return dummy

    # 将 sheet 中的单元格转换为 cellmodel
    def createCellModelAccordingRange(self, cellrange, rSheet):
        rlow, rhign, clow, chigh = cellrange
        cellvalue = rSheet.cell(rlow, clow).value
        # self, cellrange=None, value='', level = 1, child=[]
        level = rlow
        # 注意这里要新建,不然会引用同一个 list
        child = []
        Cmodel = JSCellModel(cellrange, cellvalue, level, child)
        return Cmodel

    # 树遍历, list 用于接收返回的降范式表头
    def travel(self, nodes, valuelist=[],keywords=[], combinestr='', lv=1):
        if nodes.value == '':
            combinestr = nodes.value
        else:
            combinestr = "{}/{}".format(combinestr, nodes.value)
        # 用于判断是否是最后一层
        flag = True
        for node in nodes.child:
            flag = flag & (len(node.child) == 0)
        # 如果是最后一层
        if flag:
            if len(nodes.child) == 0:
                combinestr = combinestr.lstrip('/')
                valuelist.append(combinestr)
                keywords.append(nodes.value)
                return
            for nextlevel in nodes.child:
                temp = "{}/{}".format(combinestr, nextlevel.value)
                # 由 dummy带来的空字符会在头部多一个/, 剔除
                temp = temp.lstrip('/')
                valuelist.append(temp)
                keywords.append(nextlevel.value)
            return
        else:
            for node in nodes.child:
                self.travel(node, valuelist,keywords, combinestr, lv)

    # 读取数据表格进行比对后做合并操作 -- 根据表头进行比对来抽取数据
    def ReadDataThenCompareAndExtract(self, configuremodel):
        print("开始抽取合并数据")
        # 存放表头的路径
        datapath = configuremodel.datapath
        # 读取所有数据文件
        datafilelist = JSExcelHandler.getPathFromRootFolder(datapath)
        # 存放合并结果的路径
        resultfilepath = configuremodel.storepath
        # 类型是 map {路径 : df}, 这里是按照设置的起始行读取标准模板表头,默认是 0
        headsmaplist = self.readStandardHeadFromFolder(configuremodel)
        for dfpath in headsmaplist.keys():
            # 拼接结果文件路径, 用 os.path win 和 linux 会有自己判断
            filename = os.path.split(dfpath)[-1]
            startrow = 0
            rangemap = configuremodel.validrange
            if rangemap != {}:
                if filename in configuremodel.validrange:
                    startrow = int(configuremodel.validrange[filename])
            newfile = os.path.join(resultfilepath,filename)

            '''
            这里是为所有的标准模板预先建立数据抽取的结果文件
            '''
            try:
                # 降低表头的范式
                valuelist, lastrows, keywords = self.lowerDimensionOfTitle(dfpath, startrow)
                # TODO 如果有title 的情况,要从 0 到 startrow 的 title 抽取单独写到目标文件里(暂时不做)
                # 创建文件及表头文件到result文件夹下为抽取该模板做准备
                newfile, newsheetname = self.newfilesave(dfpath, newfile, valuelist, startrow)
            except Exception as e:
                JSExcelHandler.errorlog("降范式并且新建表头文件未抽取数据做准备-{}".format(dfpath))
                pass
            # 遍历目标数据文件下所有 sheet 是否与标准模板匹配,如果匹配则进行数据抽取合并操作
            for path in datafilelist:
                try:
                    readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(path)
                    for sheetname in sheetnames:
                        # print("校验模板表：" + str(dfpath) + "\n与数据表：" + path + '下工作表sheetname：' + sheetname + "对比")
                        # 获取模板表头的行数,用于数据表中获取表头范围
                        totalrows = len(headsmaplist[dfpath].index)
                        # 根据标准表头获取数据里的表头进行比对
                        datadf = pd.read_excel(path, sheet_name=sheetname, nrows=totalrows, skiprows=startrow)

                        if datadf.equals(headsmaplist[dfpath]):
                            # 抽取模块headsmaplist
                            df = pd.read_excel(path, sheetname, skiprows=lastrows)
                            self.apendDataFrame(df, newfile, newsheetname, False)
                except Exception as e:
                    JSExcelHandler.errorlog("遍历目标数据文件下所有 sheet 与标准模板匹配后追加数据-{}".format(path))
        print("合并数据完成")

    # 获取标准的表头的文件,建立标准表头格式的 maplist
    def readStandardHeadFromFolder(self, configuremodel):
        headspath = configuremodel.headspath
        excellist = JSExcelHandler.getPathFromRootFolder(headspath)
        dfmaplist = {}
        dflist = list(dfmaplist.keys())
        # 从标准头路径中读取标准表头的格式,为比对做准备
        for excelpath in excellist:
            filename = os.path.split(excelpath)[-1]
            startrow = 0
            if filename in configuremodel.validrange:
                startrow = int(configuremodel.validrange[filename])
            try:
                newdf = pd.read_excel(str(excelpath), skiprows=startrow)
            except Exception as e:
                JSExcelHandler.errorlog("获取标准表头-{}".format(excelpath))
                pass
            if len(dflist) > 0:
                for olddf in dflist:
                    if newdf.equals(olddf) == True:
                        break
                    else:
                        dfmaplist[excelpath] = newdf
                        dflist = list(dfmaplist.keys())
            else:
                dfmaplist[excelpath] = newdf
                dflist = list(dfmaplist.values())
        return dfmaplist

    # 追加的 dataframe, 目标文件路径, 目标文件 sheet, 是否是 header 类型
    def apendDataFrame(self, apenddf, resultfile, resultsheet, header):
        # 写入的文件dataframe，engine：以操作工具包执行、mode必须为读状态否者会新增sheet
        writer = pd.ExcelWriter(resultfile, engine='openpyxl')
        book = load_workbook(resultfile)
        writer.book = book
        # 创建 sheets
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        # 获取追加行数
        try:
            dfNum = pd.DataFrame(pd.read_excel(resultfile, sheet_name=resultsheet))
        except Exception as e:
            JSExcelHandler.errorlog("追加合并数据-{}".format(resultfile))
        newRowsNum = dfNum.shape[0] + 1
        # 写入文件
        apenddf.to_excel(excel_writer=writer, sheet_name=resultsheet, index=False, header=header, startrow=newRowsNum)
        writer.save()

    # 新建合并文件
    # oldfile 模板表头文件路径, newfile 合并后数据文件路径
    # 增加是否保留检索的 header
    def newfilesave(self, oldfile, newfile, valuelist, startrow=0, headers=False):
        readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(oldfile)
        if startrow > 0:
            headers = True
        if headers:
            title = pd.read_excel(oldfile, header=None, nrows=startrow)
            xlwter, sheetnames = self.restoreHead(oldfile, sheetnames[0], newfile, title, startrow)
            df = pd.DataFrame(columns=valuelist)
            self.apendDataFrame(df, newfile, sheetnames[0], True)
        else:
            df = pd.DataFrame(columns=valuelist)
            df.to_excel(newfile, sheet_name=sheetnames[0], index=False)
        return newfile, sheetnames[0]

    # 按照 mergecell 的范围重写入表头并按原表合并单元格
    def restoreHead(self, oldfilepath, sheetname, newfile, title, startrow=0):
        readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(oldfilepath)
        writer = pd.ExcelWriter(newfile, engine='xlsxwriter')
        workbook = writer.book
        merge_format = workbook.add_format({'align': 'center'})
        title.to_excel(writer, sheet_name=sheetnames[0], index=False, header=False)
        rsheet = readOpenXlsx.sheet_by_name(sheetname)
        resultCells = rsheet.merged_cells
        mergedCells = []
        # 根据原表中的 cellrange 确定合并单元格的范围, 筛选出 header 部分有合并单元格的情况
        if startrow > 0:
            for index, cell in enumerate(resultCells):
                crlow, crhign, cclow, cchigh = cell
                if crlow < startrow:
                    mergedCells.append(cell)
        else:
            mergedCells = resultCells

        for cellrange in mergedCells:
            crlow, crhign, cclow, cchigh = cellrange
            if crlow < startrow:
                worksheet = writer.sheets[sheetnames[0]]
                value = rsheet.cell_value(crlow, cclow)
                """
                        Merge a range of cells.

                        Args:
                            first_row:    The first row of the cell range. (zero indexed).
                            first_col:    The first column of the cell range.
                            last_row:     The last row of the cell range. (zero indexed).  
                            last_col:     The last column of the cell range.
                            data:         Cell data.
                            cell_format:  Cell Format object.
                """
                # print("{}-{}-{}-{}-{}".format(crlow, cclow, crhign - 1, cchigh - 1, value))
                worksheet.merge_range(crlow, cclow, crhign - 1, cchigh - 1, value, merge_format)
        writer.save()
        return writer, sheetnames

if __name__ == '__main__':

    # GUI 测试
    controller = JSETLController()
    tkinter.mainloop()

    # 文本格式配置文件的测试
    # path = os.path.abspath('..') + "/configureFile.txt"  # 表示当前所处的文件夹上一级文件夹的绝对路径
    # View = JSConfigureView(path)
    # # view 层读取配置文件
    # View.readConfigureFile()
    # # Controller 调度执行读取模板表头文件 | 根据数据文件自动识别表头
    # controller = JSETLController()
    # # 根据标准表头进行比对和抽取合并数据
    # controller.ReadDataThenCompareAndExtract(View.configuremodel)
    # # print("ETL Finished")
    #
    # # 识别数据表中的表头,并抽取表头作为标准表头模板参考
    # controller.autoExtractSheetHead(View.configuremodel)


    # ##处理规则表头和多行表头情况测试
    #
    # # path = 'D:\\\mergeTable\\path\\heads'
    # path = '/Users/sun/Desktop/AUTOEtlTest/heads/'
    # list = JSExcelHandler.getPathFromRootFolder(path)
    # controller = JSETLController()
    # for path in list:
    #     readOpenXls, sheetnames, workpath = JSExcelHandler.OpenXls(path)
    #     for sheetname in sheetnames:
    #         rSheet = readOpenXls.sheet_by_name(sheetname)
    #         controller.lowerDimensionOfTitle(rSheet, 2)

    # controller.newfilesave('D:\\\mergeTable\\path\\result\\result.xlsx',['id', 'nmae'])
