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
import xlsxwriter
import pandas as pd
from openpyxl import load_workbook

from Controller.JSExcelController import JSExcelController
from Delegate.ConfigureDelegate import ConfigureDelegate
from Model.JSCellModel import JSCellModel
from Tools.JSExcelHandler import JSExcelHandler
from View.JSConfigureView import JSConfigureView


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
        # 这里自动识别表头参考之前范式降低过程中部分代码,思路依然是根据 mergecell 的层次来判断,到最低层 + 1 则为表头的整体范围
        # 这里涉及到一个问题, 对于合并单元格的读写问题


    # 根据 merge cell 定位多范式表格的位置, 以最小range 作为标准, 如果没有 merge cell 则默认第一行为表头
    # 应该将叶子节点的数据向根节点拼接,最后降范式的值从根节点层按列 + 1的顺序得出
    def lowerDimensionOfTitle(self, path, startrow = 0):
        readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(path)
        rSheet = readOpenXlsx.sheet_by_name(sheetnames[0])
        valuelist = []
        lastrow = 0
        # 表格的列边界
        ecol = rSheet.ncols
        resultCells = rSheet.merged_cells
        # mergeCellCount = len(mergedCells)
        mergedCells = []
        # TODO 这里要加从第几行开始降范式,默认是从 0 行开始, 默认是从 0 行开始
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
            # print(valuelist)
            print("判断--" + rSheet.name + "--的合并单元格已结束")
            # 参数返回：范式标识（False为一范式，True为多范式）、表头值、数据行坐标
            return valuelist, lastrow
        else:
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

            #最底层的层级
            lastlevel = max(setlist)
            minlevel = min(setlist)
            lastrow = lastlevel + 1
            for value in setlist:
                levelmap[value] = []
            for index, cellmodel in enumerate(levellist):
                levelmap[cellmodel.level].append(cellmodel)
            print(levelmap)

            # 创建节点树
            # 伪头结点,方便计算
            dummy = JSCellModel()
            dummy.child = levelmap[minlevel]
            # 按层级遍历,将子节点的 x 轴范围数据父节点 x 轴范围的情况进行连接
            for curlevel in range(minlevel, lastlevel):
                currentcells = levelmap[curlevel]
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
                            # print(rSheet.cell_value(lastlevel + 1, i))
                            cmodel = self.createCellModelAccordingRange(cellrange, rSheet)
                            curcell.child.append(cmodel)

            self.travel(dummy, valuelist)
            print(valuelist)
            return valuelist, lastrow

    # 将 sheet 中的单元格转换为 cellmodel
    def createCellModelAccordingRange(self, cellrange, rSheet):
        rlow, rhign, clow, chigh = cellrange
        rangex = chigh - clow
        cellvalue = rSheet.cell(rlow, clow).value
        # self, cellrange=None, value='', level = 1, child=[]
        level = rlow
        # 注意这里要新建,不然会引用同一个 list
        child = []
        Cmodel = JSCellModel(cellrange, cellvalue, level, child)
        return Cmodel

    # 树遍历, list 用于接收返回的降范式表头
    def travel(self, nodes, valuelist =[], combinestr=''):
        if len(nodes.child) == 0:
            valuelist.append(nodes.value)
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
            for nextlevel in nodes.child:
                temp = "{}/{}".format(combinestr, nextlevel.value)
                # 由 dummy带来的空字符会在头部多一个/, 剔除
                temp = temp.lstrip('/')
                valuelist.append(temp)
            return
        else:
            for node in nodes.child:
                self.travel(node, valuelist, combinestr)



    # 读取数据表格进行比对后做合并操作 -- 根据表头进行比对来抽取数据
    def ReadDataThenCompareAndExtract(self, configuremodel):
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
            newfile = resultfilepath + filename
            '''
            这里是为所有的标准模板预先建立数据抽取的结果文件
            '''
            # 降低表头的范式
            valuelist, lastrows = self.lowerDimensionOfTitle(dfpath, configuremodel.startrow)
            # TODO 如果有title 的情况,要从 0 到 startrow 的 title 抽取单独写到目标文件里(暂时不做)
            # 创建文件及表头文件到result文件夹下为抽取该模板做准备
            newfile, newsheetname = self.newfilesave(dfpath, newfile, valuelist, configuremodel.startrow)
            # 遍历目标数据文件下所有 sheet 是否与标准模板匹配,如果匹配则进行数据抽取合并操作
            for path in datafilelist:
                readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(path)
                for sheetname in sheetnames:
                    print("校验模板表：" + str(dfpath) + "\n与数据表：" + path + '下工作表sheetname：' + sheetname + "对比")
                    # 获取模板表头的行数,用于数据表中获取表头范围
                    totalrows = len(headsmaplist[dfpath].index)
                    # 根据标准表头获取数据里的表头进行比对
                    datadf = pd.read_excel(path, sheet_name=sheetname, nrows=totalrows, skiprows=configuremodel.startrow)
                    if datadf.equals(headsmaplist[dfpath]):
                        # 抽取模块
                        self.extractEqualValue(path, sheetname, newfile, newsheetname, lastrows)

    # 获取标准的表头的文件,建立标准表头格式的 maplist
    def readStandardHeadFromFolder(self, configuremodel):
        headspath = configuremodel.headspath
        excellist = JSExcelHandler.getPathFromRootFolder(headspath)
        startrow = configuremodel.startrow
        dfmaplist = {}
        dflist = list(dfmaplist.keys())
        # 从标准头路径中读取标准表头的格式,为比对做准备
        for excelpath in excellist:
            newdf = pd.read_excel(str(excelpath), skiprows= startrow)
            print(newdf)
            if len(dflist) > 0:
                for olddf in dflist:
                    if newdf.equals(olddf) == True:
                        print('is the same')
                        break
                    else:
                        dfmaplist[excelpath] = newdf
                        dflist = list(dfmaplist.keys())
            else:
                dfmaplist[excelpath] = newdf
                dflist = list(dfmaplist.values())
        return dfmaplist

    # 读取文件路径、读取文件工作簿、抽取写入的文件路径、抽取返回的工作簿、读取的文件的起始行 TODO 根据是否有 header 要求来判断
    def extractEqualValue(self, datafile, datasheet, resultfile, resultsheet, lastrows):
        df = pd.read_excel(datafile, sheet_name=datasheet, skiprows=lastrows)
        # 写入的文件dataframe，engine：以操作工具包执行、mode必须为读状态否者会新增sheet
        writer = pd.ExcelWriter(resultfile, engine='openpyxl', mode='w')
        book = load_workbook(resultfile)
        writer.book = book
        # 创建 sheets
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        # 获取追加行数
        dfNum = pd.DataFrame(pd.read_excel(resultfile, sheet_name=resultsheet))
        newRowsNum = dfNum.shape[0] + 1
        # 写入文件
        df.to_excel(excel_writer=writer, sheet_name=resultsheet, index=False, header=False, startrow=newRowsNum)
        writer.save()

    # 新建合并文件
    # oldfile 模板表头文件路径, newfile 合并后数据文件路径
    # 增加是否保留检索的 header
    def newfilesave(self, oldfile, newfile, valuelist, startrow=0, headers=False):
        readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(oldfile)
        if headers:
            title = pd.read_excel(oldfile, nrows=startrow - 1)
            title.to_excel(newfile, sheet_name=sheetnames[0], index=False, merge_cells=False)
            writer = pd.ExcelWriter(newfile, engine='openpyxl', mode='w')
            book = load_workbook(newfile)
            writer.book = book
            # 创建 sheets
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df = pd.DataFrame(columns=valuelist)
            df.to_excel(excel_writer=writer, sheet_name=sheetnames[0], index=False, startrow=startrow)
            writer.save()
        else:
            df = pd.DataFrame(columns=valuelist)
            readOpenXlsx, sheetnames, tempPath = JSExcelHandler().OpenXls(oldfile)
            df.to_excel(newfile, sheet_name=sheetnames[0], index=False)
        return newfile, sheetnames[0]

    # TODO 合并单元格
    def writeMergeCellsWithCellRange(self):
        df = pd.DataFrame(pd.np.random.randn(10, 3), columns=list('ABC'))
        column_list = df
        # Create a Pandas Excel writer using XlsxWriter engine.
        writer = pd.ExcelWriter("test.xlsx", engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', startrow=2, header=False, index=False)
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_fmt = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'color': 'white', 'fg_color': '#00007f', 'bold': True, 'border': 1})
        merge_format = workbook.add_format({'align': 'center'})
        worksheet.merge_range('A2:C2', 'Sample', merge_format)
        for idx, val in enumerate(column_list):
            worksheet.write(0, idx, val, header_fmt)
            font_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 9})
            worksheet.set_column('A:C', None, font_fmt)
            worksheet.set_row(0, None, header_fmt)
            writer.save()


if __name__ == '__main__':
    # 配置文件路径
    path = os.path.abspath('..') + "/configureFile.txt"  # 表示当前所处的文件夹上一级文件夹的绝对路径
    controller = JSETLController()
    View = JSConfigureView(path)
    # view 层读取配置文件
    View.readConfigureFile()
    # # Controller 调度执行读取模板表头文件 | 根据数据文件自动识别表头
    controller = JSETLController()
    # controller.writeMergeCellsWithCellRange()
    # 根据标准表头进行比对和抽取合并数据
    controller.ReadDataThenCompareAndExtract(View.configuremodel)


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
