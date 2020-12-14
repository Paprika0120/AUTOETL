#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 2020/02/9 14:32
# Author: Sun
# Software: PyCharm
# Description:
# excel 相关路径搜寻,读写操作
#
# ========================================
import os

import xlrd
import xlwt


class JSExcelController(object):

    def __init__(self, rootfolder, writeBookname = "Excel梳理统计.xls"):
        # 用于表计数
        self.rowCount = 0
        # 原始数据路径目录
        self.rootFolder = rootfolder

        # 生成梳理统计 xls 的名字
        self.writeWorkBookName = writeBookname

        # 各类数据的路径集合初始化
        self.excelList = []
        self.writeOpenXlsx = xlwt.Workbook()
        self.writesheet = None


    def handleExcelData(self, excelList):
        for path in excelList:
            print(path)
            try:
                fh = open("Errorlog", "a+")
                self.writeExcelData(path)
            except Exception:
                fh.write(path + '\n')
            finally:
                fh.close()
                self.saveWithName()
            continue
        # 写完一个工作表后, 行计数归零
        self.__rowCount = 0
        print("Excel数据梳理完成\n")





    # 获取源文件路径并存储在各个 list
    def getPathFromRootFolder(self, rootfolder):
        # temp是当前目录下的文件名称
        for temp in os.listdir(rootfolder):
            # 拼接绝对路径D
            filepath = os.path.join(rootfolder, temp)
            if os.path.isdir(filepath):
                # 判断是否文件夹，如果是文件夹，则继续递归遍历
                self.getPathFromRootFolder(filepath)
            else:
                (name, extension) = os.path.splitext(temp)
                extension = extension.lower()
                # 找到 excel 的类型进行读取表头操作
                if (extension == ".xls" or extension == ".xlsx"):
                    self.excelList.append(filepath)

    # 打开 xls读取数据
    def __OpenXls(self, workpath):
        print(workpath)
        tempPath = workpath
        workpath = workpath.lower()

        readOpenXlsx = None
        # 拓展名要加 '.'
        if workpath.split('.xl')[1] == 's':
            # # 因为原数据有老格式xls 的文件,xlrd 包要兼容必须开启 formatting_info=True, 但是这样就不能支持 xlsx
            readOpenXlsx = xlrd.open_workbook(workpath, formatting_info=True, on_demand=True, encoding_override='utf-8')
        else:
            readOpenXlsx = xlrd.open_workbook(workpath, on_demand=True, encoding_override='utf-8')
        # # 所有 sheet 的名字
        sheetnames = readOpenXlsx.sheet_names()
        # # 返回sheet 对象, 名字, 绝对路径
        return readOpenXlsx, sheetnames, tempPath


    # 保存文件
    def saveWithName(self):
        self.writeOpenXlsx.save(self.__writeWorkBookName)


    # 初始化写入工作薄
    def addSheet(self, sheeType):
        # EXCEL 类型
        if (sheeType == 'EXCEL'):
            self.writesheet = self.writeOpenXlsx.add_sheet(sheeType, cell_overwrite_ok=True)
            self.writesheet.write(0, 0, "编号")
            self.writesheet.write(0, 1, "类型")
            self.writesheet.write(0, 2, "路径")
            self.writesheet.write(0, 3, "工作表")
            self.writesheet.write(0, 4, "单位")
            self.writesheet.write(0, 5, "年份")
            self.writesheet.write(0, 6, "数据类型")
            self.writesheet.write(0, 7, "数据内容")
            self.__rowCount = 1
            self.saveWithName()

        return self.writesheet


    def writeExcelData(self, workpath):
        readOpenXls, sheetnames, workpath = self.__OpenXls(workpath)
        # 写入 sheet 名字
        for name in sheetnames:
            #print(name)
            rSheet = readOpenXls.sheet_by_name(name)

            # 确定表头范围，返回表头拼接字段
            head = self.handleOriginalXLSData(rSheet)

    def handleOriginalXLSData(self, rSheet):
        startRow = 0
        startCol = 0
        # 表格的行边界
        endRow = rSheet.nrows

        # 表格的列边界
        endCol = rSheet.ncols

        # 获取到 merged cell 的信息
        # /Users/sun/Desktop/test/test.xlsx
        # [(0, 3, 0, 1), (0, 1, 1, 5)]
        # 前两位 代表行的合并范围,后两位代表列的合并范围
        # (3 - 0) x (1 - 0) 的一个合并单元格
        # (1 - 0) x (5 - 1) 的一个合并单元格

        mergedCells = rSheet.merged_cells
        mergeCellCount = len(mergedCells)

        rowMax = 0
        # 处理逻辑:
        # 1.有多个合并单元格的情况,计算每个合并单元格的行边界,如果不同,说明是多级表头,取最大的为范围边界
        # 2.多级并且有平行表头的情况,这种情况难判断,先多取一行看数据,在用里面的字段来约束
        # 3.没有多级直接取第一行为表头

        # 有合并单元格的情况
        if (mergeCellCount >= 1):
            for index, cell in enumerate(mergedCells):
                rlo, rhi, clo, chi = cell
                rowMax = rhi if rowMax < rhi else rowMax
            # for cell in mergedCells:
            #     rlo, rhi, clo, chi = cell
            #     rowMax = rhi if rowMax < rhi else rowMax
            endRow = rowMax + 1
        else:
            endRow = 1

        # 数据中会有合并单元格的情况
        if (endRow >= 20):
            endRow = 20

        # 超过了边界
        if (endRow > rSheet.nrows):
            endRow = rSheet.nrows

        # start = datetime.datetime.now()
        head = ''
        for row in range(startRow, endRow):
            # 拼接表头
            for col in range(startCol, endCol):
                temp = rSheet.cell_value(row, col)
                if (temp != ''):
                    temp = str(temp).replace(' ', '')
                    head = head + temp
                    head = head + ","
                if (len(head) > 32575):
                    return "超长表头"
        return head


if __name__ == '__main__':
    path = "/Users/sun/Desktop/test/"
    handler = JSExcelController(path)
    handler.getPathFromRootFolder(path)
    handler.handleExcelData(handler.excelList)



