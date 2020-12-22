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

    def __init__(self, rootfolder):
        # 用于表计数
        self.rowCount = 0
        # 原始数据路径目录
        self.rootFolder = rootfolder

        # 生成梳理统计 xls 的名字
        # self.writeWorkBookName = writeBookname
        self.writeOpenXlsx = xlwt.Workbook()

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



