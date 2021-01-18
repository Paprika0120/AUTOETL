#!/usr/bin/env python
# coding=utf-8
# ========================================
#
# Time: 2020/12/18 10:49
# Author: Sun
# Software: PyCharm
# Description:
#
#
# ========================================
import os
import xlrd
import xlwt


class JSExcelHandler(object):

    # 获取 Excel 文件路径
    @classmethod
    def getPathFromRootFolder(cls, rootfolder):
        excelList = []
        # temp是当前目录下的文件名称
        for temp in os.listdir(rootfolder):
            # 拼接绝对路径D
            filepath = os.path.join(rootfolder, temp)
            if os.path.isdir(filepath):
                # 判断是否文件夹，如果是文件夹，则继续递归遍历
                cls.getPathFromRootFolder(filepath)
            else:
                (name, extension) = os.path.splitext(temp)
                extension = extension.lower()
                # 找到 excel 的类型进行读取表头操作
                if (extension == ".xls" or extension == ".xlsx"):
                    excelList.append(filepath)
        return excelList

    # 打开 workbook 读取 workbook 的基本信息
    @classmethod
    def OpenXls(cls, workpath):
        tempPath = workpath
        workpath = workpath.lower()
        readOpenXlsx = None
        # 拓展名要加 '.'
        if workpath.split('.xl')[1] == 's':
            # # 因为原数据有老格式xls 的文件,xlrd 包要兼容必须开启 formatting_info=True, 但是这样就不能支持 xlsx
            readOpenXlsx = xlrd.open_workbook(workpath, formatting_info=True, on_demand=True, encoding_override='utf-8')
        else:
            readOpenXlsx = xlrd.open_workbook(workpath, on_demand=True, encoding_override='utf-8')
        # 所有 sheet 的名字
        sheetnames = readOpenXlsx.sheet_names()
        # 返回 sheets 对象, 名字, 绝对路径
        return readOpenXlsx, sheetnames, tempPath

    # 新建 sheet 并保存, 标题按传入的参数数组依次写入
    @classmethod
    def addSheet(cls, titleArgs):
        # 按数据顺序来写入标题
        for title, index in enumerate(titleArgs):
            cls.writesheet.write(0, index, title)
        cls.saveWithName()
        return cls.writesheet

    # # 粒度: workbook -> sheet -> head, data 这里是返回 sheet 的基本信息,格式为 map
    # @classmethod
    # def getBasicInfoFromSheet(cls, workpath):
    #     readOpenXls, sheetnames, workpath = cls.OpenXls(workpath)
    #     # 写入 sheet 名字
    #     for name in sheetnames:
    #         # print(name)
    #         rSheet = readOpenXls.sheet_by_name(name)
    #         # 确定表头范围，返回表头拼接字段
    #         head = cls.handleOriginalXLSData(rSheet)

    # 创建目录
    @classmethod
    def mkdir(cls, path):
        # 去除首位空格
        path = path.strip()
        # 去除尾部 \ 符号
        path = path.rstrip("\\")
        # 判断路径是否存在
        isExists = os.path.exists(path)
        # 判断结果
        if not isExists:
            # 如果不存在则创建目录,创建目录操作函数
            '''
            os.mkdir(path)与os.makedirs(path)的区别是,当父目录不存在的时候os.mkdir(path)不会创建，os.makedirs(path)则会创建父目录
            '''
            # 此处路径最好使用utf-8解码，否则在磁盘中可能会出现乱码的情况
            os.makedirs(path.decode('utf-8'))
            print
            path + ' 创建成功'
            return True
        else:
            # 如果目录存在则不创建，并提示目录已存在
            print
            path + ' 目录已存在'
            return False

    @classmethod
    def readtxt(cls, filepath):
        with open(filepath) as f:
            lines = f.readlines()  # 调用文件的 readline()方法
        return lines

    @classmethod
    def errorlog(cls, text):
        with open('errorlog.txt', 'a+') as f:
            f.write('{}\n'.format(text))

    @classmethod
    def SplitPathReturnNameAndSuffix(self, path):
        filename = os.path.split(path)[-1].split('.')[0]
        suffix = os.path.split(path)[-1].split('.')[1]
        return filename, suffix
