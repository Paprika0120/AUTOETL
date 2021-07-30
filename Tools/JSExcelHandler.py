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
import pandas as pd

class JSExcelHandler(object):

    # 获取 Excel 文件路径
    @classmethod
    def getPathFromRootFolder(cls, rootfolder, excelList=[]):

        # temp是当前目录下的文件名称
        for temp in os.listdir(rootfolder):
            # 拼接绝对路径D
            filepath = os.path.join(rootfolder, temp)
            if os.path.isdir(filepath):
                # 判断是否文件夹，如果是文件夹，则继续递归遍历
                cls.getPathFromRootFolder(filepath, excelList)
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

    # 最早用于读取配置文件,现已做成 GUI 来读取配置文件
    @classmethod
    def readtxt(cls, filepath):
        with open(filepath) as f:
            lines = f.readlines()  # 调用文件的 readline()方法
        return lines

    @classmethod
    def errorlog(cls, text):
        with open('errorlog.txt', 'a+', encoding='UTF-8') as f:
            f.write('{}\n'.format(text))


    @classmethod
    def pathlog(cls, text):
        with open('pathlog.txt', 'a+', encoding='UTF-8') as f:
            f.write('{}\n'.format(text))

    @classmethod
    def excuteCheck(cls, text):
        with open('checklog.txt', 'a+', encoding='UTF-8') as f:
            f.write('{}\n'.format(text))

    # 返回文件名
    @classmethod
    def SplitPathReturnNameAndSuffix(cls, path):
        filename = os.path.split(path)[-1].split('.')[-2]
        suffix = os.path.split(path)[-1].split('.')[-1]
        return filename, suffix

    # 删除重复行
    @classmethod
    def dropDuplicatedData(cls, sourcepath, objectpath, startrow=4):
        excellist = cls.getPathFromRootFolder(sourcepath)
        sum = None
        for path in excellist:
            df = pd.read_excel(path, skiprows=startrow,header=None)
            if sum is None:
                sum = df
            else:
                sum = sum.append(df)

        # sum.drop_duplicates()
        sum = sum.drop_duplicates()
        objectpath += '/合并去重.xlsx'
        sum.to_excel(objectpath, index=False,header=None)

    # 身份证校验
    @classmethod
    def checkIDNumber(cls, num_str):
        str_to_int = {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5,
                      '6': 6, '7': 7, '8': 8, '9': 9, 'X': 10}
        check_dict = {0: '1', 1: '0', 2: 'X', 3: '9', 4: '8', 5: '7',
                      6: '6', 7: '5', 8: '4', 9: '3', 10: '2'}
        if len(num_str) != 18:
            # raise TypeError(u'请输入标准的第二代身份证号码')
            return False
        check_num = 0
        for index, num in enumerate(num_str):
            if index == 17:
                right_code = check_dict.get(check_num % 11)
                if num == right_code:
                    # print(u"身份证号: %s 校验通过" % num_str)
                    return True
                else:
                    # print(u"身份证号: %s 校验不通过, 正确尾号应该为：%s" % (num_str, right_code))
                    return False
            check_num += str_to_int.get(num) * (2 ** (17 - index) % 11)

    @classmethod
    def checkIsNull(cls, num_str):
        res = 0 if pd.isna(num_str) else 1
        return res

    # def score_verification(row):
    #     if not 0 <= row.SCORES <= 100:
    #         print(f'{row.ID}\t{row.NAME}\t{row.SCORES}')
    #
    # df = pd.read_excel(io='exls/data_verification.xlsx', sheet_name='Sheet1')
    # df.apply(score_verification, axis=1)






