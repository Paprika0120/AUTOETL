import os
import tkinter
from tkinter import *
from tkinter import ttk

from Controller.JSETLController import JSETLController
from Controller.JSVerifyController import JSVerifyController
from Model.JSConfigureModel import JSConfigureModel
from Tools.JSExcelHandler import JSExcelHandler


class JSETLWindow:
    def __init__(self, master, controller):
        self.delegate = controller

        my_path = os.path.dirname(os.path.realpath(sys.argv[0]))
        ETLFile = os.path.join(my_path, 'ETLFile')
        data = os.path.join(ETLFile, 'data')

        heads = os.path.join(ETLFile, 'heads')
        result = os.path.join(ETLFile, 'result')

        # 主窗口
        self.root = master
        self.root.title("批量数据处理工具")
        self.verifyWindow = None

        self.master = self.root
        # 基准界面initface
        self.initface = tkinter.Frame(self.master, )
        self.initface.pack()

        self.datatext = tkinter.Label(self.initface, text='原始数据位置', bd=4)
        self.datatext.grid(row=0, column=1)

        self.datalabel = tkinter.Entry(self.initface, width=80)
        self.datalabel.grid(row=0, column=2)
        self.datalabel.insert(END, data)

        self.headstext = tkinter.Label(self.initface, text='模板表头位置', bd=4)
        self.headstext.grid(row=1, column=1)
        self.headslabel = tkinter.Entry(self.initface, width=80)
        self.headslabel.grid(row=1, column=2)
        self.headslabel.insert(END, heads)

        self.resulttext = tkinter.Label(self.initface, text='合并后数据存放位置', bd=4)
        self.resulttext.grid(row=2, column=1)
        self.resultlabel = tkinter.Entry(self.initface, width=80)
        self.resultlabel.grid(row=2, column=2)
        self.resultlabel.insert(END, result)

        self.validrangetext = tkinter.Label(self.initface, text='表头起始位置', bd=4)
        self.validrangetext.grid(row=3, column=1)
        self.validrangelabel = tkinter.Entry(self.initface, width=80)
        self.validrangelabel.grid(row=3, column=2)
        self.validrangelabel.insert(END, '示例模板表头.xlsx:3')

        # 创建根据数据抽取表头的按钮
        self.headExtractButton = tkinter.Button(self.initface, command=self.headExtractButtonHandler, text="抽取表头")
        self.headExtractButton.grid(row=4, column=2)
        # 创建根据标准表头进行合并数据的按钮
        self.ETLButton = tkinter.Button(self.initface, command=self.ETLButtonHandler, text="执行合并")
        self.ETLButton.grid(row=5, column=2)

        # 跳转到校验界面
        self.ChangeButton = tkinter.Button(self.initface, command=self.changeToVerifyWindow, text="校验界面")
        self.ChangeButton.grid(row=6, column=2)

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

    def __readFile(self):
        varmap = {}
        varmap['datapath'] = self.datalabel.get().strip()
        varmap['headspath'] = self.headslabel.get().strip()
        varmap['storepath'] = self.resultlabel.get().strip()
        varmap['validrange'] = self.transformStringToMap(self.validrangelabel.get().strip())
        return varmap

    # 将对应的 表头:有效范围 分开
    def transformStringToMap(self, text=''):
        map = {}
        if text is None or text == '':
            return map
        maplist = text.split(',')
        for item in maplist:
            temp = item.split(':')
            map[temp[0]] = temp[1]
        return map

    def changeToVerifyWindow(self):
        self.initface.destroy()
        self.delegate = None
        JSVerifyWindow(self.master, JSVerifyController())


class JSVerifyWindow:

    def __init__(self, master, controller):
        my_path = os.path.dirname(os.path.realpath(sys.argv[0]))
        my_path = os.path.join(my_path, 'ETLFile')
        self.result = os.path.join(my_path, 'result')

        self.delegate = controller
        self.master = master
        # self.master.config(bg='blue')
        self.face = tkinter.Frame(self.master, )
        self.face.pack()
        # 主窗口
        # 创建根据标准表头进行合并数据的按钮
        self.ETLButton = tkinter.Button(self.face, command=self.ETLButtonHandler, text="执行校验")
        self.ETLButton.grid(row=5, column=2)

        self.datatext = tkinter.Label(self.face, text='校验文件路径', bd=4)
        self.datatext.grid(row=0, column=1)

        # self.datalabel = tkinter.Entry(self.face, width=80)
        # self.datalabel.grid(row=0, column=2)
        # self.datalabel.insert(END, result)

        self.headstext = tkinter.Label(self.face, text='身份证校验字段', bd=4)
        self.headstext.grid(row=1, column=1)
        self.headslabel = tkinter.Entry(self.face, width=80)
        self.headslabel.grid(row=1, column=2)
        self.headslabel.insert(END, '身份证')

        self.validrangetext = tkinter.Label(self.face, text='空值校验字段', bd=4)
        self.validrangetext.grid(row=3, column=1)
        self.validrangelabel = tkinter.Entry(self.face, width=80)
        self.validrangelabel.grid(row=3, column=2)
        self.validrangelabel.insert(END, '单位,金额')

        self.datetext = tkinter.Label(self.face, text='统一时间列', bd=4)
        self.datetext.grid(row=4, column=1)
        self.datelabel = tkinter.Entry(self.face, width=80)
        self.datelabel.grid(row=4, column=2)
        self.datelabel.insert(END, '入职时间')

        self.btn_back = tkinter.Button(self.face, text='返回批处理界面', command=self.back)
        self.btn_back.grid(row=6, column=2)



        self.comboxlist = ttk.Combobox(self.face, width= 75, postcommand = self.pathCallback)  # 初始化
        self.comboxlist["values"] = JSExcelHandler.getPathFromRootFolder(self.result)
        # self.comboxlist.current(0)  # 选择第一个
        # self.comboxlist.bind("<<ComboboxSelect>>",callbackFunc)  # 绑定事件,(下拉列表框被选中时，绑定go()函数)
        self.comboxlist.grid(row=0, column=2)
        # ddl_get = self.ddb_acquire.get()
        # self.comboxlist = Label(self.win, text=str(ddl_get)

    def pathCallback(self):
        datalist = JSExcelHandler.getPathFromRootFolder(self.result)
        self.comboxlist["values"] = datalist

    def __readFile(self):
        sourthpath = self.comboxlist.get().strip()
        idlist = self.headslabel.get().strip().split(',', 1)
        nulllist = self.validrangelabel.get().strip().split(',', 1)
        datelist = self.datelabel.get().strip().split(',', 1)
        return sourthpath, idlist, nulllist, datelist

    def ETLButtonHandler(self):
        # 先验证路径的正确性
        sourthpath, idlist, nullslist, datelist = self.__readFile()
        if self.delegate:
            # idlist = ['身份证']
            # nullslist = ['单位', '金额']
            self.delegate.verifyColumn(sourthpath, idlist, nullslist, datelist)

    def back(self):
        self.face.destroy()
        self.delegate = None
        JSETLWindow(self.master, JSETLController())

if __name__ == '__main__':
    controller = JSETLController()
    etlview = JSETLWindow(tkinter.Tk(), controller)
    tkinter.mainloop()








