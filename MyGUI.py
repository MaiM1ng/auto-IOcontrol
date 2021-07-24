# -*- coding: utf-8 -*-
"""
@time:2021/7/24 14:19
@Author:Yukino
"""


import pyautogui
import pyperclip
import tkinter as tk
from tkinter import filedialog
from pynput.keyboard import Controller
import time
import openpyxl as op


pyautogui.FAILSAFE = True

class MyGui(object):

    def __init__(self):

        self.root = tk.Tk()
        self.instLst = []
        self.isTopmostFlag = False  # 置顶标记
        self.isRunMinFlag = False  # 运行时最小化标记
        self.initUI()
        self.root.bind("<F1>",self.listen_F1)
        self.root.after(50, self.__updatePosition)
        self.root.mainloop()


    def initUI(self):
        self.root.title("键盘鼠标自动化")
        # 设置窗体尺寸 窗体居中
        width, height = 600, 250
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.root.geometry(alignstr)
        # 设置窗口大小不可改变
        self.root.resizable(width=False, height=False)
        # 控件初始化
        self.lbl_position = tk.Label(self.root, text='', font=('黑体', 15))
        self.lbl_position.place(x=20, y=20, width=450, height=35)

        self.btn_setTopmost = tk.Button(self.root, text='置顶',
                                        command=self.setTopmost)
        self.btn_setTopmost.place(x=490, y=20, width=80, height=35)

        self.txt_filePath = tk.Entry(self.root, text='', font=('宋体', 12))
        self.txt_filePath.place(x=20, y=70, width=450, height=35)

        self.btn_openFile = tk.Button(self.root, text='打开脚本',
                                      command=self.__openFile)
        self.btn_openFile.place(x=490, y=70, width=80, height=35)

        self.btn_load = tk.Button(self.root, text='载入脚本', font=('黑体', 12),
                                  command=self.loadFile)
        self.btn_load.place(x=20, y=180, width=100, height=50)

        self.btn_run = tk.Button(self.root, text='执行脚本',font=('黑体', 12),
                                 command=self.run)
        self.btn_run.place(x=140, y=180, width=100, height=50)

        self.btn_explain = tk.Button(self.root, text="脚本说明", font=('黑体', 12),
                                     command=self.explain)
        self.btn_explain.place(x=260, y=180, width=100, height=50)
        
        self.btn_setMin = tk.Button(self.root, text='前台运行', font=('黑体', 12),
                                    command=self.setMin)
        self.btn_setMin.place(x=400, y=180, width=170, height=50)

        self.lbl_current_file = tk.Label(self.root, text='载入的文件:', font=('黑体', 12))
        self.lbl_current_file.place(x=20, y=125, width=550, height=40)


    def getMousePosition(self):
        return pyautogui.position()

    def __getMousePositionString(self):
        x, y = pyautogui.position()
        positionStr = "X: " + str(x).rjust(4) + "   Y: " + str(y).rjust(4)
        return positionStr

    def setTopmost(self):
        if self.isTopmostFlag:
            self.isTopmostFlag = False
            self.root.wm_attributes('-topmost', 0)
            self.btn_setTopmost['text'] = '置顶'
        else:
            self.isTopmostFlag = True
            self.root.wm_attributes('-topmost', 1)
            self.btn_setTopmost['text'] = '取消置顶'

    def __updatePosition(self):
        self.lbl_position['text'] = self.__getMousePositionString()
        self.root.after(50, self.__updatePosition)

    def __openFile(self):
        fname = filedialog.askopenfilename(title='选择脚本文件',
                                           filetypes=[('text files', '.txt')])
        self.txt_filePath.delete(0, tk.END)
        self.txt_filePath.insert(0, fname)

    def loadFile(self):
        self.instLst.clear()
        fpath = ''
        fpath = self.txt_filePath.get()
        # print(fpath)
        if fpath != "":
            file = open(fpath, 'r', encoding='utf-8-sig')
            lines = file.readlines()
            print(lines)
            for line in lines:
                instStr = line.rstrip()
                if instStr == '':
                    continue
                # print(instStr)
                inst = Instruction(instStr)
                self.instLst.append(inst)
        currFileStr = "载入的文件:" + fpath
        self.lbl_current_file['text'] = currFileStr

    def run(self):
        try:
            if self.isRunMinFlag:
                self.root.state('iconic')
            for instruction in self.instLst:
                # print(instruction.toString())
                instruction.execute()
        except pyautogui.FailSafeException:
            print('停止执行')
            return


    def explain(self):
        ef = ExplainFrame()
    
    def setMin(self):
        # print(self.isRunMinFlag)
        if self.isRunMinFlag:
            self.isRunMinFlag = False
            self.btn_setMin['text'] = '前台运行'
        else:
            self.isRunMinFlag = True
            self.btn_setMin['text'] = '后台运行' 
            
    def listen_F1(self, event):
        pyperclip.copy("%d %d" % self.getMousePosition())


class ExplainFrame(object):

    def __init__(self):
        self.root = tk.Tk()
        self.root.title('脚本指令说明')
        # 设置窗体尺寸 窗体居中
        width, height = 600, 600
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.root.geometry(alignstr)
        # 设置窗口大小不可改变
        self.root.resizable(width=False, height=False)

        self.lbl_explain = tk.Label(self.root, text=self.getText(),font=('黑体', 15))
        self.lbl_explain.place(x=10, y=10, width=580, height=580)

    def getText(self):
        str = '简要说明:按下F1自动复制当前鼠标位置\n命令说明:\n' \
              'MOVE X Y(移动到XY)\n' \
              'CLICK (单击)\n' \
              'RIGHT (右击)\n' \
              'DOUBLECLICK (双击)\n' \
              'DELAY t (延时t秒)\n' \
              'KEY\n (键盘操作 最多支持两个同时按下)\n' \
              'INPUT\n (输入一句话)\n' \
              'EXCEL\n (打开一个EXCEL 路径从C盘开始\n 并用反斜杠”/“隔)\n' \
              'EXCELGET POS (获取单元格POS的内容)\n' \
              'PASTE (粘贴)\n' \
              'MID (中键点击)\n' \
              '# (不执行该行代码)'
        return str


class Instruction(object):

    def __init__(self, instStr):
        self.instStr = instStr
        self.excel = MyExcel()
        self.type = ''
        self.analyze()

    def analyze(self):
        self.instStrings = self.instStr.split(' ')
        self.type = self.instStrings[0].upper()
        if self.type[0] == '#':
            self.type = "#"
        # print(self.type)

    def execute(self):
        if self.type == 'MOVE':
            x = int(self.instStrings[1])
            y = int(self.instStrings[2])
            pyautogui.moveTo(x, y)
        elif self.type == 'DELAY':
            t = float(self.instStrings[1])
            time.sleep(t)
        elif self.type == 'CLICK':
            pyautogui.click()
        elif self.type == 'DOUBLECLICK':
            pyautogui.doubleClick()
        elif self.type == 'RIGHT':
            pyautogui.rightClick()
        elif self.type == 'MID':
            pyautogui.middleClick()
        elif self.type == 'EXCEL':
            # print(1)
            filePath = self.instStrings[1]
            self.excel.setPath(filePath)
            self.excel.load()
        elif self.type == 'PASTE':
            Str = pyperclip.paste()
            keyboard = Controller()
            # print(Str)
            keyboard.type(Str)
        elif self.type == 'EXCELGET':
            pos = self.instStrings[1]
            pyperclip.copy(self.excel.getText(pos))
        elif self.type == 'KEY':
            lens = len(self.instStrings[1].split('+'))
            if lens == 1:
                key = self.instStrings[1]
                pyautogui.press(key)
            elif lens > 1:
                keys = self.instStrings[1].split('+')
                key1, key2 = keys[0], keys[1]
                pyautogui.hotkey(key1, key2)
        elif self.type == 'INPUT':
            strings = " ".join(self.instStrings[1:])
            # print(strings)
            keyboard = Controller()
            keyboard.type(strings)

        elif self.type == "#":
            pass

    def __paste(self):
        pass

    def toString(self):
        if self.type == 'EXCELGET':
            return '1'


class MyExcel(object):

    path = ''
    excel = None
    ws = None

    def __init__(self):
        pass

    @classmethod
    def setPath(cls, path):
        cls.path = path

    @classmethod
    def load(cls):
        # print(self.path)
        cls.excel = op.load_workbook(cls.path)
        cls.ws = cls.excel.active

    @classmethod
    def getText(cls, pos):
        Str = cls.ws[pos].value
        return Str


if __name__ == '__main__':
    mg = MyGui()