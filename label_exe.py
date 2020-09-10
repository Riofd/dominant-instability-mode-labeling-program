from tkinter import *
from tkinter.filedialog import askdirectory
from PIL import Image, ImageTk
import os
import xlwt
import xlrd
import openpyxl
# image_dir1 = r'D:\Sample_Figure\PF100\1 AC line 100 50      3      9     10 2 0.1.matdelta.jpg'
# image_dir2 = r'D:\Sample_Figure\PF100\1 AC line 100 50      3      9     10 2 0.1.matvol.jpg'
# excel_dir = r'C:\Users\zhangrunfeng\Desktop\label.xlsx'


class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.image1_x = 0
        self.image1_y = 50
        self.image2_x = 0
        self.image2_y = 400
        self.fig_info, self.fig_name = self.readFile()
        self.num_fig = len(self.fig_name) / 2
        self.init_window()
        self.excel_dir = self.labelPath()
        _, self.idx = self.getExcelinfo()
        self.load_fig(self.idx)

    def init_window(self):
        self.master.title("主导失稳模式样本标注程序v1.0")

        self.pack(fill=BOTH, expand=1)

        # 实例化一个Menu对象，这个在主窗体添加一个菜单
        menu = Menu(self.master)
        self.master.config(menu=menu)

        # 创建File菜单，下面有Save和Exit两个子菜单
        file = Menu(menu)
        file.add_command(label='样本文件夹', command=self.samplePath)
        file.add_command(label='标签文件夹', command=self.labelPath)
        file.add_command(label='退出程序', command=self.client_exit)
        menu.add_cascade(label='文件', menu=file)

        # 创建Edit菜单，下面有一个Undo菜单
        edit = Menu(menu)
        edit.add_command(label='Undo')
        edit.add_command(label='Show  Image', command=self.showImg)
        edit.add_command(label='Show  Text', command=self.showTxt)
        menu.add_cascade(label='编辑', menu=edit)
        ClickButton1 = Button(self, text='稳定', command=lambda: self.readExcel(0))
        ClickButton1.place(x=750, y=200, width=100)
        ClickButton2 = Button(self, text='功角失稳', command=lambda: self.readExcel(1))
        ClickButton2.place(x=750, y=400, width=100)
        ClickButton2 = Button(self, text='电压失稳', command=lambda: self.readExcel(2))
        ClickButton2.place(x=750, y=600, width=100)

    def client_exit(self):
        exit()

    def showImg(self, image_dir, image_x, image_y):
        load = Image.open(image_dir).resize((700, 350))
        render = ImageTk.PhotoImage(load)

        img = Label(self, image=render)
        img.image = render
        img.place(x=image_x, y=image_y)

    def showTxt(self, idx):
        text = Label(self, text=self.fig_name[idx])
        text.place(x=0, y=0)

    def selectPath(self):
        path = askdirectory()
        # path.set(path)
        return path

    def samplePath(self):
        self.fig_info, self.fig_name = self.readFile()
        self.idx = 0
        self.load_fig(self.idx)

    def labelPath(self):
        path = self.selectPath()
        excel_dir = path + '/label' + '.xlsx'
        return excel_dir

    def getExcelinfo(self):
        excel_dir = self.excel_dir
        file = xlrd.open_workbook(excel_dir)
        table = file.sheets()[0]  # 得到sheet页
        nrows = table.nrows  # 总行数
        return excel_dir, nrows

    def readExcel(self, label):
        excel_dir, nrows = self.getExcelinfo()
        wb = openpyxl.load_workbook(excel_dir)
        ws = wb['Sheet1']
        ws.cell(row=nrows + 1, column=1).value = label
        wb.save(excel_dir)
        self.load_fig(self.idx)

    def readFile(self):
        path = self.selectPath()
        fig_name = []
        for root, dirs, files in os.walk(path):
            for file in files:
                fig_name.append(root + '/' + file)
        return files, fig_name

    def load_fig_dir(self, idx):
        fig_name = self.fig_name
        image_dir1 = fig_name[2*idx]
        image_dir2 = fig_name[2*idx + 1]
        return image_dir1, image_dir2

    def load_fig(self, idx):
        # print(idx)
        image_dir1, image_dir2 = self.load_fig_dir(idx)
        self.showTxt(2*idx)
        print(image_dir1, image_dir2)
        self.showImg(image_dir1, self.image1_x, self.image1_y)
        self.showImg(image_dir2, self.image2_x, self.image2_y)
        self.idx += 1


root = Tk()
root.geometry("900x750")
app = Window(root)
root.mainloop()

