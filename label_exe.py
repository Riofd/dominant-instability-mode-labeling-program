from tkinter import *
from tkinter.filedialog import askdirectory
from PIL import Image, ImageTk
import os
import glob
import time
import xlwt
import xlrd
import openpyxl


class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.image_x = 20
        self.image_y = 100
        # self.image2_x = 0
        # self.image2_y = 400
        self.fig_name = self.read_file()
        self.num_fig = len(self.fig_name) / 2
        self.init_window()
        self.excel_dir = self.label_path()
        _, self.idx = self.get_excel_info()
        self.load_fig(self.idx)

    def init_window(self):
        self.master.title("主导失稳模式样本标注程序v3.0")

        self.pack(fill=BOTH, expand=1)

        # 实例化一个Menu对象，这个在主窗体添加一个菜单
        menu = Menu(self.master)
        self.master.config(menu=menu)

        # 创建File菜单，下面有Save和Exit两个子菜单
        file = Menu(menu)
        file.add_command(label='样本文件夹', command=self.sample_path)
        file.add_command(label='标签文件夹', command=self.label_path)
        file.add_command(label='退出程序', command=self.client_exit)
        menu.add_cascade(label='文件', menu=file)

        # 创建Edit菜单，下面有一个Undo菜单
        edit = Menu(menu)
        edit.add_command(label='Undo')
        edit.add_command(label='Show  Image', command=self.show_img)
        edit.add_command(label='Show  Text', command=self.show_txt)
        menu.add_cascade(label='编辑', menu=edit)
        click_button1 = Button(self, text='稳定', command=lambda: self.write_to_excel(0))
        click_button1.place(x=840, y=200, width=100)
        click_button2 = Button(self, text='功角失稳', command=lambda: self.write_to_excel(1))
        click_button2.place(x=840, y=400, width=100)
        click_button3 = Button(self, text='电压失稳', command=lambda: self.write_to_excel(2))
        click_button3.place(x=840, y=600, width=100)
        # click_button4 = Button(self, text='电压失稳', command=lambda: self.read_excel(2))

    @staticmethod
    def client_exit():
        exit()

    def show_img(self, image_dir, image_x, image_y):
        # load = Image.open(image_dir).resize((800, 600))
        load = Image.open(image_dir)
        width, height = load.size
        load = load.resize((width//4, height//4))
        box = (0, 100, 600, 980)  # 裁剪图片空白部分
        load = load.crop(box)
        render = ImageTk.PhotoImage(load)

        img = Label(self, image=render)
        img.image = render
        img.place(x=image_x, y=image_y)

    def show_txt(self, idx):
        """
        从文件名读取 PSASP 自动判稳结果
        """
        figure_name = self.fig_name[idx]
        mode_id = figure_name[-5]
        if mode_id == '1':
            info = '自动判稳模式：稳定'
        if mode_id == '2':
            info = '自动判稳模式：功角失稳'
        if mode_id == '3':
            info = '自动判稳模式：电压失稳'
        if mode_id == '4':
            info = '自动判稳模式：频率失稳'
        text = Label(self, text=info, font=('微软雅黑', 20, 'bold'))
        text.place(x=200, y=20)

    @staticmethod
    def select_path():
        """
        路径选择函数，用以选择图片文件和标注文件
        """
        path = askdirectory()
        # path.set(path)
        return path

    def sample_path(self):
        """
        初始化样本路径并加载图片
        """
        self.fig_name = self.read_file()
        self.idx = 0
        self.load_fig(self.idx)

    def label_path(self):
        """
        返回标签文件label.exe路径
        """
        path = self.select_path()
        excel_dir = path + '/label' + '.xlsx'
        return excel_dir

    def get_excel_info(self):
        """
        读取excel表格中已经标注的样本位置
        """
        excel_dir = self.excel_dir
        file = xlrd.open_workbook(excel_dir)
        table = file.sheets()[0]  # 得到sheet页
        nrows = table.nrows  # 总行数
        print(nrows)
        return excel_dir, nrows

    def write_to_excel(self, label):
        """
        opeatation: 操作指令，default: 'write_label'，给当前图片写入标签
        可选：'backward':读上一张图片
        将标签自动写入excel文件，加载下一张图片
        """
        excel_dir, nrows = self.get_excel_info()
        wb = openpyxl.load_workbook(excel_dir)
        ws = wb['Sheet1']
        current_nrows = nrows + 1
        ws.cell(row=current_nrows, column=1).value = label
        wb.save(excel_dir)
        wb.close()
        self.load_fig(self.idx)

    @staticmethod
    def search_all_files_return_by_time_reversed(path, reverse=False):
        """
        按照时间排列顺序读取图片
        """
        return sorted(glob.glob(os.path.join(path, '*')),
                      key=lambda x: time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(x))),
                      reverse=reverse)

    def read_file(self):
        """
        读取图片文件名
        """
        path = self.select_path()
        fig_name = self.search_all_files_return_by_time_reversed(path)
        return fig_name

    def load_fig_dir(self, idx):
        """
        获取图片路径
        """
        fig_name = self.fig_name
        image_dir = fig_name[idx]
        return image_dir

    def load_fig(self, idx):
        """
        加载图片
        """
        # print(idx)
        image_dir = self.load_fig_dir(idx)
        self.show_txt(idx)
        print(image_dir)
        self.show_img(image_dir, self.image_x, self.image_y)
        self.idx += 1


root = Tk()
root.geometry("1000x1000")
app = Window(root)
root.mainloop()

