import re
import os
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.filedialog import askdirectory, askopenfilenames
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import fitz
from concurrent.futures import ThreadPoolExecutor

_global_dict = {}
temporaryDict = {}  # 用于存放处理好的多个文件以备输出


class MainGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF工具V1.1")
        self.screenWidth = self.winfo_screenwidth()  # 屏幕宽度
        self.screenHeight = self.winfo_screenheight()
        self.geometry("450x350+%d+%d" % (self.screenWidth / 2, self.screenHeight / 3))
        self.resizable(False, False)  # 禁止改变窗口大小x,y
        self.configure(background='#f3f3f3')  # 设置背景颜色

        self.tickRotate = tk.IntVar()
        self.tickSplit = tk.IntVar()
        self.tickMerge = tk.IntVar()
        self.tickToPNG = tk.IntVar()
        self.tickToPDF = tk.IntVar()
        self.angle = tk.StringVar(value=0)
        self.dpi = tk.StringVar(value=300)
        self.pageNumber = tk.StringVar(value='请使用英文逗号')
        self.filePaths = ''
        self.fileName = ''
        self.cachePath = None
        self.currentValue = 0
        self.poor = ThreadPoolExecutor(max_workers=4)

        ttk.Style().configure('T.TLabel', foreground='green')
        ttk.Style().configure('F.TLabel', foreground='red')

        self.funcFrame = ttk.LabelFrame(self, text='功能区')
        self.chooseFile = ttk.Button(self, text='选择', width=20, command=self.choose)  # 选择文件按钮
        self.pdfPathShow = ttk.Label(self, text='  ')  # 展示路径
        self.rotateShow = ttk.Entry(self.funcFrame, textvariable=self.angle, state='disabled', width=5)  # 显示角度
        self.needPages = ttk.Entry(self, textvariable=self.pageNumber, width=20)  # 输入须操作页码
        self.cb1 = ttk.Combobox(self.funcFrame, state='disabled', value=('所选页码合并', '所选页码拆为单页'))
        self.dpiShow = ttk.Entry(self.funcFrame, textvariable=self.dpi, state='disabled', width=5)
        self.feedbackShow = ttk.Label(self, text='  ', style='T.TLabel')
        self.executeButton = ttk.Button(self, text='执行操作', width=10, command=self.execute)  # 执行操作按钮
        self.exportFile = ttk.Button(self, text='导出文件', width=10, command=self.export)  # 导出文件按钮
        self.bar = ttk.Progressbar(self, length=180, mode="determinate")
        self.book()

    def book(self):
        self.chooseFile.place(x=30, y=17, width=50)
        ttk.Label(self, text='文件路径：').place(x=90, y=20)
        self.pdfPathShow.place(x=158, y=20)
        ttk.Label(self, text='需操作页码：').place(x=28, y=55)
        self.needPages.place(x=108, y=54)
        self.needPages.bind('<Button-1>', self.tip1)
        self.needPages.bind('<Key>', self.tip1)
        self.needPages.bind('<Leave>', self.tip2)
        ttk.Label(self, text='默认全部,自定义如：3-8,10,12').place(x=265, y=55)
        self.funcFrame.place(x=30, y=90)
        ttk.Checkbutton(self.funcFrame, text='旋转', variable=self.tickRotate, command=self.needRotate) \
            .grid(column=0, row=0, sticky='w')
        ttk.Button(self.funcFrame, text='逆90°', command=self.cClockwise, width=-1).grid(column=1, row=0, padx=10,
                                                                                        ipadx=2)
        self.rotateShow.grid(column=2, row=0)
        ttk.Button(self.funcFrame, text='顺90°', command=self.clockwise, width=-1).grid(column=3, row=0, padx=10,
                                                                                       ipadx=2)
        ttk.Checkbutton(self.funcFrame, text='PDF页面拆分', variable=self.tickSplit, command=self.needSplit).grid(column=0,
                                                                                                              row=1)
        ttk.Checkbutton(self.funcFrame, text='多PDF合并', variable=self.tickMerge, command=self.needMerge).grid(column=0,
                                                                                                             row=2,
                                                                                                             sticky='w')
        ttk.Checkbutton(self.funcFrame, text='PDF转图片', variable=self.tickToPNG, command=self.needToPNG).grid(column=0,
                                                                                                             row=3,
                                                                                                             sticky='w')
        ttk.Label(self.funcFrame, text='DPI值(72~1200)').grid(column=1, row=3, columnspan=2)
        ttk.Checkbutton(self.funcFrame, text='图片转PDF', variable=self.tickToPDF, command=self.needToPDF).grid(column=0,
                                                                                                             row=4,
                                                                                                             sticky='w')
        self.dpiShow.grid(column=3, row=3, sticky='w')
        self.cb1.grid(column=1, row=1, columnspan=3, padx=10, pady=3)
        self.cb1.current(0)

        self.feedbackShow.place(x=187, y=278)
        self.executeButton.place(x=125, y=305)
        self.exportFile.place(x=225, y=305)
        self.bar.place(x=125, y=255)

    def tip1(self, event):
        if self.needPages.get() == '请使用英文逗号':
            self.pageNumber.set('')

    def tip2(self, event):
        if self.needPages.get() == '':
            self.pageNumber.set('请使用英文逗号')

    def init(self):  # 输出后初始化内存数据
        self.tickRotate.set(0)
        self.tickSplit.set(0)
        self.cb1.configure(state='disabled')
        self.dpiShow.configure(state='disabled')
        self.tickMerge.set(0)
        self.tickToPNG.set(0)
        self.tickToPDF.set(0)
        self.cb1.set('所选页码合并')
        self.currentValue = 0
        self.progress(0)
        try:
            _global_dict.clear()
            temporaryDict.clear()
            self.fileName = re.search(r'[^/]*(?=\.pdf|\.jpg|\.png)', self.filePaths[0]).group(0)  # 重置文件名
        except IndexError:
            pass
        except TypeError:
            pass

    def choose(self):
        self.filePaths = askopenfilenames(filetypes=[('PDF files', '.pdf'), ('Pic files', ('.png', '.jpg'))])
        if self.filePaths != '':
            self.cachePath = self.filePaths
            self.fileName = re.search(r'[^/]*(?=\.pdf|\.jpg|\.png)', self.filePaths[0]).group(0)
            if len(''.join(self.filePaths)) < 36:
                self.pdfPathShow['text'] = self.filePaths
            else:
                self.pdfPathShow['text'] = ''.join(self.filePaths)[:32] + '...'
        elif self.filePaths == '' and self.fileName != '':
            self.filePaths = self.cachePath

    def rotate(self):  # 旋转PDF
        if len(self.filePaths) == 1:
            pdf_reader = PdfReader(self.filePaths[0])
            set_value('pdf_read', pdf_reader)
            pdf_writer = PdfWriter()
            if self.needPagesRefine()[0] == -1:
                self.feedbackShow.configure(style='F.TLabel')
                self.feedbackShow['text'] = '执行失败！'
            elif self.needPagesRefine()[0] == -99:
                for i in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[i].rotate(int(self.angle.get()))
                    pdf_writer.add_page(page)
            else:
                for i in range(len(pdf_reader.pages)):
                    if i not in self.needPagesRefine():
                        page = pdf_reader.pages[i]
                        pdf_writer.add_page(page)
                    else:
                        page = pdf_reader.pages[i].rotate(int(self.angle.get()))
                        pdf_writer.add_page(page)
            self.feedbackShow.configure(style='T.TLabel')
            self.feedbackShow['text'] = '执行成功！'
        else:
            messagebox.showinfo('', '选择文件数量不为1')
            self.feedbackShow.configure(style='F.TLabel')
            self.feedbackShow['text'] = '执行失败！'
            return
        set_value('pdf_done', pdf_writer)

    def split(self):
        if len(self.filePaths) == 1:
            if get_value('pdf_done'):
                pdf_writer = PdfWriter()  # 虽然有经过其他步骤处理的文档，但仍需要一个空白文档放置本步骤处理的内容
                if self.cb1.get() == '所选页码合并' and self.needPagesRefine()[0] not in [-99, -1]:
                    for i in self.needPagesRefine():
                        page = get_value('pdf_done').pages[i]
                        pdf_writer.add_page(page)
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                elif self.cb1.get() == '所选页码合并' and self.needPagesRefine()[0] == -99:
                    for i in range(len(get_value('pdf_read').pages)):
                        page = get_value('pdf_done').pages[i]
                        pdf_writer.add_page(page)
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                elif self.cb1.get() == '所选页码拆为单页' and self.needPagesRefine()[0] not in [-99, -1]:
                    for i in self.needPagesRefine():
                        page = get_value('pdf_done').pages[i]
                        pdf_writer.add_page(page)
                        temporaryDict[self.fileName + '_p' + str(i + 1)] = pdf_writer  # 文件名_页码：文件
                        pdf_writer = PdfWriter()
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                elif self.cb1.get() == '所选页码拆为单页' and self.needPagesRefine()[0] == -99:
                    for i in range(len(get_value('pdf_read').pages)):
                        page = get_value('pdf_done').pages[i]
                        pdf_writer.add_page(page)
                        temporaryDict[self.fileName + '_p' + str(i + 1)] = pdf_writer  # 文件名_页码：文件
                        pdf_writer = PdfWriter()
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
            elif not get_value('pdf_done'):  # 此为第一步时
                pdf_reader = PdfReader(self.filePaths[0])
                pdf_writer = PdfWriter()
                if self.cb1.get() == '所选页码合并' and self.needPagesRefine()[0] not in [-99, -1]:
                    for i in self.needPagesRefine():
                        page = pdf_reader.pages[i]
                        pdf_writer.add_page(page)
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                elif self.cb1.get() == '所选页码合并' and self.needPagesRefine()[0] == -99:
                    for i in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[i]
                        pdf_writer.add_page(page)
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                elif self.cb1.get() == '所选页码拆为单页' and self.needPagesRefine()[0] not in [-99, -1]:
                    for i in self.needPagesRefine():
                        page = pdf_reader.pages[i]
                        pdf_writer.add_page(page)
                        temporaryDict[self.fileName + '_p' + str(i + 1)] = pdf_writer
                        pdf_writer = PdfWriter()  # 没有清空pdf的方法，只能通过初始化重建一个了
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                elif self.cb1.get() == '所选页码拆为单页' and self.needPagesRefine()[0] == -99:
                    for i in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[i]
                        pdf_writer.add_page(page)
                        temporaryDict[self.fileName + '_p' + str(i + 1)] = pdf_writer
                        pdf_writer = PdfWriter()
                    set_value('pdf_done', pdf_writer)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                else:
                    self.feedbackShow.configure(style='F.TLabel')
                    self.feedbackShow['text'] = '执行失败！'
        else:
            messagebox.showinfo('', '选择文件数量不为1')
            self.feedbackShow.configure(style='F.TLabel')
            self.feedbackShow['text'] = '执行失败！'
            return False

    def merge(self):
        pdf_merger = PdfMerger(strict=False)
        pdfList = [f for f in self.filePaths if f.endswith('.pdf')]
        if len(pdfList) != 0:
            for pdf in pdfList:
                pdf_merger.append(PdfReader(pdf, strict=False))  # 合并pdf文件
            set_value('pdf_done', pdf_merger)
            self.feedbackShow.configure(style='T.TLabel')
            self.feedbackShow['text'] = '执行成功！'
        else:
            messagebox.showinfo('文件格式错误', '应为PDF文件')
            self.feedbackShow.configure(style='F.TLabel')
            self.feedbackShow['text'] = '执行失败！'

    def toPNG(self, dpi):
        if len(self.filePaths) == 1:
            try:
                int(dpi)
            except ValueError:
                messagebox.showinfo('', '请输入整数！')
                self.feedbackShow.configure(style='F.TLabel')
                self.feedbackShow['text'] = '执行失败！'
                return
            if int(dpi) < 72:
                messagebox.showinfo('输入数值过小', '过小的DPI会使图片非常模糊')
            elif int(dpi) > 1200:
                messagebox.showinfo('输入数值过大', '普通设备的最高解析度一般为1200DPI\n过高的DPI会增加电脑的处理负担，且生成较慢')
            else:
                pdf_reader = fitz.Document(self.filePaths[0])
                temporaryDict.clear()
                if self.needPagesRefine()[0] == -1:
                    self.feedbackShow.configure(style='F.TLabel')
                    self.feedbackShow['text'] = '执行失败！'
                elif self.needPagesRefine()[0] == -99:
                    for pgn in range(pdf_reader.page_count):
                        page = pdf_reader[pgn]
                        pix = page.get_pixmap(dpi=int(dpi))
                        name = '%s-%d.png' % (self.fileName, pgn + 1)
                        temporaryDict[name] = pix
                        set_value('pdfToPng', '')
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                    pdf_reader.close()
                else:
                    for pgn in self.needPagesRefine():
                        page = pdf_reader[pgn]
                        pix = page.get_pixmap(dpi=int(dpi))
                        name = '%s-%d.png' % (self.fileName, pgn + 1)
                        temporaryDict[name] = pix
                        set_value('pdfToPng', '')
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '执行成功！'
                    pdf_reader.close()
                if 'pdf_done' in _global_dict:
                    _global_dict.pop('pdf_done')


        else:
            messagebox.showinfo('', '选择文件数量不为1')
            self.feedbackShow.configure(style='F.TLabel')
            self.feedbackShow['text'] = '执行失败！'
            return

    def toPDF(self):
        if len(self.filePaths) != 0:
            pdf_reader = fitz.Document()
            for f in self.filePaths:
                img = fitz.Document(f)  # 逐一打开指定图片
                rect = img[0].rect  # 获取图片的尺寸
                pdfBytes = img.convert_to_pdf()  # 将图片转成pdf
                img.close()
                imgPDF = fitz.Document("pdf", pdfBytes)
                page = pdf_reader.new_page(width=rect.width, height=rect.height)  # 规定页面大小
                page.show_pdf_page(rect, imgPDF, 0)  # 将其写入空文件
            self.feedbackShow.configure(style='T.TLabel')
            self.feedbackShow['text'] = '执行成功！'
            set_value('pngToPDF', pdf_reader)
            _global_dict.pop('pdf_done')
            pdf_reader.close()
        else:
            self.feedbackShow.configure(style='F.TLabel')
            self.feedbackShow['text'] = '执行失败！'

    def progress(self, v):
        self.bar['value'] = v
        self.bar.update()

    def schedule(self, v):
        while self.currentValue <= 100:
            self.currentValue += 10
            self.bar.after(v, self.progress(self.currentValue))
            if self.currentValue == 100:
                self.progress(0)
            if self.feedbackShow['text'] == '执行成功！':
                self.progress(100)
                self.currentValue = 0
                self.executeButton.configure(state='normal')
                self.exportFile.configure(state='normal')
                break
            elif self.feedbackShow['text'] == '执行失败！':
                self.currentValue = 0
                self.executeButton.configure(state='normal')
                self.exportFile.configure(state='normal')
                break
        self.currentValue = 0
        self.progress(0)
        self.executeButton.configure(state='normal')
        self.exportFile.configure(state='normal')


    def needRotate(self):
        if self.tickRotate.get() == 1:
            self.tickMerge.set(0)
            self.tickToPNG.set(0)
            self.tickToPDF.set(0)
            self.dpiShow.configure(state='disabled')
            # self.rotateShow.config(state='normal')
        else:
            pass
            # self.rotateShow.config(state='disabled')暂不提供任意角度旋转

    def needSplit(self):
        if self.tickSplit.get() == 1:
            self.cb1.configure(state='readonly')
            self.dpiShow.configure(state='disabled')
            self.tickMerge.set(0)
            self.tickToPNG.set(0)
            self.tickToPDF.set(0)
        else:
            self.cb1.configure(state='disabled')

    def needMerge(self):
        if self.tickMerge.get() == 1:
            self.tickRotate.set(0)
            self.tickSplit.set(0)
            self.tickToPNG.set(0)
            self.tickToPDF.set(0)
            self.cb1.configure(state='disabled')
            self.dpiShow.configure(state='disabled')
        else:
            pass

    def needToPNG(self):
        if self.tickToPNG.get() == 1:
            self.dpiShow.configure(state='normal')
            self.tickRotate.set(0)
            self.tickSplit.set(0)
            self.tickMerge.set(0)
            self.tickToPDF.set(0)
            self.cb1.configure(state='disabled')
        else:
            self.dpiShow.configure(state='disabled')

    def needToPDF(self):
        if self.tickToPDF.get() == 1:
            self.tickRotate.set(0)
            self.tickSplit.set(0)
            self.tickMerge.set(0)
            self.tickToPNG.set(0)
            self.cb1.configure(state='disabled')
            self.dpiShow.configure(state='disabled')
        else:
            pass

    def cClockwise(self):
        self.angle.set(int(self.rotateShow.get()) - 90)
        if self.rotateShow.get() == '-360':
            self.angle.set(0)

    def clockwise(self):
        self.angle.set(int(self.rotateShow.get()) + 90)
        if self.rotateShow.get() == '360':
            self.angle.set(0)

    def fileTitle(self):
        if self.tickRotate.get() == 1:
            self.fileName += '_已旋转'
        if self.tickSplit.get() == 1:
            self.fileName += '_已拆分'
        if self.tickMerge.get() == 1:
            self.fileName += '_已合并'
        if self.tickToPNG.get() == 1:
            self.fileName += '_转图片'

    def execute(self):
        self.fileTitle()
        self.executeButton.configure(state='disabled')
        self.exportFile.configure(state='disabled')
        self.poor.submit(self.schedule, 100)
        if self.tickRotate.get() == 1:
            self.poor.submit(self.rotate)
        if self.tickSplit.get() == 1:
            self.poor.submit(self.split)
        if self.tickMerge.get() == 1:
            self.poor.submit(self.merge)
        if self.tickToPNG.get() == 1:
            self.poor.submit(self.toPNG, self.dpi.get())
        if self.tickToPDF.get() == 1:
            self.poor.submit(self.toPDF)

    def export(self):
        saveDir = askdirectory()
        pdfPath = os.path.join(saveDir, self.fileName + '.pdf')
        if saveDir != '' and 'pdf_done' in _global_dict:  # 选择了保存目录并且执行过任一操作
            if self.cb1.get() != '所选页码拆为单页':  # 旋转和所选页码合并走这里
                with open(pdfPath, 'wb') as out:
                    get_value('pdf_done').write(out)
                self.feedbackShow.configure(style='T.TLabel')
                self.feedbackShow['text'] = '保存成功！'
            elif self.cb1.get() == '所选页码拆为单页':  # 所选页码为单文件走这里
                if not os.path.exists(saveDir + '/' + self.fileName):  # 如果不存在拆分文件夹
                    os.mkdir(saveDir + '/' + self.fileName)
                    for name, file in temporaryDict.items():
                        with open(saveDir + '/' + self.fileName + '/' + name + '.pdf', 'wb') as out:
                            file.write(out)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '保存成功！'
                else:
                    for name, file in temporaryDict.items():
                        with open(saveDir + '/' + self.fileName + '/' + name + '.pdf', 'wb') as out:
                            file.write(out)
                    self.feedbackShow.configure(style='T.TLabel')
                    self.feedbackShow['text'] = '保存成功！'
        elif saveDir != '' and 'pdfToPng' in _global_dict:
            if not os.path.exists(saveDir + '/' + self.fileName):
                os.mkdir(saveDir + '/' + self.fileName)
                for name, file in temporaryDict.items():
                    file.save('%s/%s' % (saveDir + '/' + self.fileName, name))
                self.feedbackShow.configure(style='T.TLabel')
                self.feedbackShow['text'] = '保存成功！'
        elif saveDir != '' and 'pngToPDF' in _global_dict:
            get_value('pngToPDF').save("%s/%s-merged.pdf" % (saveDir, self.fileName))
            self.feedbackShow.configure(style='T.TLabel')
            self.feedbackShow['text'] = '保存成功！'
        else:
            self.feedbackShow.configure(style='F.TLabel')
            self.feedbackShow['text'] = '保存失败！'
        self.init()

    def needPagesRefine(self):  # 这个函数用于整理用户输入的页码
        pageStr = self.needPages.get()
        page_list = re.split(r'[,，\s]+', pageStr)
        #page_list = pageStr.split(',')  # 转成列表格式方便操作
        pageList = []  # 这里犯了个愚蠢的错误，直接把数据加入page_list导致循环迭代并报错，因此新建一个列表
        if page_list == [''] or page_list == ['请使用英文逗号']:
            return [-99]
        else:
            try:
                for i in page_list:
                    if '-' in i:
                        try:
                            a = re.findall(r'\b\d+', i)[0]
                            b = re.findall(r'\b\d+', i)[1]
                            new = list(range(int(a) - 1, int(b)))
                            pageList.extend(new)
                        except IndexError:
                            messagebox.showinfo('错误提示', '负数？')
                            return [-1]
                    else:
                        c = int(i)
                        pageList.append(c - 1)
                return list(set(pageList))
            except ValueError:
                messagebox.showinfo('错误提示', '可能输入了中文标点或非数字')
                return [-1]


def doc():
    messagebox.showinfo('帮助', '1.执行的操作不会叠加生效，先保存\n'
                              '2.选择文件时右下角可切换文件类型\n'
                              '3.过大的PDF转图片可能会卡住\n'
                              '4.由于PDF的特殊性，转成图片必然有质量亏损\n'
                              '5.其他已知或未知、不可见人的bug\n'
                              '6.作者：易言未济')


class Menu(MainGUI):
    def __init__(self):
        super().__init__()
        self.menu_bar = tk.Menu(self)
        self.setting_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="其他", menu=self.setting_menu)
        self.setting_menu.add_command(label="关于", command=doc)

        self.menu_init()

    def menu_init(self):
        self.config(menu=self.menu_bar)


def _init():  # 初始化
    global _global_dict
    _global_dict = {}


def set_value(key, value):
    # 定义一个全局变量
    _global_dict[key] = value


def get_value(key):
    # 获得一个全局变量，不存在则提示读取对应变量失败
    try:
        return _global_dict[key]
    except KeyError:
        print('读取' + key + '失败\r\n')
