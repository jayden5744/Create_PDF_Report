# -*- coding: utf-8 -*-

import ttk
import excel2pdf
import tkFileDialog
import tkMessageBox
from Tkinter import *


class Pdf:
    def __init__(self, master):
        self.message = tkMessageBox
        self.F = Frame(master)
        self.F.pack()
        self.path = ''
        self.save_path = ''
        # 파일경로
        self.path_b = Button(self.F, text='파일경로'.decode('utf-8'), command=lambda: self.path_button_press())
        self.path_b.grid(row=0, column=0)
        self.path_label = Label(self.F, text=self.path, anchor='center')
        self.path_label.grid(row=0, column=1, columnspan=3, sticky='ew')

        # 해당 파일종류 선택
        self.comb_box = ttk.Combobox(self.F)
        self.comb_box['values'] = ("SA8", "SA9", "SA9_1", "SA10", "SA10_1", "SA11",  "SA12", "SA13", "SA14", "SA15")
        self.comb_box.pack()
        self.comb_box.set("시험종류 선택")
        self.comb_box.grid(row=1, column=1, sticky='w')

        # title
        self.title_label = Label(self.F, text='시험 제목'.decode('utf-8'))
        self.title_entry = Entry(self.F)
        self.title_label.grid(row=2, column=0)
        self.title_entry.grid(row=2, column=1, columnspan=2, sticky='ew')

        # description
        self.description = Label(self.F, text='시험 설명'.decode('utf-8'))
        self.des_Text = Text(self.F)
        self.description.grid(row=3, column=0)
        self.des_Text.grid(row=3, column=1, columnspan=3)

        # 저장경로 설정
        self.save_b = Button(self.F, text='저장경로 설정'.decode('utf-8'), command=lambda: self.save_button_press(), width=10)
        self.save_b.grid(row=4, column=0, sticky='es')
        self.save_label = Label(self.F, text=self.save_path, anchor='w')
        self.save_label.grid(row=4, column=1, sticky='w')

        # 변환 시작
        self.convert_b = Button(self.F, text='변환 실행'.decode('utf-8'), command=lambda: self.convert_pdf(), width=2)
        self.convert_b.grid(row=4, column=3, sticky='ew')

    def path_button_press(self):
        value = tkFileDialog.askopenfilename()
        self.path = value
        self.path_label = Label(self.F, text=self.path, anchor='center')
        self.path_label.grid(row=0, column=1, columnspan=3, sticky='ew')

    def save_button_press(self):
        value = tkFileDialog.askdirectory()
        self.save_path = value
        self.save_label = Label(self.F, text=self.save_path, anchor='w')
        self.save_label.grid(row=4, column=1, sticky='w')

    def info(self):
        self.message.showinfo("변환완료", 'Excel파일이 PDF로 변환되었습니다.')

    def convert_pdf(self):
        file_type = self.comb_box.get()
        name = self.path.split('/')[-1]
        name = name.split('.')[0]
        path = self.path.split('/')[0:-1]
        path = str('/'.join(path))
        title = self.title_entry.get()
        description = self.des_Text.get(1.0, 20.30)
        save_path = str(self.save_path)

        if file_type == 'SA8':
            excel2pdf.convert_sa8(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA9':
            excel2pdf.convert_sa9(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA9_1':
            excel2pdf.convert_sa9_1(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA10':
            excel2pdf.convert_sa10(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA10_1':
            excel2pdf.convert_sa10_1(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA11':
            excel2pdf.convert_sa11(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA12':
            excel2pdf.convert_sa12(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA13':
            excel2pdf.convert_sa13(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA14':
            excel2pdf.convert_sa14(name, title, description, path, save_path)
            self.info()
        elif file_type == 'SA15':
            excel2pdf.convert_sa15(name, title, description, path, save_path)
            self.info()


def init():
    tk = Tk()
    tk.title("PDF 변환기".decode('utf-8'))
    tk.minsize(640, 400)
    return tk


if __name__ == '__main__':
    root = init()
    Pdf(root)
    root.mainloop()