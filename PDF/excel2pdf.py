# -*- coding: utf-8 -*-
import os
import docx
import xlrd
import argparse
import openpyxl
import datetime
import numpy as np
import pandas as pd
import excel2pdf_gui
import comtypes.client
from docx import Document
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from docx.shared import Inches, RGBColor
import sys

reload(sys)
sys.setdefaultencoding('utf-8')


def get_args():
    parser = argparse.ArgumentParser()
    # feature information
    parser.add_argument('--gui', action='store_true', help='GUI를 실행합니다.'.decode('utf-8'))
    parser.add_argument('--sa12', action='store_true', help='SA12파일을 변환합니다.'.decode('utf-8'))
    parser.add_argument('--sa13', action='store_true', help='SA13파일을 변환합니다.'.decode('utf-8'))
    parser.add_argument('--path', type=str, help='엑셀파일이 들어있는 폴더의 경로를 입력하세요.'.decode('utf-8'))
    parser.add_argument('--save_path', type=str, help='PDF를 저장할 폴더의 경로를 입력하세요.'.decode('utf-8'))
    parser.add_argument('--filename', type=str, help='엑셀파일 이름을 입력하세요.'.decode('utf-8'))
    parser.add_argument('--title', type=str, help='시험제목을 입력하세요.'.decode('utf-8'))
    parser.add_argument('--description', default='', type=str, help='시험 설명을 입력하세요.'.decode('utf-8'))

    return parser.parse_args()


# 스타일 선언
def style(document):
    style_1 = document.styles.add_style('Heading_1', docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)
    style_1.base_style = document.styles['Heading 1']
    style_1.font.name = 'HY 견고딕'.decode('utf-8')
    style_1._element.rPr.rFonts.set(qn('w:eastAsia'), 'HY 견고딕'.decode('utf-8'))  # 한글 폰트를 따로 설정해 준다
    style_1.font.size = docx.shared.Pt(15)
    style_1.font.bold = True
    style_1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    style_2 = document.styles.add_style('text', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    style_2.font.name = '돋움체'.decode('utf-8')
    style_2._element.rPr.rFonts.set(qn('w:eastAsia'), '돋움체'.decode('utf-8'))  # 한글 폰트를 따로 설정해 준다
    style_2.font.size = docx.shared.Pt(10)
    style_2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    return style_1, style_2


# xls파일을 xlsx로 변환
def xls2xlsx(name, path=None, **kw):
    if path is None:
        xls_name = name + '.xls'
    else:
        xls_name = path + '/' + name + '.xls'
    book_xls = xlrd.open_workbook(xls_name, formatting_info=True, ragged_rows=True, **kw)
    book_xlsx = openpyxl.workbook.Workbook()

    sheet_names = book_xls.sheet_names()
    for sheet_index in range(len(sheet_names)):
        sheet_xls = book_xls.sheet_by_name(sheet_names[sheet_index])
        if sheet_index == 0:
            sheet_xlsx = book_xlsx.active
            sheet_xlsx.title = sheet_names[sheet_index]
        else:
            sheet_xlsx = book_xlsx.create_sheet(title=sheet_names[sheet_index])
        for c_range in sheet_xls.merged_cells:
            rlo, rhi, clo, chi = c_range
            sheet_xlsx.merge_cells(start_row=rlo + 1, end_row=rhi,
                                   start_column=clo + 1, end_column=chi, )

        def _get_xlrd_cell_value(cell):
            value = cell.value
            if cell.ctype == xlrd.XL_CELL_DATE:
                datetime_tup = xlrd.xldate_as_tuple(value, 0)
                if datetime_tup[0:3] == (0, 0, 0):  # time format without date
                    value = datetime.time(*datetime_tup[3:])
                else:
                    value = datetime.datetime(*datetime_tup)
            return value

        for row in range(sheet_xls.nrows):
            sheet_xlsx.append((
                _get_xlrd_cell_value(cell)
                for cell in sheet_xls.row_slice(row, end_colx=sheet_xls.row_len(row))
            ))
    if path is None:
        book_xlsx.save(name + '.xlsx')
    else:
        book_xlsx.save(path + '/' + name + '.xlsx')


def convert_sa12(sa12_name, sa12_title, sa12_description, sa12_path=None, sa12_save_path=None):
    print '-----------------SA12 연구파일 pdf변환을 시작합니다.---------------------'.decode('utf-8')
    # data_only = True로 해줘야 수식이 아닌 값으로 받아온다.
    if sa12_path is None:  # 파일과 같은 폴더에 있을 때
        try:
            load_wb = load_workbook(sa12_name + ".xlsx")
        except IOError:  # xls파일인 경우
            xls2xlsx(sa12_name, sa12_path)
            load_wb = load_workbook(sa12_name + ".xlsx")
    else:  # 따로 경로를 지정했을 때
        try:
            load_wb = load_workbook(str(sa12_path) + '\\' + sa12_name + ".xlsx")
        except IOError:  # xls파일인 경우
            xls2xlsx(sa12_name, sa12_path)
            load_wb = load_workbook(str(sa12_path) + '\\' + sa12_name + ".xlsx")
    print '-----------------Excel File을 성공적으로 불러왔습니다.---------------------'.decode('utf-8')

    # 시트이름으로 불러오기
    load_ws = load_wb['Index']
    all_value = []
    for row in load_ws.rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        all_value.append(row_value)
    all_value = pd.DataFrame(all_value[1:]).sort_values(by=3)
    p_title = all_value[1].dropna(axis=0).reset_index(drop=True)  # 테이블 제목
    r_title = all_value[2].dropna(axis=0).reset_index(drop=True)  # 결과 검토
    result_summary = str(load_ws['A2'].value)
    load_ws2 = load_wb[result_summary]

    all_values = []
    for row in load_ws2.rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        all_values.append(row_value)
    column = all_values[0]
    data = all_values[1:]
    df = pd.DataFrame(data=data, columns=column)
    df2 = df[['Power Level (%)', 'Iteration', 'PF Target', 'PF Actual 1', 'PF Actual 2', 'PF Actual 3']]
    load_wb.close()

    # 사용하기 위한 변수 선언
    document = Document()
    sa12_title = sa12_title.encode('utf-8')
    # sa12_title = unicode(sa12_title).encode('utf-8')
    sa12_description = sa12_description.encode('utf-8')
    # sa12_description = unicode(sa12_description).encode('utf-8')
    # 제목
    style_1, style_2 = style(document)
    document.add_paragraph(sa12_title.decode('utf-8'), style=style_1)
    last_paragraph = document.paragraphs[-1]
    last_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER  # 중앙정렬
    # 시험 설명
    document.add_paragraph('시험 설명'.decode('utf-8'), style='ListBullet')
    document.add_paragraph(sa12_description.decode('utf-8'), style=style_2)
    # 기능시험 결과
    document.add_paragraph('기능시험 결과'.decode('utf-8'), style='ListBullet')
    num = df2['PF Target'].unique()
    num = np.sort(num[1:])

    # 테이블 작성
    for i in range(len(p_title)):
        mer_title = p_title[i]
        document.add_paragraph(mer_title, style='ListNumber')
        table = document.add_table(rows=1, cols=6, style='Light Shading')
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '출력 (%)'.decode('utf-8')
        hdr_cells[1].text = '반복횟수'.decode('utf-8')
        hdr_cells[2].text = '목표 역률'.decode('utf-8')
        hdr_cells[3].text = '실제 역률 (A)'.decode('utf-8')
        hdr_cells[4].text = '실제 역률 (B)'.decode('utf-8')
        hdr_cells[5].text = '실제 역률 (C)'.decode('utf-8')
        for a, b, c, d, e, f in df2[df2['PF Target'] == num[i]].values.tolist():
            row_cells = table.add_row().cells
            row_cells[0].text = str(a)
            row_cells[1].text = str(b)
            row_cells[2].text = str(c)
            row_cells[3].text = str(d)
            row_cells[4].text = str(e)
            row_cells[5].text = str(f)
        try:
            temp = r_title[i]
            mer_title2 = '* 결과검토: ' + temp
            document.add_paragraph(mer_title2, style=style_2)
        except KeyError:
            pass
    document.add_paragraph('기능시험 결과 요약'.decode('utf-8'), style='List Bullet')
    document.add_paragraph(str(load_ws['C2'].value).decode('utf-8'), style=style_2)

    # docx파일을 생성을 위한 save('파일명')
    document.save(sa12_name.decode('utf-8') + '.docx')
    print '-----------------Docs File을 성공적으로 불러왔습니다.---------------------'.decode('utf-8')

    # 파일 경로 절대경로로
    in_file = os.path.abspath(sa12_name.decode('utf-8') + '.docx')
    if sa12_save_path is None:
        out_file = os.path.abspath(sa12_name)
    else:
        out_file = os.path.abspath(str(sa12_save_path) + '\\' + sa12_name)
    # word형식의 파일을 열기
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    # PDF형식으로 저장
    doc.SaveAs(out_file, FileFormat=17)
    doc.Close()
    word.Quit()
    print '-----------------PDF File을 성공적으로 만들었습니다.---------------------'.decode('utf-8')


def convert_sa13(sa13_name, sa13_title, sa13_description, sa13_path, sa13_save_path):
    print '-----------------SA13 연구파일 pdf변환을 시작합니다.---------------------'.decode('utf-8')
    # data_only = True로 해줘야 수식이 아닌 값으로 받아온다.
    if sa13_path is None:  # 파일과 같은 폴더에 있을 때
        try:
            load_wb = load_workbook(sa13_name + ".xlsx")
        except IOError:  # xls파일인 경우
            xls2xlsx(sa13_name, sa13_path)
            load_wb = load_workbook(sa13_name + ".xlsx")
    else:  # 따로 경로를 지정했을 때
        try:
            load_wb = load_workbook(str(sa13_path) + '/' + sa13_name + ".xlsx")
            print(str(sa13_path) + '/' + sa13_name + ".xlsx")
        except IOError:  # xls파일인 경우
            xls2xlsx(sa13_name, sa13_path)
            load_wb = load_workbook(str(sa13_path) + '/' + sa13_name + ".xlsx")
    print '-----------------Excel File을 성공적으로 불러왔습니다.---------------------'.decode('utf-8')
    # 시트이름으로 불러오기
    load_ws = load_wb['Index']
    all_value = []
    for row in load_ws.rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        all_value.append(row_value)
    all_value = pd.DataFrame(all_value[1:])
    p_title = all_value[1][1:].dropna(axis=0).reset_index(drop=True)  # 결과 검토
    r_title = all_value[2].dropna(axis=0).reset_index(drop=True)  # 결과 검토
    result_summary = str(load_ws['A2'].value)
    load_ws2 = load_wb[result_summary]

    all_values = []
    for row in load_ws2.rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        all_values.append(row_value)
    column = all_values[0]
    data = all_values[1:]
    df = pd.DataFrame(data=data, columns=column)
    # 그림 만들기
    index_down = df[df['Dataset File'].str.contains('down') == True].index
    index_up = df[df['Dataset File'].str.contains('up') == True].index
    for i in range(len(index_down)):
        fig = plt.figure()
        plt.rcParams["figure.figsize"] = (10, 6)
        plt.rcParams['axes.grid'] = True

        img_title = df['Dataset File'][index_down[i]]
        if i == 0:
            vv_1_1000 = df[0:index_down[i]]
        else:
            vv_1_1000 = df[index_up[i - 1] + 1:index_down[i]]
        ax = fig.add_subplot(1, 1, 1)
        ax.plot(vv_1_1000['Average Voltage (pu)'], vv_1_1000['Var Actual 1'] / 4000, linestyle='', marker='o',
                color='blue', label='Power')
        ax.plot(vv_1_1000['Average Voltage (pu)'], vv_1_1000['Var Target 1'] / 4000, color='black', label='VV curve')
        ax.plot(vv_1_1000['Average Voltage (pu)'], vv_1_1000['Var Min Allowed 1'] / 4000, linestyle=':', color='red',
                label='VV pass/fail band')
        ax.plot(vv_1_1000['Average Voltage (pu)'], vv_1_1000['Var Max Allowed 1'] / 4000, linestyle=':', color='red')
        ax.set_title('Volt-Var Function1')
        ax.set_xlabel('Grid Voltage(% nominal)')
        ax.set_ylabel('Reactive Power(% nameplate)')
        ax.set_xticks([0.9, 0.95, 1, 1.05, 1.1])
        ax.set_xticklabels(['90', '95', '100', '105', '110'])
        ax.set_yticks([-1.5, -1, -0.5, 0, 0.5, 1.0, 1.5])
        ax.set_yticklabels(['-150', '-100', '-50', '0', '50', '100', '150'])

        plt.savefig('img/' + img_title + '.png')

    # 사용하기 위한 변수 선언
    document = Document()
    sa13_title = sa13_title.encode('utf-8')
    sa13_description = sa13_description.encode('utf-8')

    # 제목
    style_1, style_2 = style(document)  # 스타일 설정
    document.add_paragraph(sa13_title.decode('utf-8'), style=style_1)
    last_paragraph = document.paragraphs[-1]
    last_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER  # 중앙정렬
    # 시험설명
    document.add_paragraph('시험 설명'.decode('utf-8'), style='ListBullet')
    document.add_paragraph(sa13_description.decode('utf-8'), style=style_2)
    # 판정기준
    document.add_paragraph('판정기준'.decode('utf-8'), style='ListBullet')
    document.add_paragraph(str(load_ws['B2'].value).decode('utf-8'), style=style_2)

    # 기능시험결과
    document.add_paragraph('기능시험 결과'.decode('utf-8'), style='ListBullet')
    for i in range(len(index_down)):
        img_title = df['Dataset File'][index_down[i]]
        mer_title = str(p_title[i])
        document.add_paragraph(mer_title.decode('utf-8'), style='ListNumber')
        document.add_picture('img/' + str(img_title) + '.png', width=Inches(5))  # 그림 불러와서 넣기
        last_paragraph = document.paragraphs[-1]
        last_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER  # 중앙정렬
        caption = '<' + str(mer_title) + '>'  # 캡션 달기
        document.add_paragraph(caption.decode('utf-8'), style=style_2)
        last_paragraph = document.paragraphs[-1]
        last_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER  # 중앙정렬
        try:
            # 결과검토 / 결과검토가 없을경우 발생하는 에러를 위해 try except구문
            temp = r_title[i]
            mer_title2 = '* 결과검토: ' + temp
            document.add_paragraph(mer_title2, style=style_2)
        except KeyError:
            pass
    # 기능시험 결과 요약
    document.add_paragraph('기능시험 결과 요약'.decode('utf-8'), style='ListBullet')
    document.add_paragraph(str(load_ws['C2'].value).decode('utf-8'), style=style_2)

    # docx파일을 생성을 위한 save('파일명')
    document.save(sa13_name.decode('utf-8') + '.docx')
    print '-----------------Docs File을 성공적으로 불러왔습니다.---------------------'.decode('utf-8')

    # 파일 경로 절대경로로
    in_file = os.path.abspath(sa13_name.decode('utf-8') + '.docx')
    if sa13_save_path is None:
        out_file = os.path.abspath(sa13_name)
    else:
        out_file = os.path.abspath(str(sa13_save_path) + '\\' + sa13_name)
    # word형식의 파일을 열기
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    # PDF형식으로 저장
    doc.SaveAs(out_file, FileFormat=17)
    doc.Close()
    word.Quit()
    print '-----------------PDF File을 성공적으로 만들었습니다.---------------------'.decode('utf-8')


if __name__ == '__main__':
    args = get_args()
    # -------------------------------------------------------------------------------------------------------------- #
    # command :
    #       - python excel2pdf.py --파일종류 --filename 파일이름 --title 시험제목 --description 시험설명
    #       - python excel2pdf.py --파일종류 --filename 파일이름 --title 시험제목 --description 시험설명 --path 경로
    #       [파일이 다른 경로에 있을때]
    #
    #       ex) python excel2pdf.py --sa12 --filename SA12 --title 시험제목 --description 시험설명
    #       ex) python excel2pdf.py --sa13 --filename SA13 --title 시험제목 --description 시험설명 --path C:\\python
    #
    # sa12/sa13 : 시험파일의 종류
    # filename : 해당 폴더에 있는 파일 이름
    # title : PDF 제목으로 들어가게될 내용
    # optional
    # description : PDF 내 시험설명으로 들어가게될 내용
    # path : 해당 엑셀파일이 있는 폴더의 경로
    # save_path : PDF의 저장 경로
    # -------------------------------------------------------------------------------------------------------------- #
    while args.sa12 or args.sa13:
        args.title = unicode(args.title.decode('cp949'))
        args.description = unicode(args.description.decode('cp949'))
        if args.sa12:
            try:
                convert_sa12(args.filename, args.title, args.description, args.path, args.save_path)
            except IOError:
                print "파일을 찾지 못했습니다.".decode('utf-8')
                print 'path: ' + args.path
                print 'file name: ' + args.filename

        elif args.sa13:
            try:
                convert_sa13(args.filename, args.title, args.description, args.path, args.save_path)
            except IOError:
                print "파일을 찾지 못했습니다.".decode('utf-8')
                print 'path: ' + str(args.path).decode('utf-8')
                print 'file name: ' + args.filename.decode('utf-8')

        repeat = raw_input('Do you want to continue?(y/n)'.decode('utf-8'))
        if repeat == 'y' or repeat == 'Y':
            filename = raw_input('Enter the following file name.')
            title = raw_input('Enter the following title')
            description = raw_input('Enter the following file description')
            args.filename = filename
            args.title = title
            args.description = description
        else:
            break
    # -------------------------------------------------------------------------------------------------------------- #
    # command :
    #       - python excel2pdf.py --gui
    #
    # -------------------------------------------------------------------------------------------------------------- #
    if args.gui:
        root = excel2pdf_gui.init()
        excel2pdf_gui.Pdf(root)
        root.mainloop()