# -*- encoding: cp949 -*-
# doc -> pdf 가능 https://qkqhxla1.tistory.com/402
import os
import comtypes.client

wdFormatPDF = 17

in_file = os.path.abspath('demo.docx')
out_file = os.path.abspath('basic_pdf')

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()