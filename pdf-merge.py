'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
Date: 2022-09-08 11:41:17
LastEditTime: 2022-09-08 11:54:13
LastEditors: Yingyu Wang
Description: Version: 合并pdf, 并在单页pdf后加入空白页
FilePath: \crawer\pdf-merge.py
'''
import os
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileReader

path = os.getcwd()
pdf_list = [file for file in os.listdir(path) if file.endswith('.pdf') or file.endswith('.PDF')]
# pdf_list = [os.path.join(path,filename) for filename in pdf_list]

file_merger = PdfFileMerger()
blank = 'blank.pdf'
pdf_list.remove(blank)

for pdf in pdf_list:
    pdf_file = open(pdf,'rb')
    reader = PdfFileReader(pdf_file)
    page = reader.getNumPages()
    file_merger.append(pdf)
    if page % 2 == 1:
        file_merger.append(blank)

merge_name = 'merger to ' + pdf_list[-1]
file_merger.write(merge_name)