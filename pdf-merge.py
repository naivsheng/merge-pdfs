'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
Date: 2022-09-08 11:41:17
LastEditTime: 2024-01-10 14:31:16
LastEditors: naivsheng naivsheng@outlook.com
Description: 合并pdf,提供选项是否需要在单页后加入空白页，默认添加;根据PyPDF2更新代码
FilePath: \crawer\pdf-merge.py
'''
import os
from PyPDF2 import PdfMerger
from PyPDF2 import PdfReader

path = os.getcwd()
pdf_list = [file for file in os.listdir(path) if file.endswith('.pdf') or file.endswith('.PDF')]
# pdf_list = [os.path.join(path,filename) for filename in pdf_list]

file_merger = PdfMerger()
blank = 'blank.pdf'
pdf_list.remove(blank)
check = input('mit blank oder ohne blank(m/o):')
if len(pdf_list)<1:
    print("Do not have any PDF-Documnet to merge")
else:
    if check =="o":
        for pdf in pdf_list:
            pdf_file = open(pdf,'rb')
            reader = PdfReader(pdf_file)
            file_merger.append(pdf)
    else:
        for pdf in pdf_list:
            pdf_file = open(pdf,'rb')
            reader = PdfReader(pdf_file)
            page = len(reader.pages)
            file_merger.append(pdf)
            if page % 2 == 1:
                file_merger.append(blank)

merge_name = 'merger to ' + pdf_list[-1]
file_merger.write(merge_name)