'''
Author: naivsheng naivsheng@outlook.com
Date: 2024-10-09 10:52:50
LastEditors: naivsheng naivsheng@outlook.com
LastEditTime: 2024-10-09 11:30:04
FilePath: \merge-pdfs\Velten-split.py
Beschreibung: 拆分当前文件夹下的pdf并标页码, 移动源文件到source文件夹
'''
import fitz
import os
import shutil
def split_pdf(input_pdf_path):
    # 打开PDF文件
    doc = fitz.open(input_pdf_path)
    for page_num in range(doc.page_count):
        # 创建一个新的PDF文档用于保存该页
        new_pdf = fitz.open()
        new_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
        page_name = input_pdf_path.replace('.pdf','') + f"_{page_num + 1}"
        # 保存为单独的PDF文件
        output_pdf_path = f"{page_name}.pdf"
        new_pdf.save(output_pdf_path)
        new_pdf.close()
    doc.close()
path = os.getcwd()
srcpath = f'{path}\\source\\'
if not os.path.exists(srcpath):
    os.makedirs(srcpath)
for file in os.listdir(path):
    if '.PDF' in file.upper():
        split_pdf(file)
        shutil.move(file,srcpath + file)    
