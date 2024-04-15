'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
Date: 2024-04-15
LastEditTime: 2024-04-15 09:22:47
LastEditors: naivsheng naivsheng@outlook.com
Description: get当前文件夹内全部文件名以归档
FilePath: \crawer\file_liste.py
'''
import os
import pandas as pd

path = os.getcwd()
pdf_list = [file for file in os.listdir(path) if file.endswith('.pdf') or file.endswith('.PDF')]
# pdf_list = [os.path.join(path,filename) for filename in pdf_list]

if len(pdf_list) < 1:
    print("Do not have any PDF-Documnet to merge")
else:
    pdf_list.insert(0,path)
    df = pd.DataFrame(pdf_list)
    df.to_excel('original Preis liste.xlsx',index=False)
