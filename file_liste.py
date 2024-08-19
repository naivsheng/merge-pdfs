'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
Date: 2024-04-15
LastEditTime: 2024-04-16 10:37:33
LastEditors: naivsheng naivsheng@outlook.com
Description: get当前文件夹内全部文件名、文件时间以归档
FilePath: \crawer\file_liste.py
'''
import os
import pandas as pd
import time

path = os.getcwd()
pdf_list = {'filename':[],'ctime':[],'mtime':[]}
# pdf_list = {file:[time.strftime("%Y-%m-%d",time.localtime(os.stat(file).st_mtime)),time.strftime("%Y-%m-%d",time.localtime(os.stat(file).st_ctime))] for file in os.listdir(path) if file.endswith('.pdf') or file.endswith('.PDF')}
# pdf_list = [os.path.join(path,filename) for filename in pdf_list]
for file in os.listdir(path):
    if file.endswith('.pdf') or file.endswith('.PDF'):
        pdf_list['filename'].append(file)
        pdf_list['mtime'].append(time.strftime("%Y-%m-%d",time.localtime(os.stat(file).st_mtime)))
        pdf_list['ctime'].append(time.strftime("%Y-%m-%d",time.localtime(os.stat(file).st_ctime)))                        

if len(pdf_list) < 1:
    print("Do not have any PDF-Documnet to merge")
else:
    df = pd.DataFrame(pdf_list)
    df.to_excel('original Preis liste time.xlsx',index=False)
