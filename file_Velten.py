'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
LastEditTime: 2024-10-08 17:27:19
LastEditors: naivsheng naivsheng@outlook.com
Description: get当前文件夹内全部文件名以归档
FilePath: \crawer\file_Verten.py
'''
import os
import pandas as pd

path = os.getcwd()
folder_list = []
with open('files.txt','r',encoding='utf-8') as f:
    FLs = f.readlines()
pdf_list = {x.replace('\n','') for x in FLs}
for file in os.listdir(path):
    if '.' not in file:
        folder_list.append(file)
for folder in folder_list:
    target_file = f'{folder}.xlsx'
    target_path = f'{path}\\{folder}'
    output = os.listdir(target_path)
    for file in output:
        if file.endswith('.pdf') or file.endswith('.PDF'):
            filiale = file.split('-')[0]
            if filiale == 'Stuttgart':
                filiale = 'Stuttgart-2'
            pdf_list[filiale] = file 
    df = pd.DataFrame(pdf_list)
    df.to_excel(target_file,index=False)
