'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
LastEditTime: 2024-10-09 10:23:59
LastEditors: naivsheng naivsheng@outlook.com
Description: get当前文件夹内全部文件名以归档: 获取文件夹名并创建对应的sheet页, 将文件夹中的pdf文件汇总到表格中
FilePath: \crawer\file_Velten.py
'''
import os
import openpyxl
path = os.getcwd()
folder_list = []
with open('files.txt','r',encoding='utf-8') as f:
    FLs = f.readlines()
for file in os.listdir(path):
    if '.' not in file:
        folder_list.append(file)
wb = openpyxl.load_workbook('Velten.xlsx')
sheets = wb.worksheets
for folder in folder_list:
    pdf_list = {x.replace('\n',''):[] for x in FLs}
    if folder not in sheets:
        ws=wb.create_sheet(folder,0)
        headers = ['Filiale','Files']
        ws.append(headers)
    else:
        ws = wb[folder]
    target_path = f'{path}\\{folder}'
    output = os.listdir(target_path)
    for file in output: # get pdf Datein
        if file.endswith('.pdf') or file.endswith('.PDF'):
            filiale = file.split('-')[0]
            if filiale == 'Stuttgart':
                filiale = 'Stuttgart-2'
            elif filiale not in pdf_list:
                pdf_list[filiale] = []
                with open('files.txt','a+',encoding='utf-8')as f:
                    f.write(f'{filiale}\n')
            pdf_list[filiale].append(file) 
    for fl,info in pdf_list.items(): # writedown
        row = [fl] + info
        ws.append(row)
wb.save('Velten.xlsx')