'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
LastEditTime: 2024-10-09 12:18:16
LastEditors: naivsheng naivsheng@outlook.com
Description: get当前文件夹内全部文件名以归档: 获取文件夹名并创建对应的sheet页, 将文件夹中的pdf文件汇总到表格中, 调整列宽
FilePath: \crawer\file_Velten.py
'''
def read_file():
    with open('files.txt','r',encoding='utf-8') as f:
        FLs = f.readlines()
    return FLs
def write_file(filiale):
    with open('files.txt','w',encoding='utf-8')as f:
        f.writelines(line + "\n" for line in filiale)
import os
import openpyxl
import openpyxl.worksheet
path = os.getcwd()
folder_list = []
for file in os.listdir(path):
    if os.path.isdir(path):
        folder_list.append(file)
wb = openpyxl.load_workbook('Velten.xlsx')
sheets = wb.worksheets
for folder in folder_list:
    FLs = read_file()
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
            pdf_list[filiale].append(file) 
    for fl,info in pdf_list.items(): # writedown
        row = [fl] + info
        ws.append(row)
    write_file(list(pdf_list.keys()))
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 50
sheets = wb.worksheets
if '目录' not in sheets:
    ws = wb.create_sheet('目录',0)
else: ws = wb['目录']
for i, sheet_name in enumerate(sheets):
    link_address = f"'{sheet_name}'!A1"  # 获取对应的链接地址，此处假设跳转到其他工作表的A1单元格
    cell = ws.cell(row=i+1, column=1)  # 获取对应的单元格
    cell.value = f"跳转到 {sheet_name}"
    cell.hyperlink = openpyxl.worksheet.hyperlink(ref=cell.coordinate, location=link_address)  # 设置跳转链接
    # 在其他工作表中添加跳转到新增工作表的链接
    hyperlink_to_new_sheet = openpyxl.worksheet.hyperlink(ref=f"'{ws.title}'!A{cell.row}", location=f"'{ws.title}'!A{cell.row}")
    # 创建超链接文本
    text = "返回总表"
    # 将超链接和文本添加到工作表中
    worksheet = wb[sheet_name]
    max_column = worksheet.max_column
    return_cell = worksheet.cell(row=1, column=max_column+1)
    return_cell.value = text
    return_cell.hyperlink = hyperlink_to_new_sheet
wb.save('Velten.xlsx')