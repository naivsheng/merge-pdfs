'''
# -*- coding: UTF-8 -*-
Author: Yingyu Wang
LastEditTime: 2024-10-09 14:26:24
LastEditors: naivsheng naivsheng@outlook.com
Description: get当前文件夹内全部文件名以归档: 获取文件夹名并创建对应的sheet页, 将文件夹中的pdf文件汇总到表格中, 调整列宽
FilePath: \crawer\file_Velten.py
'''
def read_file():
    with open('files.txt','r',encoding='utf-8') as f:
        FLs = f.readlines()
    return FLs
def write_file(filiale):
    elements_with_char = [s for s in filiale if '.pdf' in s]
    for elem in elements_with_char:
        filiale.pop(filiale.index(elem))
    with open('files.txt','w',encoding='utf-8')as f:
        f.writelines(line + "\n" for line in filiale)
import os
import openpyxl
import openpyxl.worksheet
path = os.getcwd()
folder_list = []
for file in os.listdir(path):
    if os.path.isdir(file):
        folder_list.append(file)
wb = openpyxl.load_workbook('Velten.xlsx')
sheets_name = wb.sheetnames
for folder in folder_list:
    FLs = read_file()
    pdf_list = {x.replace('\n',''):[] for x in FLs}
    if folder not in sheets_name:
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
            if 'Turmstr' in filiale:
                filiale = 'BerlinTurmstr'
            elif filiale not in pdf_list:
                pdf_list[filiale] = []
            pdf_list[filiale].append(file) 
    ws.delete_cols(1,ws.max_column)
    for fl,info in pdf_list.items(): # writedown
        row = [fl] + info
        ws.append(row)
    write_file(list(pdf_list.keys()))
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 50
sheets = wb.sheetnames
if '目录' in sheets:
    del wb['目录']
ws = wb.create_sheet('目录', 0)
for i, sheet_name in enumerate(sheets):
    if sheet_name == '目录':
        continue
    # 在目录工作表中添加跳转到其他工作表的链接
    cell = ws.cell(row=i + 1, column=1)
    cell.value = f"Link to {sheet_name}"  # 设置单元格显示文本
    cell.hyperlink = f"#{sheet_name}!A1"  # 设置跳转链接
    cell.style = "Hyperlink"  # 设置超链接样式
    # 在目标工作表中添加返回目录的超链接
    worksheet = wb[sheet_name]
    return_cell = worksheet.cell(row=1, column=3)
    # 创建或更新“返回目录”的超链接
    if return_cell.hyperlink is None or return_cell.hyperlink.target != "目录!A1":
        return_cell.value = "返回目录"
        return_cell.hyperlink = f"#目录!A1"  # 设置跳转到新目录页的链接
        return_cell.style = "Hyperlink"  # 设置超链接样式
# 保存Excel文件
wb.save('Velten.xlsx')