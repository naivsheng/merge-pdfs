import fitz
import tkinter as tk
from tkinter import simpledialog

def split_pdf(input_pdf_path):
    # 打开PDF文件
    doc = fitz.open(input_pdf_path)
    with open('Readme.txt', 'r',encoding='utf-8') as f:
        filenames = f.readlines()
    for page_num in range(doc.page_count):
        '''
        page = doc.load_page(page_num)  # 加载当前页
        # 提取页面文本，用于命名文件
        page_text = page.get_text()
        page_name = extract_page_name(page_text)
        # 如果页面名为空，则使用默认命名
        if not page_name:
            page_name = f"Page_{page_num + 1}"
        
        # 确保文件名合法（去除非法字符）
        page_name = "".join(char for char in page_name if char.isalnum() or char in (" ", "_")).strip()
        '''
        
        # 创建一个新的PDF文档用于保存该页
        new_pdf = fitz.open()
        new_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
        page_name = input_pdf_path.replace('.pdf','') + "_" + filenames[page_num].split(' = ')[1] + "_" + filenames[page_num].split(' = ')[0]
        page_name = page_name.replace('\n','')
        # 保存为单独的PDF文件
        output_pdf_path = f"{page_name}.pdf"
        new_pdf.save(output_pdf_path)
        new_pdf.close()
    
    doc.close()
def get_filename_prefix():
    # 使用tkinter创建一个弹窗输入框
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    filename_prefix = simpledialog.askstring(title="Input", prompt="Enter the filename prefix end with .pdf:")
    root.destroy()  # 销毁主窗口
    return filename_prefix
'''
def extract_page_name(page_text):
    # 页面名提取逻辑
    # 假设页面名在文本的第一行
    return page_text.splitlines()[0] if page_text else "Unnamed_Page"
'''
filename_prefix = get_filename_prefix()
if '.pdf' in filename_prefix:
    split_pdf(filename_prefix)
else:
    filename_prefix = filename_prefix + '.pdf'
    split_pdf(filename_prefix)