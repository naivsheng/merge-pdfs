<!--
 * @Author: naivsheng naivsheng@outlook.com
 * @Date: 2024-01-10 14:21:04
 * @LastEditors: naivsheng naivsheng@outlook.com
 * @LastEditTime: 2024-10-09 10:27:52
 * @FilePath: \merge-pdfs\README.md
-->
# merge-pdfs pdf合并
merge pdfs and add blank page to single files

从当前文件目录读取pdf或PDF文件，并将它们按文件名进行合并，为保证双面打印的顺序，在单数页后加入空白页

20240306更新：

若blank.pdf不存在于目标目录，则不添加空页；

否则提供选项是否需要在单页后加入空白页

20240110更新：

提供选项是否需要在单页后加入空白页，默认添加

# file_Velten get当前文件夹内全部文件夹及文件以归档:
## 2024-10-09
获取文件夹名并创建对应的sheet页, 将文件夹中的pdf文件汇总到表格中

用于记录各货柜的文件夹中是否有各店(files.txt)的点货单

# pdf-split 拆分pdf
将pdf的各页以Readme.txt文件的页标进行拆分并命名