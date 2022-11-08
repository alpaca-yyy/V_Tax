#!/usr/bin/env python
# coding: utf-8

# In[4]:


import win32com.client as win32
import os

def exchange(dir):
#:param dir: product_count,product_trend,product_before15 文件夹
#:return:
    #返回当前脚本的绝对路径 利用.split方法path分割成目录和文件名二元组 返回目录
    path = os.path.abspath(__file__).split('src')[0]
    path = os.path.join(path,'file','source_file',dir)
    files = os.listdir(path)
    for file_name in files:
        if file_name.rsplit('.',1)[-1]=='xls':
            fname = os.path.join(path,file_name)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)
            #在原来的位置创建出：原名+'.xlsx'文件
            wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            os.remove(fname)

