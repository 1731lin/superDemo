# -*- coding: utf-8 -*-
"""
Created on Thu May 16 16:07:36 2019
实现的功能：利用Python将两个Excel表格对应位置相加，并存储在一个新的Excel文件中。
但是因两个Excel表的第一列和第一行都是一样的，所以不需要相加，保持原样存储至新的Excel表格中：


@author: 17319
"""
import xlrd
import xlwt
import os



def combine_excel(file_1,file_2,i2):
    wb_pri = xlrd.open_workbook(file_1)   #打开原始文件
    wb_tar = xlrd.open_workbook(file_2)   #打开目标文件
    wb_result = xlwt.Workbook()           #新建一个文件，用来保存结果
    sheet_result = wb_result.add_sheet('Sheet1',cell_overwrite_ok=True)

    sheet_pri = wb_pri.sheet_by_index(0)  # 通过index获取每个sheet
    sheet_tar = wb_tar.sheet_by_index(0)  # 通过index获取每个sheet
    ncols = sheet_pri.ncols               # Excel列的数目  原Excel和目标Excel的列表的长度相同
    row_0=sheet_pri.row_values(0)         #获取第一行的值
    col_0=sheet_pri.col_values(0)         #获取第一列的值
    for i,key in enumerate(row_0):       #写入新Excel表的第一行
        sheet_result.write(0,i,key)
    for i,key in enumerate(col_0):        #写入新Excel表的第一列
        sheet_result.write(i,0,key)
    for i1 in range(1,ncols):              #将Excel表格对应位置相加
        l_p = sheet_pri.col_values(i1,start_rowx=1,end_rowx=None)#每列的元素
        l_t = sheet_tar.col_values(i1,start_rowx=1,end_rowx=None)
        l_r=[]
        
        for i in range(0, len(l_p)):#除了第一次都是字符串，后面都是一个float，一个字符串
            l_p_s=l_p[i]
            l_t_s=l_t[i]
            if isinstance(l_p[i],str):#去','去' '
                l_p[i]=l_p[i].replace(',', '')
                l_p_s=l_p[i].strip()
            if isinstance(l_t[i],str):
                l_t[i]=l_t[i].replace(',', '')
                l_t_s=l_t[i].strip()
                
            if l_p_s=='' and l_t_s!='':
                l_r.append(float(l_t[i]))
            elif l_t_s=='' and l_p_s!='':
                l_r.append(float(l_p[i]))
            elif l_t_s!='' and l_p_s!='':
                l_r.append(float(l_p[i])+float(l_t[i]))
            else:
                l_r.append(0)
        print(l_r)
        for j,key in enumerate(l_r):
            sheet_result.write(j+1,i1,key)
    wb_result.save(r"./people/save"+str(i2)+".xls")
 
if __name__=="__mian__": 
    path=r'./people'
    filenames = os.listdir(path)#获取所有xls
    file_1 = path+"/"+filenames[0]
    file_2 =  ''
    for i2, filename in enumerate(filenames):
        if i2>0:
            file_2=path+"/"+filename
            combine_excel(file_1,file_2,i2)
            file_1=path+"/save"+str(i2)+".xls"
            
            print(file_2)
    for i2 in range(1,7):
        try:
            os.remove(r"./people/save"+str(i2)+".xls")     
        except Exception:
            print("异常")

