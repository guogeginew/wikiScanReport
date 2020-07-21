#!/usr/bin/env python
# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils import copy
import os
from wikivar import *
 
# 功能点wiki
url_api = url_api
 
 
#文件存放路径
file_path = localpath
excel_name = FacilityTransferTaskFileName
 
#判断文件目录的存在性
def dir_exists(file_path):
    is_exists = os.path.exists(file_path)
    if not is_exists:
        os.makedirs(file_path)
        return file_path
    else:
        return file_path
 
#判断excel文件的存在性
def file_excel_exists(file_name):
    #调用判断文件目录存在性的方法
    file_path_name = dir_exists(file_path)
    file_excel_name = os.path.join(file_path_name,file_name)
    is_exsits = os.path.exists(file_excel_name)
    if not is_exsits:
        #创建excel文件
        work_book = xlwt.Workbook(encoding='ascii')
        # 对一个单元格重复操作会引发错误，以cell_overwrite_ok方式新增则不会出现错误
        work_book.add_sheet(u'移交申请（进行中）',cell_overwrite_ok=True)
        work_book.add_sheet(u'移交任务',cell_overwrite_ok=True)
        work_book.add_sheet(u'移交任务（已完成）',cell_overwrite_ok=True)
        work_book.save(file_excel_name)
        return file_excel_name
    else:
        return file_excel_name
        
        
#爬取页面信息
def get_html_content():
    conn = requests.session()
    conn.auth = (wikiusername,wikipassword)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
    }
    res_api = conn.get(url=url_api,headers = headers)
    
    #print(res_api.text)
 
    #生成BeautifulSoup对象，后面查找网页内容时使用
    res_soup_api = BeautifulSoup(str(res_api.text), 'lxml')
    #1.查找符合table，class=confluenceTable的标签
    #2.由于查找到的符合步骤1的标签共10个，而使用时只需要用到第2个，所以取数组下标为1的标签内容
    #3.再次初始化为BeautifulSoup对象，方便后面的使用
    print ("table: ", len(res_soup_api.find_all('table',class_='aui metadata-summary-macro null')))
    export_content(res_soup_api,'aui metadata-summary-macro null',u"移交申请（进行中）")
    export_content(res_soup_api,'aui aui-table-interactive tasks-report',u"移交任务")    
    export_content(res_soup_api,'aui aui-table-interactive tasks-report',u"移交任务（已完成）")    

#导出每一个sheet页
def export_content(res_soup_api,class_name,sheet_name):
    table_class = res_soup_api.find_all('table',class_=class_name)
    if(len(table_class)==0):
       print("no content to be exported")
       return
    table_api_value = BeautifulSoup(str(table_class[0]),'lxml')
    if sheet_name == u"移交任务（已完成）":
        table_api_value = BeautifulSoup(str(table_class[1]),'lxml')
    #生成标题
    export_head = table_api_value.thead.contents
    print("export_head len: ", len(export_head))
    for i in range(len(export_head)):
        #将获取到的tbody标签解析为lxml，并按数组下标逐个获取各个tr标签
        export_head_all_value = BeautifulSoup(str((table_api_value.thead.contents)[i]),'lxml')
        #通过contents属性将获取得的tr直接子节点，得到的是一个数组列表
        export_head_tr = export_head_all_value.tr.contents
        print ("thead.contents len: ", len(export_head_tr))
        for n in range(len(export_head_tr)):
            msg_value = export_head_tr[n].text
            #调用方法向excel表内写入数据
            print("msg_value: ", msg_value)
            excel_head_write(sheet_name,i,n,msg_value)
    print('文件头写入完成...')
 
    #生成内容
    #获取table下的第1个tbody标签下的标签内容，contents只能获取到第一个
    # 通过观察发现tbody标签下的内容正是爬取数据所需要的，得到的是一个数组
    export_content_tbody = table_api_value.tbody.contents
    print("export_content_tbody len: ", len(export_content_tbody))
    for i in range(len(export_content_tbody)):
        #将获取到的tbody标签解析为lxml，并按数组下标逐个获取各个tr标签
        export_content_tbody_all_value = BeautifulSoup(str((table_api_value.tbody.contents)[i]),'lxml')
        #通过contents属性将获取得的tr直接子节点，得到的是一个数组列表
        export_content_tbody_tr = export_content_tbody_all_value.tr.contents
        #print ("tr.contents len: ", len(export_content_tbody_tr))
        for n in range(len(export_content_tbody_tr)):
            msg_value = export_content_tbody_tr[n].text
            #调用方法向excel表内写入数据
            #print("msg_value: ", msg_value)
            excel_data_write(sheet_name,i,n,msg_value)
    print('文件内容导出完成...')
    print("*****************************************************")
    print("*****************************************************")
     

#向excel中写入head
def excel_head_write(sheet_name,i, j, excel_data):
    file_excel_name = file_excel_exists(excel_name)
    workbook = xlrd.open_workbook(file_excel_name)
    # 复制文件并保留格式
    workbook = copy.copy(workbook)
    # 索引到第一个Sheet页
    worksheet = workbook.get_sheet(sheet_name)
    worksheet.write(i, j, label=excel_data)
    workbook.save(file_excel_name)
 
# 向excel中写入数据
def excel_data_write(sheet_name,i, j, excel_data):
    file_excel_name = file_excel_exists(excel_name)
    workbook = xlrd.open_workbook(file_excel_name)
    # 复制文件并保留格式
    workbook = copy.copy(workbook)
    # 索引到第一个Sheet页
    worksheet = workbook.get_sheet(sheet_name)
    worksheet.write(i+1, j, label=excel_data)
    workbook.save(file_excel_name)
 
    #name获取标签名称
    # print(tb1.th.name)
    #text属性获取文本值
    # print(tb1.tr.text)
    # print(tb1.th.div.text)
    #attrs返回所有属性
    # print(tb1.th.attrs)
    # print(tb1.th['class'])
    # print(tb1.th.text)
    # print(tb[1])
    # print(tb[1].text)
 
if __name__ == '__main__':
    #爬取页面信息
    get_html_content()