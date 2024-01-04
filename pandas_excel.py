#!/usr/bin/env python
# -*- coding:utf-8-*-
# Time      : 2024/1/4 2:24 PM
# Author    : crhstack
# Contact   : crhuazai@163.com
# File      : pandas_excel.py
# Software  : PyCharm
import os.path
import pandas as pd
import datetime


def create_excel(file,sheet):
    '''创建excel'''
    if os.path.exists(file):
        print(datetime.datetime.now(), "文件已存在")
        return ""
    # 创建一个数据字典
    data = {}

    # 将数据字典转换为DataFrame对象
    df = pd.DataFrame(data)
    # 创建一个ExcelWriter对象
    writer = pd.ExcelWriter(file)

    # 将DataFrame写入Excel文件
    df.to_excel(writer)

    # 保存Excel文件
    writer.save()
    print(datetime.datetime.now(), "创建文件成功")


def write_excel(file, sheet_name, data):
    '''excel追加数据'''
    #读取源数据
    original_data = pd.read_excel(file)
    append_data = pd.DataFrame(data)

    # 将新数据与旧数据合并起来
    all_data = original_data.append(append_data)
    all_data.to_excel(file, sheet_name, index=False)
    print(datetime.datetime.now(), "数据写入成功")


def read_excel(file, sheet_name):
    '''读取excel'''
    data = pd.read_excel(file, sheet_name)
    print(datetime.datetime.now(), data.to_json())
    return data.to_json


if __name__ == '__main__':
    file = 'test.xls'
    sheet = '表格1'
    create_excel(file, sheet)
    data = {'Name': ['111Alice', 'Bob', 'Charlie', 'David'],
            'Age': [25, 30, 35, 40],
            'Salary': [50000, 60000, 70000, 80000]}
    write_excel(file, sheet, data)
    # read_excel(file, sheet)
