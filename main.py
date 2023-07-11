#!/usr/bin/python3

import os
import decimal

from openpyxl import Workbook
from openpyxl import load_workbook

# 全局就指定一个xls
workbook = load_workbook(filename='综合.xlsx')

# 遍历文件
def findAllFile(base):
    for root, ds, fs in os.walk(base):
        for f in fs:
            if f.endswith('.xlsx'):
                fullname = os.path.join(root, f)
                yield fullname

def take_avg(d):
    return d['avg']
    
def take_index(d):
    return d['index']

def take_sum(d):
    return d['sum']


# 初始化
def init():
    # 检查报告如果存在就先删除    
    if os.path.exists("report.xlsx"):
        os.remove("report.xlsx")


# 01.解析BUG提交
def parse_bug_commit():
    print(workbook)
    sheet = workbook["bug提交"]
    data = []
    
    i = 2
    while(sheet['D'+str(i)].value != None):
        col_range = sheet['D'+str(i)+':T'+str(i)]    
        sum = 0
        month = 0
        for row in col_range:
            for cell in row:
                if(isinstance(cell.value, int)):
                    sum += cell.value
                    month += 1
        data.append({'name':sheet['D'+str(i)].value, 'sum':sum, 'month':month, 'avg': round(sum/month, 4), 'index': i-2 })
        i += 1
    
    data.sort(key=take_avg, reverse=True)
    for index in range(len(data)):
        data[index]['order'] = index + 1
    data.sort(key=take_index)

    return data

# 01.写入BUG提交
def write_bug_commit(data):
    sheet = workbook["bug提交"]

    # 找出第一名
    first = None
    for d in data:
        if d['order'] == 1:
            first = d
            break

    for index in range(len(data)):
        d = data[index]
        sheet['U'+str(d['index']+2)] = d['avg']
        sheet['V'+str(d['index']+2)] = d['order']
        sheet['W'+str(d['index']+2)] = round((d['avg'] / first['avg']) * 100)


# 02.解析文档评审
def parse_doc_review():
    sheet = workbook["文档评审"]
    data = []
    
    i = 2
    while(sheet['E'+str(i)].value != None):
        col_range = sheet['G'+str(i)+':R'+str(i)]
        sum = 0
        h1 = 0
        month = 0
        for row in col_range:
            for cell in row:
                if(isinstance(cell.value, int)):
                    if(month <= 5):
                        h1 += cell.value
                    sum += cell.value
                    month += 1
        data.append({'name':sheet['E'+str(i)].value, 'sum':sum, 'h1':h1, 'index': i-2 })
        i += 1
    
    data.sort(key=take_sum, reverse=True)
    for index in range(len(data)):
        data[index]['order'] = index + 1
    data.sort(key=take_index)

    return data


def write_doc_review(data):
    sheet = workbook["文档评审"]

    # 找出第一名
    first = None
    for d in data:
        if d['order'] == 1:
            first = d
            break

    for index in range(len(data)):
        d = data[index]
        sheet['S'+str(d['index']+2)] = d['sum']
        sheet['T'+str(d['index']+2)] = d['h1']
        
        #sheet['U'+str(d['index']+2)] = round((d['avg'] / first['avg']) * 100)
        #result = d['h1'] - sheet['T'+str(d['index']+2)].value
        
        sheet['V'+str(d['index']+2)] = round((d['h1'] / first['h1']) * 100)


# 主函数
def main():
    
    init()

    bug_commit = parse_bug_commit()
    write_bug_commit(bug_commit)

    doc_review = parse_doc_review()
    write_doc_review(doc_review)

    workbook.save("report.xlsx")


if __name__ == '__main__':
    main()
