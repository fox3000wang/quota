#!/usr/bin/python3

import os
import decimal

from openpyxl import Workbook
from openpyxl import load_workbook

# 遍历文件
def findAllFile(base):
    for root, ds, fs in os.walk(base):
        for f in fs:
            if f.endswith('.xlsx'):
                fullname = os.path.join(root, f)
                yield fullname

# 检查报告如果存在就先删除
def checkReport():
    if os.path.exists("report.xlsx"):
        os.remove("report.xlsx")

# 解析xlsx
def parseXlsx(path):
    # 读取
    workbook = load_workbook(filename=path)
    sheet = workbook["Sheet1"]
    print(sheet['B1'].value)

# 生成报告
def generateReport():
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 42
    wb.save("report.xlsx")

# 解析综合文档
def parseDocx():
    workbook = load_workbook(filename='综合.xlsx')
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
    #print(data)
    
    def takeAvg(d):
        return d['avg']
    
    def takeIndex(d):
        return d['index']    

    data.sort(key=takeAvg, reverse=True)
    for index in range(len(data)):
        data[index]['order'] = index + 1
    data.sort(key=takeIndex)

    return data


# 写入分析数据
def writeDocx(data):
    workbook = load_workbook(filename='综合.xlsx')
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
        
        #sheet['W'+str(d['index']+2)] = round((d['avg'] / first['avg']) * 100)
        if d['order'] > 15:
            sheet['W'+str(d['index']+2)] = 40
        else:
            sheet['W'+str(d['index']+2)] = round((d['avg'] / first['avg']) * 100)
        
    workbook.save("report.xlsx")


# 主函数
def main():
    checkReport()
    
    '''
    base = '.'
    for i in findAllFile(base):
        parseXlsx(i)
    generateReport()    
    '''
    summary = parseDocx()
    print(summary)
    # sort_data = handleDocx(summary)
    # print(sort_data)

    writeDocx(summary)

if __name__ == '__main__':
    main()
