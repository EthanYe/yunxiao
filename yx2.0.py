# 云霄项目小波动统计

import os
import numpy as np
import pandas as pd
import xlwt as xw
import xlrd as xr
import csv



def SteadyFlow(dict0, jz):
    # 读取极值结果文件，由于调压室为4，数据读取位置固定
    jz=3
    b = pd.read_csv('恒定流结果文件.xls', sep='\t',skipfooter=st+1, encoding='gbk',engine='python')

    b = b.iloc[:,1:] 
    b = b.values  # convert to ndarray
    b[:, 0] = b[:, 0]*100
    dict0[dir2] = [b, jz]


def SteadyFlow2Excek(dict0,
                     path=None,
                     filename='小波动恒定流统计结果.xls',
                     sheetname='Sheet1'):

    wb = xw.Workbook()
    ws = wb.add_sheet(sheetname)
    jz=3
    style = xw.XFStyle()
    style.num_format_str = '0.00'  # 保留两位小数
    line = 0  # 写入Excel行号
    index = 0  # 工况数列表索引

    for i in range(xNumber):
        index += 1
        n = i+1
        preCondition='X'
        condition = preCondition+str(n)     # Write condition name
        data = dict0[condition][0]          # 从字典获取数据
        line1 = line+jz-1                   # 本次写入最后一行
        ws.write_merge(line, line1, 0, 0, condition)    # 写入第一列工况名
        for k in range(jz):
            row = line+k
            strUnit = str(4+k)+'#'
            ws.write(row, 1, strUnit)
            for j in range(4):
                ws.write(row, j+2, data[k, j], style)
        line = line1+1

    wb.save(path + '/' + filename)


def SmallTransient(dict0):
    b = pd.read_csv('小波动计算整理.xls', sep='\t'
                    , encoding='gbk', engine='python')
    jz=int((b.shape[0]-6)/2)

    b = b.iloc[:jz+5, 1:]
    b = b.values  # convert to ndarray
    dict0[dir2]=[b,jz]


def SmallTransient2Excel(dict0,
                         path=None,
                         filename='小波动转速特征值结果.xls',
                         sheetname='Sheet1'):
    wb = xw.Workbook()
    ws = wb.add_sheet(sheetname)
    wsTank = wb.add_sheet("调压室水位特征值")
    
    style = xw.XFStyle()
    style.num_format_str = '0.00'  # 保留两位小数
    line = 0  # 写入ExcelxNumber
    index = 0  # 工况数列表索引
    lineTank=0
    

    #first page
    for i in range(xNumber):
        index += 1
        n = i+1
        preCondition='X'
        condition = preCondition+str(n)     # Write condition name
        data = dict0[condition][0]          # 从字典获取数据
        jz = dict0[condition][1]            # 从字典获取机组数
        line1 = line+jz-1                   # 本次写入最后一行
        ws.write_merge(line, line1, 0, 0, condition)    # 写入第一列工况名
       
        for k in range(jz):
            row = line+k
            strUnit = str(4+k)+'#'
            ws.write(row, 1, strUnit)
            for j in range(11):
                ws.write(row, j+2, data[k, j], style)
        
        wsTank.write_merge(lineTank, lineTank+1, 0, 0, condition)    # 写入第一列工况名
        for k in range(2):
            row = lineTank+k
            if k==0:
                strTank = '上游'
            else:
                strTank='下游'
            wsTank.write(row, 1, strTank)
            for j in range(10):
                if j in [5,6,7]:
                    continue
                wsTank.write(row, j+2, data[k+5, j], style)
        line = line1+1
        lineTank=lineTank+2
    #second page
    wb.save(path + '/' + filename)


if __name__ == '__main__':
    # 自定义参数
    itemName = '云霄'       # 项目名称
    st = 4                  # 调压室个数
    jz=3
    xNumber= 9              #小波动工况数   
    jzName=[4,10,14]
    ratedM=306.1
    maxRatedM=ratedM*1.1

    dict_steady={}
    dict_smalltransition={}
    dict_gl
    dict_waterdisturbance={}
    
    # prepath = r'E:\DeskFile\项目\云霄\计算\计算\方案八\水力干扰'  # 计算文件夹路径
    prepath = r'E:\DeskFile\项目\云霄\计算\计算\方案八\小波动'  # 计算文件夹路径
    dir_list = prepath.split('\\')
    lenth = len(dir_list)+1  # 根目录长度
    # tpFolder = os.listdir(prepath)  # 水轮机、水泵工况文件夹名称列表
    for root1, dirs1, files1 in os.walk(prepath):   # 遍历水轮机、水泵工况文件夹
        dir_list = root1.split('\\')  # 目录分段
        if len(files1)>4:
            dir2 = dir_list[lenth].split('（')[0]
            os.chdir(root1)
            SteadyFlow(dict_steady, jz)
            SmallTransient(dict_smalltransition)
            # WaterDisturbance(dict_waterdisturbance)
    SteadyFlow2Excek(dict_steady,  prepath, '小波动恒定流统计结果.xls')
    SmallTransient2Excel(dict_smalltransition,prepath,'小波动转速特征值.xls')

