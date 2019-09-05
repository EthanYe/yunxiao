# 云霄项目水力干扰统计

import os
import numpy as np
import pandas as pd
import xlwt as xw
import xlrd as xr
import csv



def WaterDisturbance(dict0, dict1):

    c = pd.read_csv('极值结果文件.xls', sep='\t',
                    encoding='gbk')
    jz = int((c.shape[0]-6)/2)           # number of unit
    # get initial level 
    d = pd.read_csv('恒定流结果文件.xls', sep='\t',
                    encoding='gbk')
    d=d.iloc[:,1:].values
    iniLevel=np.array([[float(d[jz+4][0])],[float(d[jz+1][0])]])
    #get max & min level
    c = c.iloc[[jz+4,jz+1], 1:5].values
    maxLevel=np.hstack((iniLevel,c.astype(float)))
    #get amptitude
    upper_amptitude = maxLevel[:, 1]-maxLevel[:, 0]
    lower_amptitude = maxLevel[:, 0]-maxLevel[:, 3]
    amptitude=np.vstack((upper_amptitude,lower_amptitude))
    level=np.hstack((maxLevel,amptitude.T))


    maxT = 0.0
    if '频率调节' in dir_list:
        df_small = pd.read_csv('小波动计算整理.xls', sep='\t',
                        encoding='gbk', engine='python')
        jz = int((df_small.shape[0]-6)/2)           # number of unit
        listJZ = df_small.iloc[:jz, 0].values       # name list of unit   
        jz_gr = len(listJZ)
        for i in range(jz_gr):
            listJZ[i]=int(listJZ[i][4:] )           # get int in unit name
        
        for i,k in zip(listJZ,range(jz_gr)):
            a = pd.read_csv('机组J'+str(i)+'非恒定流详细信息.xls', sep='\t',
                            encoding='gbk')
            a = a.values
            iniM = a[0][6]
            maxM = a[:, 6].max()
            overM = maxM-ratedM
            relativeOverM = overM/ratedM*100
            
            for t, j in a[:, [0, 6]]:
                if j > maxRatedM:
                    maxT = t+0.4
            adjustTime = df_small.iloc[k, 7]
            vibraNum = df_small.iloc[k, 9]
            if i == listJZ[0]:
                b = np.array([iniM, maxM, overM, relativeOverM, maxT,adjustTime,vibraNum])
            else:
                b = np.vstack((b, [iniM, maxM, overM, relativeOverM, maxT,adjustTime,vibraNum]))
        b=b.astype(float)
        dict0[dir2]=[b,listJZ,level]
    elif '功率调节' in dir_list:

        for i in dict0[dir2][1]:
            a = pd.read_csv('机组J'+str(i)+'非恒定流详细信息.xls', sep='\t',
                            encoding='gbk')
            a = a.values

            iniM = a[0][6]
            maxM = a[:, 6].max()
            overM = maxM-ratedM
            relativeOverM = overM/ratedM*100
            for t, j in a[:, [0, 6]]:
                if j > maxRatedM:
                    maxT = t+0.4
            if i == dict0[dir2][1][0]:
                b = np.array([iniM, maxM, overM, relativeOverM, maxT])
            else:
                b = np.vstack((b, [iniM, maxM, overM, relativeOverM, maxT]))
        b = b.astype(float)
        dict1[dir2] = [b, dict0[dir2][1],level]


def WaterDisturbance2Excel(dict0,dict1,
                           path=None, filename='水力干扰统计结果.xls',
                           sheetname='Sheet1'):

    wb = xw.Workbook()
    ws1 = wb.add_sheet('机组——频率调节')
    ws2 = wb.add_sheet('调压室——频率调节')
    ws3 = wb.add_sheet('机组——功率调节')
    ws4 = wb.add_sheet('调压室——功率调节')

    style = xw.XFStyle()
    style.num_format_str = '0.00'  # 保留两位小数
    line = 0  # 写入ExcelxNumber
    index = 0  # 工况数列表索引
    lineTank = 0

    # 机组——频率调节
    for i in range(grNumber):
        index += 1
        n = i+1
        preCondition = 'GR'
        condition = preCondition+str(n)     # Write condition name
        
        # 从字典获取数据
        data1 = dict0[condition][0]          
        data2=dict0[condition][2]
        data3 = dict1[condition][0]          # 从字典获取数据
        data4 = dict1[condition][2]

        listUnit = dict0[condition][1]        
        jz = len(listUnit)      # 从字典获取机组数
        line1 = line+jz-1                   # 本次写入最后一行

        ws1.write_merge(line, line1, 0, 0, condition)    # 写入第一列工况名
        ws2.write_merge(lineTank, lineTank+1, 0, 0, condition)    # 写入第一列工况名
        ws3.write_merge(line, line1, 0, 0, condition)    # 写入第一列工况名
        ws4.write_merge(lineTank, lineTank+1, 0, 0, condition)    # 写入第一列工况名
        for k in range(jz):
            row = line+k
            if listUnit[k]==4:
                strUnit = str(4)+'#'
            elif listUnit[k]==10:
                strUnit = str(5)+'#'
            elif listUnit[k] == 14:
                strUnit = str(6)+'#'
            ws1.write(row, 1, strUnit)
            ws3.write(row, 1, strUnit)
            if jz==1:
                for j in range(7):
                    ws1.write(row, j+2, data1[j], style)
                for j in range(5):
                    ws3.write(row, j+2, data3[j], style)
            else:
                for j in range(7):
                    ws1.write(row, j+2, data1[k, j], style)
                for j in range(5):
                    ws3.write(row, j+2, data3[k, j], style)

        for k in range(2):
            row = lineTank+k
            if k == 0:
                strTank = '上游'
            else:
                strTank = '下游'
            ws2.write(row, 1, strTank)
            ws4.write(row, 1, strTank)
            for j in range(7):
                ws2.write(row, j+2, data2[k, j], style)
                ws4.write(row, j+2, data4[k, j], style)
        line = line1+1
        lineTank = lineTank+2
    #second page
    wb.save(path + '/' + filename)


if __name__ == '__main__':
    # 自定义参数
    itemName = '云霄'       # 项目名称
    st = 4                  # 调压室个数
    jz = 3
    xNumber = 9             # 小波动工况数
    grNumber=18             # 水力干扰工况数
    jzName = [4, 10, 14]
    ratedM = 306.1
    maxRatedM = ratedM*1.1

    dict_steady = {}
    dict_smalltransition = {}
    dict_waterdisturbance1 = {}
    dict_waterdisturbance2 = {}

    prepath = r'E:\DeskFile\项目\云霄\计算\计算\方案八\水力干扰'
    prepath1 = r'E:\DeskFile\项目\云霄\计算\计算\方案八\水力干扰\频率调节'  # 计算文件夹路径
    prepath2 = r'E:\DeskFile\项目\云霄\计算\计算\方案八\水力干扰\功率调节'  # 计算文件夹路径
    # prepath = r'E:\DeskFile\项目\云霄\计算\计算\方案八\小波动'  # 计算文件夹路径
    dir_list = prepath1.split('\\')
    lenth = len(dir_list)  # 根目录长度
    # tpFolder = os.listdir(prepath)  # 水轮机、水泵工况文件夹名称列表
    for root1, dirs1, files1 in os.walk(prepath1):   # 遍历水轮机、水泵工况文件夹
        dir_list = root1.split('\\')  # 目录分段
        if len(files1) > 4:
            dir2 = dir_list[lenth].split('（')[0]
            os.chdir(root1)
            # SteadyFlow(dict_steady, jz)
            # SmallTransient(dict_smalltransition)
            WaterDisturbance(dict_waterdisturbance1, dict_waterdisturbance2)

    for root1, dirs1, files1 in os.walk(prepath2):   # 遍历水轮机、水泵工况文件夹
        dir_list = root1.split('\\')  # 目录分段
        if len(files1) > 4:
            dir2 = dir_list[lenth].split('（')[0]
            os.chdir(root1)
            # SteadyFlow(dict_steady, jz)
            # SmallTransient(dict_smalltransition)
            WaterDisturbance(dict_waterdisturbance1, dict_waterdisturbance2)
    # SteadyFlow2Excek(dict_steady,  prepath, '小波动恒定流统计结果.xls')
    # SmallTransient2Excel(dict_smalltransition,prepath,'小波动转速特征值.xls')
    WaterDisturbance2Excel(dict_waterdisturbance1, dict_waterdisturbance2,prepath,'水力干扰统计结果.xls')
