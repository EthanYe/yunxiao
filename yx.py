# 云霄项目大波动统计

import os
import numpy as np
import pandas as pd
import xlwt as xw
import xlrd as xr
import csv


def Extremum(dict0):
    # 读取压力脉动文件
    b = pd.read_csv('考虑压力脉动与计算误差后的修正结果.xls', sep='\t', encoding='gbk')
    b = b.values  # convert to ndarray
    if(b[-1][0] == '1#'):
        jz = 2
    elif (b[-1][0] == '2#'):
        jz = 3
    b = b[-(jz*6):, 1:]
    b = b.astype('float')
    for i in [1, 3, 5]:
        # 1-蜗壳，3-尾水管，5-转速
        sta = jz*i  # 起始行号
        end = sta+3  # 终止行号
        p = list(b[sta:end, 6])  # 取出各机组的极值
        if(i == 3):                 # 获取最大极值
            max_p = min(p)
        else:
            max_p = max(p)
        pos = p.index(max_p)
        t_p = b[sta+pos, 3]  # 获取对应时间
        list_p = [max_p, t_p, dir2]
        if i == 3:
            item = '尾水位进口最小动水压力'
            if item in dict0:
                if dict0[item][0] > max_p:
                    dict0[item] = list_p
            else:
                dict0[item] = list_p
        else:
            if i == 1:
                item = '蜗壳末端最大动水压力'
            elif i == 5:
                item = '转速最大上升率'
            if item in dict0:
                if dict0[item][0] < max_p:
                    dict0[item] = list_p
            else:
                dict0[item] = list_p

    # 读取极值结果文件，由于调压室为4，数据读取位置固定
    a = np.genfromtxt('极值结果文件.txt',
                      delimiter="\t",
                      skip_header=4,
                      max_rows=st + 1,
                      usecols=range(1, 5))
    if pd.isnull(a[0][0]):              # 第一个数据是否为nan（若是，则3个机组）
        a = np.delete(a, 0, axis=0)     # 此处删除第一行的元素
    # 更新上游调压室最高涌浪
    if '上游调压室最高涌浪' in dict0:
        if dict0['上游调压室最高涌浪'][0] < a[p1][0]:
            dict0['上游调压室最高涌浪'] = [a[p1][0], a[p1][1], dir2]
    else:
        dict0['上游调压室最高涌浪'] = [a[p1][0], a[p1][1], dir2]

    # 更新上游调压室最低涌浪
    if '上游调压室最低涌浪' in dict0:
        if dict0['上游调压室最低涌浪'][0] > a[p1][2]:
            dict0['上游调压室最低涌浪'] = [a[p1][2], a[p1][3], dir2]
    else:
        dict0['上游调压室最低涌浪'] = [a[p1][2], a[p1][3], dir2]

    # 更新上闸最高涌浪
    if '上闸最高涌浪' in dict0:
        if dict0['上闸最高涌浪'][0] < a[p2][0]:
            dict0['上闸最高涌浪'] = [a[p2][0], a[p2][1], dir2]
    else:
        dict0['上闸最高涌浪'] = [a[p2][0], a[p2][1], dir2]

    # 更新上闸最低涌浪
    if '上闸最低涌浪' in dict0:
        if dict0['上闸最低涌浪'][0] > a[p2][2]:
            dict0['上闸最低涌浪'] = [a[p2][2], a[p2][3], dir2]
    else:
        dict0['上闸最低涌浪'] = [a[p2][2], a[p2][3], dir2]

    # 更新尾水调压室最高涌浪
    if '尾水调压室最高涌浪' in dict0:
        if dict0['尾水调压室最高涌浪'][0] < a[p3][0]:
            dict0['尾水调压室最高涌浪'] = [a[p3][0], a[p3][1], dir2]
    else:
        dict0['尾水调压室最高涌浪'] = [a[p3][0], a[p3][1], dir2]

    # 更新尾水调压室最低涌浪
    if '尾水调压室最低涌浪' in dict0:
        if dict0['尾水调压室最低涌浪'][0] > a[p3][2]:
            dict0['尾水调压室最低涌浪'] = [a[p3][2], a[p3][3], dir2]
    else:
        dict0['尾水调压室最低涌浪'] = [a[p3][2], a[p3][3], dir2]

    # 更新下闸最高涌浪
    if '下闸最高涌浪' in dict0:
        if dict0['下闸最高涌浪'][0] < a[p4][0]:
            dict0['下闸最高涌浪'] = [a[p4][0], a[p4][1], dir2]
    else:
        dict0['下闸最高涌浪'] = [a[p4][0], a[p4][1], dir2]

    # 更新下闸最低涌浪
    if '下闸最低涌浪' in dict0:
        if dict0['下闸最低涌浪'][0] > a[p4][2]:
            dict0['下闸最低涌浪'] = [a[p4][2], a[p4][3], dir2]
    else:
        dict0['下闸最低涌浪'] = [a[p4][2], a[p4][3], dir2]


def UnitResult(dict0):
    b = pd.read_csv('考虑压力脉动与计算误差后的修正结果.xls', sep='\t', encoding='gbk')
    b = b.values  # convert to ndarray
    if(b[-1][0] == '1#'):
        jz = 2
    elif (b[-1][0] == '2#'):
        jz = 3
    b = b[-(jz*6):, 1:]
    b = b.astype('float')
    i = 2*jz
    j = 4*jz
    delList = list(range(i, i+jz))
    delList = delList+list(range(j, j+jz))
    b = np.delete(b, delList, 0)
    dict0[dir2] = [b, jz]


def UnitResult2excel(dict0,
                     path=None,
                     filename='统计结果.xls',
                     sheetname='Sheet1'):
    wb = xw.Workbook()
    ws = wb.add_sheet(sheetname)

    style = xw.XFStyle()
    style.num_format_str = '0.00'  # 保留两位小数
    listN = [8, 16, 6, 6]  # 工况数列表
    listCol = [1, 2, 3, 7]  # 保留两位小数列表
    line = 0  # 写入Excel行号
    index = 0  # 工况数列表索引
    for m in listN:
        index += 1
        for i in range(m):
            n = i+1
            if index == 1:
                preCondition = 'SJT'
            elif index == 2:
                preCondition = 'JHT'
            elif index == 3:
                preCondition = 'SJP'
            elif index == 4:
                preCondition = 'JHP'
            condition = preCondition+str(n)     # Write condition name
            data = dict0[condition][0]          # 从字典获取数据
            jz = dict0[condition][1]            # 从字典获取几组数
            line1 = line+jz*4-1                 # 本次写入最后一行
            ws.write_merge(line, line1, 0, 0, condition)  # 写入第一列工况名
            line2 = line                               # 第二列索引
            line3 = line2+jz-1
            ws.write_merge(line2, line3, 1, 1, '球阀上游最大动水压力(m)')
            for k in range(jz):
                strUnit = str(4+k)+'#'
                ws.write(line2+k, 2, strUnit)
            line2 = line3+1                               # 第二列索引
            line3 = line2+jz-1
            ws.write_merge(line2, line3, 1, 1, '蜗壳最大动水压力(m)')
            for k in range(jz):
                strUnit = str(4+k)+'#'
                ws.write(line2+k, 2, strUnit)
            line2 = line3+1                               # 第二列索引
            line3 = line2+jz-1
            ws.write_merge(line2, line3, 1, 1, '尾水管最小动水压力 (m)')
            for k in range(jz):
                strUnit = str(4+k)+'#'
                ws.write(line2+k, 2, strUnit)
            line2 = line3+1                               # 第二列索引
            line3 = line2+jz-1
            ws.write_merge(line2, line3, 1, 1, '机组最大转速 (rpm)')
            for k in range(jz):
                strUnit = str(4+k)+'#'
                ws.write(line2+k, 2, strUnit)
            for j in range(7):
                for k in range(jz*4):
                    row = line+k
                    if j+1 in listCol:
                        ws.write(row, j+3, data[k, j], style)
                    else:
                        ws.write(row, j+3, data[k, j])
            line = line1+1

    wb.save(path + '/' + filename)


def Extremum2excel(dict0,
                   dict00,
                   path=None,
                   filename='统计结果.xls',
                   sheetname='Sheet1'):
    wb = xw.Workbook()
    ws = wb.add_sheet(sheetname)
    i = 0
    for key in dict0.items():
        ws.write(0, i, key[0])
        ws.write(1, i, key[1][0])
        ws.write(2, i, key[1][1])
        ws.write(3, i, key[1][2])
        i += 1
    i = 0
    for key in dict00.items():
        ws.write(5, i, key[0])
        ws.write(6, i, key[1][0])
        ws.write(7, i, key[1][1])
        ws.write(8, i, key[1][2])
        i += 1
    wb.save(path + '/' + filename)


def SteadyFlow(dict0, jz):
    # 读取极值结果文件，由于调压室为4，数据读取位置固定
    a = np.genfromtxt('恒定流结果文件.txt',
                      delimiter="\t",
                      skip_header=1,
                      max_rows=jz,
                      usecols=range(1, 5))
    a[:, 0] = a[:, 0]*100
    dict0[dir2] = [a, jz]


def SteadyFlow2Excek(dict0,
                     path=None,
                     filename='恒定流统计结果.xls',
                     sheetname='Sheet1'):

    wb = xw.Workbook()
    ws = wb.add_sheet(sheetname)

    style = xw.XFStyle()
    style.num_format_str = '0.00'  # 保留两位小数
    line = 0  # 写入Excel行号
    index = 0  # 工况数列表索引
    for m in listN:
        index += 1
        for i in range(m):
            n = i+1
            if index == 1:
                preCondition = 'SJT'
            elif index == 2:
                preCondition = 'JHT'
            elif index == 3:
                preCondition = 'SJP'
            elif index == 4:
                preCondition = 'JHP'
            condition = preCondition+str(n)     # Write condition name
            data = dict0[condition][0]          # 从字典获取数据
            jz = dict0[condition][1]            # 从字典获取几组数
            line1 = line+2                   # 本次写入最后一行
            ws.write_merge(line, line1, 0, 0, condition)    # 写入第一列工况名
            for k in range(3):
                if k > jz-1:
                    continue
                row = line+k
                strUnit = str(4+k)+'#'
                ws.write(row, 1, strUnit)
                for j in range(4):
                    ws.write(row, j+2, data[k, j], style)
            line = line1+1

    wb.save(path + '/' + filename)


def GateTank(dict_1, dict_2, dict_3, dict_4, jz):
    # 读取极值结果文件，由于调压室为4，数据读取位置固定
    a = np.genfromtxt('恒定流结果文件.txt',
                      delimiter="\t",
                      skip_header=jz+2,
                      max_rows=st,
                      usecols=range(1, 2))

    b = np.genfromtxt('极值结果文件.txt',
                      delimiter="\t",
                      skip_header=jz+2,
                      max_rows=st,
                      usecols=range(1, 9))

    dict_1[dir2] = [a[p2]] + list(b[p2])    # 上闸
    dict_2[dir2] = [a[p1]] + list(b[p1])    # 上调
    dict_3[dir2] = [a[p3]] + list(b[p3])    # 下调
    dict_4[dir2] = [a[p4]] + list(b[p4])    # 下闸


def GateTank2Excel(dict_1, dict_2, dict_3, dict_4):
    df_1 = pd.DataFrame(dict_1)
    df_1 = df_1.T
    writer1 = pd.ExcelWriter(prepath+'\上闸.xlsx')					# 写入Excel文件
    # ‘page_1’是写入excel的sheet名
    df_1.to_excel(writer1, 'page_1', float_format='%.2f')
    writer1.save()
    writer1.close()

    df_2 = pd.DataFrame(dict_2)
    df_2 = df_2.T
    writer2 = pd.ExcelWriter(prepath+'\上调.xlsx')					# 写入Excel文件
    # ‘page_2’是写入excel的sheet名
    df_2.to_excel(writer2, 'page_2', float_format='%.2f')
    writer2.save()
    writer2.close()

    df_3 = pd.DataFrame(dict_3)
    df_3 = df_3.T
    writer3 = pd.ExcelWriter(prepath+'\下调.xlsx')					# 写入Excel文件
    # ‘page_3’是写入excel的sheet名
    df_3.to_excel(writer3, 'page_3', float_format='%.2f')
    writer3.save()
    writer3.close()

    df_4 = pd.DataFrame(dict_4)
    df_4 = df_4.T
    writer4 = pd.ExcelWriter(prepath+'\下闸.xlsx')					# 写入Excel文件
    # ‘page_4’是写入excel的sheet名
    df_4.to_excel(writer4, 'page_4', float_format='%.2f')
    writer4.save()
    writer4.close()


if __name__ == '__main__':
    # 自定义参数
    itemName = '云霄'  # 项目名称
    st = 4  # 调压室个数
    dict1 = {'name': '设计工况'}
    dict2 = {'name': '校核工况'}

    dict_s = {}
    dict3 = {}
    dict_steady = {}
    dict_upperGate = {}
    dict_lowerGate = {}
    dict_upperTank = {}
    dict_lowerTank = {}
    p1 = 3  # 上调
    p2 = 1  # 上闸
    p3 = 0  # 下调
    p4 = 2  # 下闸
    listN = [8, 16, 6, 6]  # 工况数列表
    # index_list = [
    #     '蜗壳末端最大动水压力', '尾水管最小压力', '转速最大上升率', '上游调压室最高涌浪', '上游调压室最低涌浪', '上闸最高涌浪',
    #     '上闸最低涌浪', '尾水调压室最高涌浪', '尾水调压室最低涌浪', '下闸最高涌浪', '下闸最低涌浪'
    # ]

    prepath = r'E:\DeskFile\项目\云霄\计算\计算\方案八\大波动'  # 计算文件夹路径
    dir_list = prepath.split('\\')
    lenth = len(dir_list)+1  # 根目录长度
    # tpFolder = os.listdir(prepath)  # 水轮机、水泵工况文件夹名称列表
    for root1, dirs1, files1 in os.walk(prepath):   # 遍历水轮机、水泵工况文件夹
        dir_list = root1.split('\\')  # 目录分段
        if('启动' in dir_list or '先甩' in dir_list):
            continue
        if (('水轮机设计工况'in dir_list) or ('水泵设计工况' in dir_list)):
            dict_s = dict1
        elif (('水轮机校核工况'in dir_list) or ('水泵校核工况' in dir_list)):
            dict_s = dict2
        else:
            continue
        if files1:
            dir2 = dir_list[lenth].split('（')[0]
            os.chdir(root1)
            Extremum(dict_s)
            UnitResult(dict3)
            jz = dict3[dir2][1]
            # SteadyFlow(dict_steady, jz)
            GateTank(dict_upperGate, dict_lowerGate,
                     dict_upperTank, dict_lowerTank, jz)

    Extremum2excel(dict1, dict2, prepath, '云霄统计结果.xls')
    UnitResult2excel(dict3,  prepath, '机组参数统计结果.xls')
    # SteadyFlow2Excek(dict_steady,  prepath, '恒定流统计结果.xls')
    GateTank2Excel(dict_upperGate, dict_lowerGate,
                   dict_upperTank, dict_lowerTank)
