#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
#from openpyxl import Workbook
'''
import os

os.chdir('D:\\workspace\\Python36\\test')

def quDeGongZuoBiao(gongZuoBuMingCheng,gongZuoBiaoMingCheng):
    #取得excel的xlsx工作簿文件中的工作表并返回，返回值为openpyxl.worksheet.worksheet.Worksheet类型
    #需要import openpyxl。2.4.5版本openpyxl不支持get_highest_row()方法，引入max_row,max_column属性替代。
    jieGuo=openpyxl.load_workbook(gongZuoBuMingCheng).get_sheet_by_name(gongZuoBiaoMingCheng)
    return jieGuo
'''
def jiaoYiCelue(TRiShouPanJia,TRiMA20Zhi,T_1RiShouPanJia,T_1RiMA20Zhi,T_2RiShouPanJia,T_2RiMA20Zhi):
    '''
    当t日收盘价<t日ma20 并且 t-1日收盘价<t-1日ma20时，策略为空仓
    当t日收盘价<t日ma20 并且 t-1日收盘价>t-1日ma20时，策略为卖出
    当t日收盘价<t日ma20 并且 t-1日收盘价=t-1日ma20 并且 t-2日收盘价<t-2日ma20时，策略为空仓
    当t日收盘价<t日ma20 并且 t-1日收盘价=t-1日ma20 并且 t-2日收盘价>t-2日ma20时，策略为卖出
    -------------------------------------------------------------------------------------
    当t日收盘价>t日ma20 并且 t-1日收盘价<t-1日ma20时，策略为买入
    当t日收盘价>t日ma20 并且 t-1日收盘价>t-1日ma20时，策略为持有
    当t日收盘价>t日ma20 并且 t-1日收盘价=t+1日ma20 并且 t-2日收盘价<t-2日ma20时，策略为买入
    当t日收盘价>t日ma20 并且 t-1日收盘价=t+1日ma20 并且 t-2日收盘价>t-2日ma20时，策略为持有
    -------------------------------------------------------------------------------------
    当t日收盘价=t日ma20时，策略为不变
    -------------------------------------------------------------------------------------
    '''
    if TRiShouPanJia < TRiMA20Zhi:
        if T_1RiShouPanJia < T_1RiMA20Zhi:
            jieGuo='空仓'
        elif T_1RiShouPanJia > T_1RiMA20Zhi:
            jieGuo='卖出'
        elif T_1RiShouPanJia == T_1RiMA20Zhi:
            if T_2RiShouPanJia < T_2RiMA20Zhi:
                jieGuo='空仓'
            elif T_2RiShouPanJia > T_2RiMA20Zhi:
                jieGuo='卖出'
            else:jieGuo='1.1.1错误'
        else:
            jieGuo = '1.1错误'
    elif TRiShouPanJia > TRiMA20Zhi:
        if T_1RiShouPanJia < T_1RiMA20Zhi:
            jieGuo = '买入'
        elif T_1RiShouPanJia > T_1RiMA20Zhi:
            jieGuo = '持有'
        elif T_1RiShouPanJia == T_1RiMA20Zhi:
            if T_2RiShouPanJia < T_2RiMA20Zhi:
                jieGuo = '买入'
            elif T_2RiShouPanJia > T_2RiMA20Zhi:
                jieGuo = '持有'
            else:
                jieGuo = '1.1.2错误'
        else:
            jieGuo = '1.2错误'
    elif TRiShouPanJia == TRiMA20Zhi:
        jieGuo='维持'
    else:jieGuo = '1错误'
    return jieGuo
#------------------------------------------------------------
def zhiXingJiaoYi(TRiJiaoYiCeLue,T_1RiCangWeiShuLiang,T_1RiZhangHuYuE,guJia):
    '''
    当t日策略为“买入”并且t-1日仓位数量=0时，t日执行买入操作，既持仓数+1并且账户余额=账户余额额-股价
    当t日策略为“买入”并且t-1日仓位数量<>0时，t日报告错误
    当t日策略为“卖出”并且t-1日仓位数量>0时，t日执行卖出操作既持仓数-1并且账户余额=账户余额额+股价
    当t日策略为“卖出”并且t-1日仓位数量<=0时，t日报告错误
    当t日策略为“持有”并且t-1日仓位数量>0时，t日执行持有操作既持仓数不变并且账户余额不变
    当t日策略为“持有”并且t-1日仓位数量<=0时，t日报告错误
    当t日策略为“空仓”并且t-1日仓位数量=0时，t日执行空仓操作既持仓数不变并且账户余额不变
    当t日策略为“空仓”并且t-1日仓位数量<>0时，t日报告错误
    当t日策略为“维持”时，t日持仓数不变
    其他情况报告错误
    '''
    jieGuo2=0
    if TRiJiaoYiCeLue == '买入':
        if T_1RiCangWeiShuLiang == 0:
            jieGuo = T_1RiCangWeiShuLiang+1
            jieGuo2=T_1RiZhangHuYuE-guJia
        else:jieGuo = '2.1错误'
    elif TRiJiaoYiCeLue == '卖出':
        if T_1RiCangWeiShuLiang > 0:
            jieGuo = T_1RiCangWeiShuLiang - 1
            jieGuo2 = T_1RiZhangHuYuE + guJia
        else:jieGuo = '2.2错误'
    elif TRiJiaoYiCeLue == '持有':
        if T_1RiCangWeiShuLiang > 0:
            jieGuo = T_1RiCangWeiShuLiang
            jieGuo2 = T_1RiZhangHuYuE
        else:jieGuo = '2.3错误'
    elif TRiJiaoYiCeLue == '空仓':
        if T_1RiCangWeiShuLiang == 0:
            jieGuo = T_1RiCangWeiShuLiang
            jieGuo2 = T_1RiZhangHuYuE
        else:jieGuo = '2.4错误'
    elif TRiJiaoYiCeLue == '维持':
        jieGuo=T_1RiCangWeiShuLiang
        jieGuo2 = T_1RiZhangHuYuE
    else: jieGuo = '2错误'
    return (jieGuo,jieGuo2)
#------------------------------------------------------------
def zhiXingJiaoYi2(TRiJiaoYiCeLue,T_1RiCangWeiShuLiang):
    '''
    当t日策略为“买入”并且t-1日仓位数量=0时，t日执行买入1
    当t日策略为“买入”并且t-1日仓位数量<>0时，t日报告错误
    当t日策略为“卖出”并且t-1日仓位数量>0时，t日执行卖出操作
    当t日策略为“卖出”并且t-1日仓位数量<=0时，t日报告错误
    当t日策略为“持有”并且t-1日仓位数量>0时，t日执行持有操作
    当t日策略为“持有”并且t-1日仓位数量<=0时，t日报告错误
    当t日策略为“空仓”并且t-1日仓位数量=0时，t日执行空仓操作
    当t日策略为“空仓”并且t-1日仓位数量<>0时，t日报告错误
    其他情况报告错误
    '''
    if TRiJiaoYiCeLue == '买入':
        if T_1RiCangWeiShuLiang == 0:
            jieGuo=T_1RiCangWeiShuLiang+1
        else:jieGuo = '2.1错误'
    elif TRiJiaoYiCeLue == '卖出':
        if T_1RiCangWeiShuLiang > 0:
            jieGuo = T_1RiCangWeiShuLiang - 1
        else:jieGuo = '2.2错误'
    elif TRiJiaoYiCeLue == '持有':
        if T_1RiCangWeiShuLiang > 0:
            jieGuo = T_1RiCangWeiShuLiang
        else:jieGuo = '2.3错误'
    elif TRiJiaoYiCeLue == '空仓':
        if T_1RiCangWeiShuLiang == 0:
            jieGuo = T_1RiCangWeiShuLiang
        else:jieGuo = '2.4错误'
    else: jieGuo = '2错误'
    return jieGuo
#------------------------------------------------------------
def jiSuanYuE(TRiCangWeiShuLiang,T_1RiZhangHuYuE,guJia):
    if TRiCangWeiShuLiang == 1 or TRiCangWeiShuLiang == -1:
        jieGuo=T_1RiZhangHuYuE-guJia
    else:jieGuo = T_1RiZhangHuYuE
    return jieGuo
#------------------------------------------------------------
print('打开工作表')
#gongZuoBiao=quDeGongZuoBiao('510050.xlsx','510050')

gongZuoBu=openpyxl.load_workbook('510050.xlsx')
gongZuoBiao=gongZuoBu.get_sheet_by_name('510050')
print('单元格读取中……')
for chuShiZhi in range(24,400,1):
    gongZuoBiao['m'+str(chuShiZhi)].value = jiaoYiCelue(gongZuoBiao['e'+str(chuShiZhi)].value, gongZuoBiao['i'+str(chuShiZhi)].value, gongZuoBiao['e'+str(chuShiZhi-1)].value,
                                           gongZuoBiao['i'+str(chuShiZhi-1)].value, gongZuoBiao['e'+str(chuShiZhi-2)].value, gongZuoBiao['i'+str(chuShiZhi-2)].value)
    gongZuoBiao['n' + str(chuShiZhi)].value,gongZuoBiao['o' + str(chuShiZhi)].value = zhiXingJiaoYi(gongZuoBiao['m'+str(chuShiZhi)].value,gongZuoBiao['n'+str(chuShiZhi-1)].value,gongZuoBiao['o'+str(chuShiZhi-1)].value,gongZuoBiao['e'+str(chuShiZhi)].value)
    #gongZuoBiao['o' + str(chuShiZhi)].value = jiSuanYuE(gongZuoBiao['n'+str(chuShiZhi)].value,gongZuoBiao['o'+str(chuShiZhi-1)].value,gongZuoBiao['e'+str(chuShiZhi)].value,)
#gongZuoBiao['m24'].value=jiaoYiCelue(gongZuoBiao['e24'].value,gongZuoBiao['i24'].value,gongZuoBiao['e23'].value,gongZuoBiao['i23'].value,gongZuoBiao['e22'].value,gongZuoBiao['i22'].value)
#gongZuoBiao['m46'].value=jiaoYiCelue(gongZuoBiao['e46'].value,gongZuoBiao['i46'].value,gongZuoBiao['e45'].value,gongZuoBiao['i45'].value,gongZuoBiao['e44'].value,gongZuoBiao['i44'].value)
'''
print(gongZuoBiao['m22'].value)
riQi=gongZuoBiao['a22'].value
shouPanJia=gongZuoBiao['e22'].value
MA20=gongZuoBiao['i22'].value
print(riQi,shouPanJia,MA20)
if shouPanJia>MA20:
    gongZuoBiao['m22']=1
elif shouPanJia==MA20:
    gongZuoBiao['m22'] = 0
elif shouPanJia < MA20:
    gongZuoBiao['m22'] = -1
else:
    gongZuoBiao['m22'] = '错误'
print(gongZuoBiao['m22'].value)
gongZuoBiao['N22'].value = '误'


def secondvalue(a, b):
    c = a + b
    return (a, b, c)


x, y, z = secondvalue(1, 2)
print( 'x:', x, 'y:', y, 'z:', z)
'''
gongZuoBu.save('111.xlsx')



print('保存完成。')


