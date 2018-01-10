# encoding: utf-8
import re
from lxml import etree
import xlwt
import xlrd
from xlutils.copy import copy
from user import *

def get_son(father,List):
    father_xpath = etree.HTML(etree.tostring(father))
    son=father_xpath.xpath('//body/*/*')
    if len(son) == 0:
        return False
    for i in range(len(son)):
        aa=etree.HTML(etree.tostring(father)).xpath('//body/*/text()')
        if len(aa)==3 and aa[1].strip()!='':
            if List[-1]!=aa[0].strip()+'\r\n'+aa[1].strip():
                List.pop()
                List.append(aa[0].strip()+'\r\n'+aa[1].strip())
        if son[i].text != None:
            item = re.split(r'\s+?', son[i].text)
            if len(item)==1:
                # print item[0]#时间
                List.append(item[0])
            if len(item)==9:
                # print item[4],#空课
                List.append(item[4])
            if len(item) == 10:
                # print item[9],#课程
                List.append(item[9])
            if len(item) == 17:
                # print item[8],#星期
                List.append(item[8])
        if get_son(son[i],List):
            print
    return List

def write_schedule(List):
    m = 0
    data = xlrd.open_workbook(user_xlsx(user(0)))
    wb = copy(data)
    sheet = wb.get_sheet(0)

    #写课程
    for i in range(6):
        for j in range(7):
            sheet.write_merge(i*2+1,i*2+2,j+1,j+1,List[m])
            m+=1

    #单元格长宽
    sheet.col(0).width = 5000
    for i in range(7):
        sheet.col(i+1).width = 11000
    for i in range(13):
        sheet.row(i).height = 400
    wb.save(user_xlsx(user(0)))

def cut(List):
    row, col = 1, 0
    pop_num = []
    excel = xlwt.Workbook()
    sheet = excel.add_sheet('schedule')

    List.pop(8)
    List.pop(8)
    List.pop(40)
    List.pop(40)
    List.pop(40)
    List.pop(72)
    List.pop(72)
    List.pop(72)
    # 写星期/节数
    for k in range(len(List)):
        if k < 8:
            sheet.write(0, col, List[k])
            col += 1
            pop_num.append(k)
        if k > 7 and k % 8 == 0:
            sheet.write(row, 0, List[k])
            row += 1
            pop_num.append(k)
    for i in range(len(pop_num)):
        List.pop(pop_num[i] - i)
    #去重
    pop_num=[]
    for i in range(len(List)):
        if i>6 and List[i]==List[i-7] and (i/7)%2!=0:
            pop_num.append(i)
    for i in range(len(pop_num)):
        List.pop(pop_num[i]-i)
    excel.save(user_xlsx(user(0)))
    return List