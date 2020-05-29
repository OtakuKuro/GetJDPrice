import urllib.request as urllib2
import json
import xlwt
import xlrd
import os
from xlutils.copy import copy
import time
import random

def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

def check_xls():
    if os.path.exists("test.xls"):  
        pass
    else:
        f = xlwt.Workbook()
        sheet1 = f.add_sheet('price',cell_overwrite_ok=True)
        row0 = ["时间","CPU","GPU","内存", "固态", "机械", "散热", "电源", "机箱", "总计"]

        #写第一行
        for i in range(0, len(row0)):
            sheet1.write(0,i,row0[i],set_style('Times New Roman',220))

        f.save('test.xls')


def get_price():
    xlsrd = xlrd.open_workbook('test.xls')
    rs = xlsrd.sheet_by_index(0)
    row = rs.nrows
    
    xlswt = copy(xlsrd)
    sheet = xlswt.get_sheet(0)
    
    sum_price = 0

    table_date = rs.cell_value(row-1, 0)
    #print (table_date)
    current_date = time.strftime("%m-%d", time.localtime())
    if table_date == current_date:
        row -= 1
    sheet.write(row,0,current_date)

    cpu = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuIds=J_100004330867',timeout=5)
    price=json.loads(cpu.read())
    cpu.close()
    sheet.write(row,1,price[0]['p'])
    sum_price += float(price[0]['p'])

    gpu = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuIds=J_100005239235',timeout=5)
    price=json.loads(gpu.read())
    gpu.close()
    sheet.write(row,2,price[0]['p'])
    sum_price += float(price[0]['p'])

    memory = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuid=J_100003138151',timeout=5)
    price=json.loads(memory.read())
    memory.close()
    sheet.write(row,3,price[0]['p'])
    sum_price += float(price[0]['p'])*2

    ssd = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuid=J_100005926989',timeout=5)
    price=json.loads(ssd.read())
    ssd.close()
    sheet.write(row,4,price[0]['p'])
    sum_price += float(price[0]['p'])

    disk = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuid=J_3843702',timeout=5)
    price=json.loads(disk.read())
    disk.close()
    sheet.write(row,5,price[0]['p'])
    sum_price += float(price[0]['p'])

    fan = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuid=J_598827',timeout=5)
    price=json.loads(fan.read())
    fan.close()
    sheet.write(row,6,price[0]['p'])
    sum_price += float(price[0]['p'])

    power = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuid=J_100004925348',timeout=5)
    price=json.loads(power.read())
    power.close()
    sheet.write(row,7,price[0]['p'])
    sum_price += float(price[0]['p'])

    case = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuid=J_100006587871',timeout=5)
    price=json.loads(case.read())
    case.close()
    sheet.write(row,8,price[0]['p'])
    sum_price += float(price[0]['p'])

    sheet.write(row,9,sum_price)

    xlswt.save('test.xls')

if __name__ == "__main__":
    check_xls()
    get_price()