import xlrd
from xlutils.copy import copy
import requests
import time
import json
import pprint

# 获得根据地址返回的json字符串
def get_mapLngLat(addr):
    url = 'http://api.map.baidu.com/geocoder/v2/?address={}&output=json&ak=T9wg1BibN8VMVEDudgyh6RkqwEvcGQjF'.format(
        addr)
    res = requests.get(url)
    # 转换为json串
    json_data = json.loads(res.text)
    return json_data

# 循环地址得到经纬度
def getLngLat(addr):
    # 定义空数组，用来存放经纬度，作为返回值。供写入文件时调用
    long_lat = []
    for item in addr:
        if item == "":
            long_lat.append(None)
        else:
            json_data = get_mapLngLat(item)
            # pprint.pprint(json_data)  基本与print相等，只不过打印出的数据结构更清晰
            long = json_data['result']['location']['lng']
            lat = json_data['result']['location']['lat']
            long_lat.append(str(long) + ',' + str(lat))
            print(str(long) + ',' + str(lat))
            # time.sleep(0.5)
    return long_lat


# 打开excel文件，路径看你自己文件在哪里了
excel = xlrd.open_workbook("fileSource/四个阶段清单列表.xlsx")
# 第一个sheet页
sheet = excel.sheets()[0]
#获取列数据 第7列，第二行开始(这里是excel中的中文地址)
col = sheet.col_values(6,1)

# 调用遍历获取经纬度方法，返回值为一个数组
longLat = getLngLat(col)
# 以复制一份文件的方式写入数据，使用 xlutils.copy 中的copy方法
fileCopy = copy(excel)
#获取到复制的文件的第一个sheet页
csheet = fileCopy.get_sheet(0)
# 定义行的变量，用于遍历写入使用
rows = 1
# cols = 7
# 循环写入数据
for val in longLat:
    # 写入数据
    csheet.write(rows, 7, val)   #行，列，值
    rows = rows+1   #行加一
# 保存文件
fileCopy.save('fileSource/test.xls')