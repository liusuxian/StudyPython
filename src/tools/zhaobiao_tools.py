# -*- coding: utf-8 -*-
import calendar
import datetime
import os
import time
import re
from docx import Document
import requests

cityList = [
    '成都', '四川',
    '潍坊', '山东',
    '唐山', '河北',
    '西安', '陕西',
    '南京', '江苏'
]


def handelDateUrl(date: str):
    dateUrl = "https://www.bidcenter.com.cn/newsmore-" + date + '.html'
    print('dateUrl:', dateUrl)
    result = requests.get(dateUrl)
    if result.status_code == 200:
        resUrlList = re.findall('<li><a href="(.*?)".*?>.*?</a>', result.text, re.S)
        for resUrl in resUrlList:
            handelResUrl('https://www.bidcenter.com.cn' + resUrl)
    else:
        print('handelDateUrl ERROR:', result.status_code, dateUrl)


def handelResUrl(resUrl: str):
    result = requests.get(resUrl)
    if result.status_code == 200:
        realUrl = re.findall('<link rel.*?href="(.*?)" />', result.text, re.S)
        if len(realUrl) > 0:
            time.sleep(1)
            handelRealUrl(realUrl[0])
    else:
        print('handelResUrl ERROR:', result.status_code, resUrl)


def handelRealUrl(realUrl: str):
    result = requests.get(realUrl)
    if result.status_code == 200:
        city = re.findall('<ul class="xiangm-xx">.*?<li>.*?<a .*?>(.*?)</a>', result.text, re.S)
        if len(city) > 0:
            city = city[0]
            for c in cityList:
                if c in city:
                    title = re.findall('<p class="text-title">(.*?)</p>', result.text, re.S)
                    if len(title) > 0:
                        title = title[0].replace("\r\n", "").replace(" ", "")
                        print('title:', title, 'city:', city, 'realUrl:', realUrl)
                        p = document.add_paragraph(title + '\n' + city + '\n' + realUrl + '\n')
                        p.add_run('bold').bold = True
                        p.add_run('and some')
                        p.add_run('italic.').italic = True
                        # 保存文档
                        document.save('url.docx')
                    break
    else:
        print('handelRealUrl ERROR:', result.status_code, realUrl)


# 创建 docx 文件
isExists = os.path.exists('url.docx')
if isExists:
    os.remove('url.docx')
# 创建文档对象
document = Document()
# 设置文档标题，中文要用unicode字符串
document.add_heading(u'招标公告信息', 0)

# 爬取最近6个月的数据
today = datetime.date.today()
endYear = today.year
endMonth = today.month
endDay = today.day
print('today:', today)
curYear = endYear
curMonth = endMonth
for i in range(0, 6):
    print('curYear:', curYear, 'curMonth:', curMonth)
    if curMonth == endMonth:
        for day in range(1, endDay + 1):
            handelDateUrl(str(curYear) + '-' + str(curMonth) + '-' + str(day))
    else:
        _, monthRange = calendar.monthrange(curYear, curMonth)
        for day in range(1, monthRange + 1):
            handelDateUrl(str(curYear) + '-' + str(curMonth) + '-' + str(day))

    curMonth = curMonth - 1
    if curMonth <= 0:
        curYear = curYear - 1
        curMonth = 12
