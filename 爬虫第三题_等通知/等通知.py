###刘易哲###
##202200130103##
import requests
import re
import xlwt
import xlrd
from bs4 import BeautifulSoup
import os
import time
import lxml

#path是用来存放脚本的路径
path = '.\\等通知.\\script.vbs'


#UA伪装及url
headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36'
}
url = 'https://www.bkjx.sdu.edu.cn/sanji_list.jsp'
frequency = input("请输入每几分钟刷新一次:")

#循环，每次循环sleep五分钟   此循环一直到末尾
while True:

    # 如果不是第一次运行(之前已经建立了文件夹),就读取上一次表格的第一个url，方便与之后的url对比，判断有无新通知
    first_url = ''
    if os.path.exists('./等通知./近期通知.xls'):
        workbook = xlrd.open_workbook('./等通知./近期通知.xls')
        sheet = workbook.sheet_by_index(0)
        first_url = sheet.cell_value(1, 1)

    # 初始化一个excel文件
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('can can 信息', cell_overwrite_ok=True)
    sheet.col(0).width = 256 * 70
    sheet.col(1).width = 256 * 40
    sheet.col(2).width = 256 * 15
    col = ('标题', '网址', '日期')
    for i in range(0, 3):
        sheet.write(0, i, col[i])

#爬取内容
    #for page in range(1,2):  此处可以进行翻页操作爬取更多通知，将信息存入列表后一起写入excel。但情景是要最新的通知，所以我就不写了。（才不是因为我懒呢）
    page = 1
    params = {
        'totalpage': '154',
        'PAGENUM': page,
        'urltype': 'tree.TreeTempUrl',
        'wbtreeid': '1010'
    }
    page_text = requests.get(url=url,headers=headers,params=params).text
    soup = BeautifulSoup(page_text,'lxml')
    ex = '<div style="float:right;">(.*?)</div>'
    try:
        for i in range(0,15):
            news_url = soup.select('.leftNews3 a')[i]['href']
            if re.findall('https:(.*?)',news_url,re.S) == []:
                news_url = 'https://www.bkjx.sdu.edu.cn/'+news_url
            news_title = soup.select('.leftNews3 a')[i]['title']
            news_date = re.findall(ex,page_text)[i]

            #爬取后写入excel并保存到等通知文件夹
            sheet.write(i + 1, 0, news_title)
            sheet.write(i + 1, 1, news_url)
            sheet.write(i + 1, 2, '  '+news_date)
            print(news_title)
            print(news_url)
            print(news_date)
        if not os.path.exists('./等通知'):
            os.mkdir('./等通知')
        save_path = './等通知./近期通知.xls'
        book.save(save_path)

        #如果写入数据时用户已经打开excel会出错，此时提醒用户
    except:
        with open(path, 'w') as fp:
            fp.write('msgbox"excel已打开，无法写入",48,"attention"')
        os.system(path)
        print("excel已打开，无法写入")
        continue


#读取现在excel中的数据与之前的数据对比，找出有几条新通知
    workbook2 = xlrd.open_workbook('./等通知./近期通知.xls')
    sheet_read = workbook2.sheet_by_index(0)
    count = 1
    while True:
        #第一次打开直接跳过
        if first_url == '':
            count = -1
            break
        first_url2 = sheet_read.cell_value(count, 1)
        first_title = sheet_read.cell_value(count, 0)
        if first_url2 != first_url:
            count = count+1
        else:
            break
    if count == 1:
        print('-------暂无新通知-------')
    elif count == -1:
        print('首次使用，消息已存储到./等通知./近期通知.xls')


    # 写脚本提醒有新通知
    else:
        cnt = count #cnt存放新通知的条数
        with open(path,'w') as fp:
            fp.write('dim ws'+'\n'+'set ws = createobject("wscript.shell")'+'\n')
        with open(path,'a') as fp:
            while count > 1:
                first_url2 = sheet_read.cell_value(count, 1)
                first_title = sheet_read.cell_value(count, 0)
                fp.write('u'+str(count)+'="'+first_url2+'"\n')
                fp.write('t' + str(count) + '="' + first_title+'"\n')
                count -= 1
            count = cnt
            fp.write('a = msgbox("有新通知!!!"')
            while count > 1:
                fp.write('&(chr(13))&t'+str(count)+'&(chr(13))&'+'u'+str(count))
                count -= 1
            fp.write('&(chr(13))&"是否打开全部网址?",4,"新通知")'+'\n'+'if a=6 then'+'\n')
            count = cnt
            while count > 1:
                fp.write('ws.run u'+str(count)+'\n'+'wscript.sleep 500\n')
                count -= 1
            fp.write("end if")
        fp.close()
        os.system(path)
    time.sleep(60*float(frequency))


