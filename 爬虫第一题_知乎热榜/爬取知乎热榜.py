###刘易哲###
##202200130103##
import requests
import re
import xlwt
import os
print('Please wait...')
# 初始化表格信息
book = xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet1 = book.add_sheet('小时榜',cell_overwrite_ok=True)
sheet2 = book.add_sheet('日榜',cell_overwrite_ok=True)
sheet3 = book.add_sheet('周榜',cell_overwrite_ok=True)
col = ('标题', '话题', '网址', '关注增量', '浏览增量', '回答增量', '赞同增量')
sheet1.col(7).width = 256*10
for i in range(0,7):
    sheet1.write(0, i, col[i])
    sheet2.write(0, i, col[i])
    sheet3.write(0, i, col[i])
sheet1.write(0,7,'热力值')

# 初始化表格大小
def init_size(sheet):
    sheet.col(0).width = 256*90
    sheet.col(1).width = 256*30
    sheet.col(2).width = 256*40
init_size(sheet1)
init_size(sheet2)
init_size(sheet3)


headers = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36'
}
url_hour = 'https://www.zhihu.com/api/v4/creators/rank/hot'




#以下部分为爬取各类信息的函数
def get_info(time,sheet):

    #请求含有信息的页面
    topics = ''
    params = {
        'domain': '0',
        'period': time,
        'limit': '100',
        'offset': '0'
    }
    response = requests.get(url=url_hour,headers=headers,params=params)
    page_text = response.text


#定义各类信息的正则表达式
    ex_url = '"question":{"url":"(.*?)",'
    ex_title = '"title":"(.*?)"'
    ex_hot = '"score":(.*?),"score_level"'
    ex_topicsList = '"topics":(.*?),"label"'
    ex_topic = '"name":"(.*?)"}'
    ex_new_follow = '"new_follow_num":(.*?),'
    ex_new_look = '"new_pv":(.*?),'
    ex_new_answer ='"new_answer_num":(.*?),'
    ex_new_agree = '"new_upvote_num":(.*?),'

#用正则表达式筛选信息
    url_list = re.findall(ex_url,page_text)
    title_list = re.findall(ex_title,page_text)
    #hot_list = re.findall(ex_hot,page_text)
    topics_list = re.findall(ex_topicsList,page_text)
    answer_list = re.findall(ex_new_answer,page_text)
    follow_list = re.findall(ex_new_follow,page_text)
    look_list = re.findall(ex_new_look,page_text)
    agree_list = re.findall(ex_new_agree,page_text)

#写入信息到excel
    l = len(url_list)
    for i in range(0,l):
        sheet.write(i + 1, 0, title_list[i])
        topics_pre = re.findall(ex_topic, topics_list[i])
        for j in range(0,len(topics_pre)):
            topics += '#'+topics_pre[j]
        sheet.write(i + 1, 1, topics)
        topics = ''
        sheet.write(i + 1, 2, url_list[i])
        sheet.write(i + 1, 3, follow_list[i])
        sheet.write(i + 1, 4, look_list[i])
        sheet.write(i + 1, 5, answer_list[i])
        sheet.write(i + 1, 6, agree_list[i])
    if time == 'hour':    #如果是小时榜，爬取热力值
        hot_list = re.findall(ex_hot, page_text)
        for i in range(0,l):
            sheet.write(i+1,7,hot_list[i])


get_info('hour',sheet1)
get_info('day',sheet2)
get_info('week',sheet3)
if not os.path.exists('./知乎热榜'):
    os.mkdir('./知乎热榜')
save_path = './知乎热榜./知乎榜单.xls'
book.save(save_path)
print("done")
