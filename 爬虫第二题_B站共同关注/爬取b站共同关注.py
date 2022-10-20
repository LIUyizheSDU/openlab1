###刘易哲###
##202200130103##
import requests
import re
import xlwt
import os


# 初始化一个excel文件
book = xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet = book.add_sheet('共同关注',cell_overwrite_ok=True)
col = ('UP主的UID','UP主的昵称','等级','UP主的粉丝数(个)')
for i in range(0,4):
    sheet.write(0,i,col[i])
sheet.col(0).width = 256*15
sheet.col(1).width = 256*20
sheet.col(2).width = 256*6
sheet.col(3).width = 256*16

# UA伪装
headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36'
}

# 建立列表，存储爬取的信息
Ulist1 = []
Ulist2 = []
nameList = []
levelList = []
followerList = []
uid1 = input("请输入uid1:")
uid2 = input("请输入uid2:")

# 该函数获取某个uid的关注列表中的up主的uid
def GetUid(Ulist,uid):
    url = 'https://api.bilibili.com/x/relation/followings'
    for pn in range(1,6):
        params = {
        'vmid': uid,
        'pn': pn,
        'ps': '20',
        }
        page_text = requests.get(url=url,headers=headers,params=params).text
        ex_uid = '"mid":(.*?),'
        uid_list = re.findall(ex_uid,page_text,re.S)
        for uid_src in uid_list:
            Ulist.append(uid_src)

# 该函数获取某个uid的昵称和等级信息，并将其存入列表
def get_info(uid,NameList,LevelList):
    url = 'https://api.bilibili.com/x/space/acc/info?mid='+uid
    page_info = requests.get(url=url,headers=headers).text
    ex_level = '"level":(.*?),"jointime"'
    ex_name = '"name":"(.*?)","sex"'
    levelList = re.findall(ex_level,page_info,re.S)
    nameList = re.findall(ex_name,page_info,re.S)
    if levelList == []:
        levelList.append('error')
    if nameList == []:
        nameList.append('该账号已注销')
    NameList.append(nameList[0])
    LevelList.append(levelList[0])


# 该函数获取某个uid的的粉丝数并存入列表
def get_followers(uid,FollowerList):
    url = 'https://api.bilibili.com/x/relation/followers'
    params = {
    'vmid': uid,
    'pn': '1',
    'ps': '20'
    }
    page_followers = requests.get(url=url,headers=headers,params=params).text
    ex_followers = '"total":(.*?)}'
    follower_list = re.findall(ex_followers,page_followers,re.S)
    if follower_list == []:
        follower_list.append('error')
    FollowerList.append(follower_list[0])

#爬取时间太长辣，让用户等等，防止他关掉程序
print("Please wait...")

# 获取两个uid的共同关注列表，取交集
GetUid(Ulist1,uid1)
GetUid(Ulist2,uid2)
UList = [id for id in Ulist1 if id in Ulist2]
print(UList)


for id in UList:
    #print(id)
    get_info(id,nameList,levelList)
    get_followers(id,followerList)

    # 写入execl并保存
l = len(UList)
for i in range(0,l):
    sheet.write(i + 1, 0, UList[i])
    sheet.write(i + 1, 1, nameList[i])
    sheet.write(i + 1, 2, levelList[i])
    sheet.write(i + 1, 3, followerList[i])
if not os.path.exists('./B站共同关注'):
    os.mkdir('./B站共同关注')
save_path = './B站共同关注./'+uid1+'和'+uid2+'的共同关注.xls'
book.save(save_path)
print('done.')
