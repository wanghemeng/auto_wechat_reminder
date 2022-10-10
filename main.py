import itchat
import xlrd

msg=' 该做核酸了宝子'

workbook = xlrd.open_workbook(r'./data.xls')
sheet_name = workbook.sheet_names()[0]
sheet = workbook.sheet_by_index(0)

names = sheet.col_values(0)
dates = sheet.col_values(12)

print(names)
print(dates)

itchat.auto_login()

for i in range(len(names)):
    if dates[i] == '3':
        print(names[i])
        for friend_info in itchat.get_friends():
            RemarkName = friend_info['RemarkName']
            if (RemarkName.find(names[i]) != -1): 
                people = itchat.search_friends(name=RemarkName)
                userName = people[0]['UserName']
                itchat.send(names[i]+msg, toUserName=userName)
                break

itchat.logout()

# 下面的代码不支持模糊搜索
# for i in range(len(names)):
#     if dates[i] == '3' or dates[i] == 3:
#         print(names[i])
#         people = itchat.search_friends(name=names[i])
#         print(people)
#         userName = people[0]['UserName']
#         itchat.send(names[i]+msg, toUserName=userName)
