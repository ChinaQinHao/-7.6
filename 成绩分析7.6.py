import xlrd
import matplotlib.pyplot as plt
# import numpy as np
from matplotlib import font_manager
# 导入模块
import os

plt.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文显示问题
plt.rcParams['axes.unicode_minus'] = False  # 解决中文显示问题

my_font = font_manager.FontProperties(fname="C:\Windows\Fonts\msyh.ttc")
work_book = xlrd.open_workbook('成绩单.xls')  # 读取成绩单

x_axis_data = []

y_axis_data = []
sheet_1 = work_book.sheet_by_index(0)
'''
student=[]
colors = ['bs--', 'yo--', 'r*--', 'c+--', 'g^--']
name = ['尹永杰', '宋志楠', '秦昊', '李勇昊', '柏昕彤']


for i in range(1,6):
    for j in range(1,9):
        cell_0 = sheet_1.cell_value(i, j)
        student.append(cell_0)
    y_axis_data.append(student)
    student=[]
for i in range(1,9):
    cell_1 = sheet_1.cell_value(0, i)
    x_axis_data.append(cell_1)

print(x_axis_data)
print(y_axis_data)



for i in range(5):
    plt.plot(x_axis_data, y_axis_data[i], colors[i],
             alpha=0.5, linewidth=1, label=(name[i]))
    # 'bo-'表示蓝色实线，数据点实心原点标注
## plot中参数的含义分别是横轴值，纵轴值，线的形状（'s'方块,'o'实心圆点，'*'五角星   ...，颜色，透明度,线的宽度和标签 ，

plt.legend()  # 显示上面的label
plt.xlabel('科目',fontproperties=my_font)  # x_label
plt.ylabel('成绩', fontproperties=my_font)  # y_label


plt.xticks(x_axis_data, fontproperties=my_font)

plt.legend(prop=my_font)
#plt.ylim(-1,1)#仅设置y轴坐标范围
plt.show()

'''


# 创建文件夹
def 创建文件夹(path):
    folder = os.path.exists(path)

    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径

        print("---  创建文件夹成功  ---")

    else:
        print("---  文件夹被创建过了！  ---")


def 显示并保存图片():
    for i in range(9):

        fraces = [58 - 个人各科成绩排名[i], 个人各科成绩排名[i]]
        plt.axes(aspect=1)
        plt.pie(x=fraces, autopct='%3.1f%%', labels=['您超过的', '全班人数'])
        plt.title(科目name[i] + '超过全班')
        photo = '{1}/{0}{1}.png'.format(科目name[i], username)
        plt.savefig(photo)  # 保存图片
        # plt.show()
        plt.close()


# 按列读取
data_col = [sheet_1.col_values(i) for i in range(sheet_1.ncols)]
# 按行读取
data_row = []
for row in range(sheet_1.nrows):
    data_row.append(sheet_1.row_values(row))
# print(data_col)
# 设置全班名字
data_row[0].pop(0)
allname = data_col[0]
allname.pop(0)
科目 = ['yuwen', 'shuxue', 'yingyu', 'zhengzhi', 'lishi', 'dili', 'shenngwu', 'wuli','zongfen']
科目name = ['语文', '数学', '英语', '政治', '历史', '地理', '生物', '物理', '总分']
各科成绩排名 = []
科目字典 = ['语文', '数学', '英语', '政治', '历史', '地理', '生物', '物理', '总分']
for i in range(9):
    # 语文成绩单
    科目[i] = data_col[i + 1]
    科目[i].pop(0)
    各科成绩排名.append(sorted(科目[i], reverse=True))
    # 把姓名与语文成绩结合成字典
    科目[i] = zip(allname, 科目[i])
    # yuwen.sort(reverse = True)
    科目[i] = dict(科目[i])
    科目字典[i] = dict(科目[i])
    科目[i] = sorted(科目[i].items(), key=lambda x: x[1], reverse=True)


    # print(科目name[i],科目[i])

# print(各科成绩排名)

个人各科成绩 = ['语文', '数学', '英语', '政治', '历史', '地理', '生物', '物理','总分']
个人各科成绩排名 = []
username = input("请输入您的姓名")

if username in allname:
    print('查询到以下结果')
    for i in range(9):
        #print(科目字典)
        个人各科成绩[i] = 科目字典[i].get(username)
        个人各科成绩排名.append(各科成绩排名[i].index(个人各科成绩[i]))
        print(科目name[i], 个人各科成绩[i])

        # 创建文件夹
        # 调用函数
    file = "{0}".format(username)
    创建文件夹(file)
    显示并保存图片()  # 这是个函数！！！！！！！！！！在前面定义的























else:
    print('对不起，没有找到相关学生')


# print(个人各科成绩排名)






