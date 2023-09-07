import xlrd  # 用来读取xls文件,不能修改数据
import xlwt  # 创建xls文件并对其进行操作，但不能对已有的xls文件进行修改
import os
import re
from unicodedata import name

totWeek = 16    # 总周数


def find_pattern_line(text, pattern):
    lines = text.split('\n')
    matching_lines = []

    for line in lines:
        if re.search(pattern, line):
            matching_lines.append(line)

    return matching_lines


Spare = []    # 创建空的无课表列表
for i in range(1, totWeek+2):
    perWeekList = [
        [[], [], [], [], [], [], [], []],
        [[], [], [], [], [], [], [], []],
        [[], [], [], [], [], [], [], []],
        [[], [], [], [], [], [], [], []],
        [[], [], [], [], [], [], [], []],
        [[], [], [], [], [], [], [], []],
    ]
    Spare.append(perWeekList)


# 获取文件夹中所有.xls文件
current_folder = os.path.dirname(os.path.abspath(__file__))
file_dir = './input'
L = []
for i, j, k in os.walk(file_dir):
    for file in k:
        if os.path.splitext(file)[1] == '.xls':
            L.append(os.path.join(file))
print(L)

# 按顺序处理各个课表有无课程数据
for names in L:
    workbook = xlrd.open_workbook(file_dir + '/' + names)  # 打开个人课表
    table = workbook.sheet_by_name('Sheet1')  # 通过名称获取sheet表
    text = table.row_values(0)[0]  # 获取表头的人名，用于区分不同同学课表
    print(text)
    pattern = r'[\u4e00-\u9fa5]+'  # 匹配汉字的范围
    matches = re.findall(pattern, text)

    if len(matches) >= 3:
        name = matches[1]
        print(name)  # 输出姓名
    else:
        print("表头格式错误，请检查xls文件是否从教务下载")

    # 按顺序读取每个人无课情况，存入列表
    schedule_row = [3, 4, 5, 6, 7]  # 原课程表表格中课程行数
    row = -1
    for n in schedule_row:
        row += 1
        thisRow = table.row_values(n)  # 一行中的课程列表

        # 提取课程列表
        for i in range(1, 8):
            pattern = r'\(\[周\]\)'  # 匹配 "([周])"
            matchedLines = find_pattern_line(thisRow[i], pattern)
            # print(row, i)
            weekList = []
            remaining_list = []
            for weekStringTemp in matchedLines:
                weekString = weekStringTemp[:-5]
                weekListTemp = weekString.split(',')

                # 提取课程周数
                for num in range(0, len(weekListTemp)):
                    if '-' in weekListTemp[num]:
                        nums = weekListTemp[num].split('-')
                        left = int(nums[0])
                        right = int(nums[1])
                        new_weeks = list(range(left, right + 1))  # 创建连续周数的新列表
                        weekList = weekList + new_weeks  # 将新列表添加到原weekList中
                    else:
                        weekList.append(int(weekListTemp[num]))
            weekList.sort()  # 排序周数
            int_numbers = [int(x) for x in weekList]

            # 获取剩余的数字组成的列表
            remaining_list = [x for x in range(1, totWeek + 1) if x not in int_numbers]
            print(remaining_list)
            for week in remaining_list:
                Spare[week][row][i-1].append(name + ' ')

for i in range(1, totWeek + 1):
    print(i)
    for j in range(6):
        print(Spare[i][j])

for i in range(1, totWeek + 1):
    # 创建一个新的 xls 文件
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Sheet1')
    worksheet.col(0).width = 3000
    for j in range(1, 8):
        worksheet.col(j).width = 7000

    # 设置边框样式
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style = xlwt.XFStyle()
    style.borders = borders
    style.alignment.wrap = 1

    weekday = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    session = ['第一大节', '第二大节', '第三大节', '第四大节', '晚课']
    # 输出表头
    s = 0
    for j in weekday:
        s += 1
        worksheet.write(0, s, j, style)

    s = 0
    for j in session:
        s += 1
        worksheet.write(s, 0, j, style)

    # 将Spare列表储存的数据导入xls文件
    for x in range(0, 5):
        for y in range(0, 7):
            worksheet.write(x+1, y+1, Spare[i][x][y], style)
    workbookPath = './output/第' + str(i) + '周.xls'
    workbook.save(workbookPath)  # 保存无课表
    print(f"第{i}周无课表已生成,文件路径{workbookPath}")
input()
