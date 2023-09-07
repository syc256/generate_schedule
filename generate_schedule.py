import xlrd  # 用来读取xls文件,不能修改数据
import xlwt  # 创建xls文件并对其进行操作，但不能对已有的xls文件进行修改
import os
import re
from unicodedata import name


def find_pattern_line(text, pattern):
    lines = text.split('\n')
    matching_lines = []

    for line in lines:
        if re.search(pattern, line):
            matching_lines.append(line)

    return matching_lines


# 单周课表
oddSpare = [
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
]

# 双周课表
evenSpare = [
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
    [[''], [''], [''], [''], [''], [''], ['']],
]

# 获取文件夹中所有.xls文件
current_folder = os.path.dirname(os.path.abspath(__file__))
file_dir = current_folder
L = []
for i, j, k in os.walk(file_dir):
    for file in k:
        if os.path.splitext(file)[1] == '.xls':
            L.append(os.path.join(file))
print(L)

# 按顺序处理各个课表有无课数据
for names in L:
    workbook = xlrd.open_workbook(names)  # 打开个人课表
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
    row = 0
    for n in schedule_row:
        thisRow = table.row_values(n)  # 一行中的课程列表
        row += 1
        for i in range(1, 8):
            pattern = r'\(\[周\]\)'  # 匹配 "([周])"
            matchedLines = find_pattern_line(thisRow[i], pattern)
            if '形势与政策' in thisRow[i]:
                print('跳过形势与政策')  # 形势与政策不算
                oddSpare[row][i - 1].append(name + ' ')
                # oddSpare[row][i - 1].append(' ')
                evenSpare[row][i - 1].append(name + ' ')
                # evenSpare[row][i - 1].append(' ')
            elif len(matchedLines) < 1:
                print("此单元格内没有([周])")
                oddSpare[row][i - 1].append(name)
                oddSpare[row][i - 1].append(' ')
                evenSpare[row][i - 1].append(name)
                evenSpare[row][i - 1].append(' ')
            else:
                weekList = []
                remaining_list = []
                for weekStringTemp in matchedLines:
                    # weekStringTemp = matchedLines[0]
                    weekString = weekStringTemp[:-5]
                    print(weekString)
                    weekListTemp = weekString.split(',')
                    print(weekListTemp)

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
                remaining_list = [x for x in range(1, 17) if x not in int_numbers]
                print(remaining_list)

                # 判断单双周的划分
                odd_count = 0
                even_count = 0
                for num in remaining_list:
                    if num % 2 == 0:
                        even_count += 1
                    else:
                        odd_count += 1
                print('odd is', odd_count)
                print('even is', even_count)
                if odd_count + even_count <= 4:  # 所有周次加起来小于等于4节的课不算
                    print('该课程数量小于等于4')
                else:
                    if odd_count > 2:  # 单周课程大于两节
                        oddSpare[row][i - 1].append(name)
                        oddSpare[row][i - 1].append(' ')
                        print('添加到单周')
                    if even_count > 2:  # 双周课程大于两节
                        evenSpare[row][i - 1].append(name)
                        evenSpare[row][i - 1].append(' ')
                        print('添加到双周')

print('单周：')
for i in range(6):
    print(oddSpare[i])
print('双周：')
for i in range(6):
    print(evenSpare[i])

# 创建一个新的 xls 文件
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Sheet1')
worksheet.col(0).width = 3000
for i in range(1, 8):
    worksheet.col(i).width = 7000

# 设置边框样式
borders = xlwt.Borders()
borders.left = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.bottom = xlwt.Borders.THIN
style = xlwt.XFStyle()
style.borders = borders
style.alignment.wrap = 1

# 合并 A1~H1，写入“单周”
worksheet.write_merge(0, 0, 0, 7, "单周", style)

# 合并 A10~H10，写入“双周”
worksheet.write_merge(7, 7, 0, 7, "双周", style)

weekday = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
session = ['第一大节', '第二大节', '第三大节', '第四大节', '晚课']
s = 0
for i in weekday:
    s += 1
    worksheet.write(1, s, i, style)
    worksheet.write(8, s, i, style)

s = 1
for i in session:
    s += 1
    worksheet.write(s, 0, i, style)
    worksheet.write(s + 7, 0, i, style)

# 将spare列表储存的数据导入xls文件
for x in range(1, 8):
    for y in range(2, 7):
        worksheet.write(y, x, oddSpare[y - 1][x - 1], style)
    for y in range(9, 14):
        worksheet.write(y, x, evenSpare[y - 8][x - 1], style)

workbook.save('无课表.xls')  # 保存无课表
print("xls 文件已生成")
