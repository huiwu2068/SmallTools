import xlwings as xw
import re

# 创建新的Excel文件
wb = xw.Book()
sheet = wb.sheets[0]  # 获取第一个工作表

# 写入表头
header = ['章节',  '项目', '描述', '优先级']
sheet.range('A1').value = header

# 从文件读取示例内容
with open('spec.txt', 'r',encoding='utf-8') as file:
    example_content = file.read()

# 初始化变量
category = ""
chapter = ""
subchapter = ""
item = ""
description = ""
priority = ""

# 将示例内容按行分割
lines = example_content.strip().split('\n')

row = 2  # 从第二行开始写入数据
for line in lines:
    line = line.strip()
    if line:
        if line.startswith('('):
            # 项目行
            item = re.sub(r'\([^)]+\)\t', '', line)
            sheet.range(f'B{row}').value = item
        elif line.startswith('优先级'):
            # 优先级行
            priority = re.sub(r'优先级：', '', line)
            sheet.range(f'D{row}').value = priority
            row += 1
            description = ""
        elif re.match(r'^\d[\d.]*\t', line):
            sheet.range(f'A{row}').value = line
            row += 1
        else:
            # 描述行
            description = description +'\n'+line
            sheet.range(f'C{row}').value = description

# 保存并关闭Excel文件
wb.save('output.xlsx')
wb.close()
