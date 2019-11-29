from openpyxl import *
import random


filename = 'N2核心词汇.xlsx'

# 读取.xlsx文件
wb = load_workbook(filename)

# 读取Excel里的第一张表
sheet = wb[wb.sheetnames[0]]

# # 遍历B列
# for i in range(1, sheet.max_row+1):
#     print(sheet["B"+str(i)].value)

print('--------------------------------------------------')
print('-----------答题开始，请关闭单词Excel--------------')

epoch_null = False
cols = ['B','C','D']
allAnswers = [i for i in range(2,sheet.max_row+1)]
for epoch in range(10000):
    if epoch%7 == 0 or epoch_null:
        epoch_null = False
        epoch_answers = random.sample(allAnswers, 30) 
    q = random.sample(cols, 2) 
    answers = []
    for i in epoch_answers:
        if sheet['F'+str(i)].value >= 1:
            continue
        n = max(0,sheet['G'+str(i)].value) + 1
        for j in range(n):
            answers.append(i)
    print(len(answers))
    if not answers:
        epoch_null = True
        continue
    row = random.choice(answers)
    wrongAnswers = [i for i in epoch_answers]
    del wrongAnswers[wrongAnswers.index(row)]#删除正确的答案
    wrongAnswers=random.sample(wrongAnswers,3)#随机选取三个错误答案
    answersOptions=wrongAnswers+[row,]#将正确答案和错误答案连接起来
    random.shuffle(answersOptions)#打乱四个答案的顺序
    print('--------------------------------------------------')
    print('问题'+str(epoch+1)+':  '+sheet[q[0]+str(row)].value)
    print('1.'+sheet[q[1]+str(answersOptions[0])].value)
    print('2.'+sheet[q[1]+str(answersOptions[1])].value)
    print('3.'+sheet[q[1]+str(answersOptions[2])].value)
    print('4.'+sheet[q[1]+str(answersOptions[3])].value)
    while 1:
        try:
            ans = int(input())
            break
        except:
            print('请输入1-4的有效数字~')
            pass
    if(row == answersOptions[ans-1]):
        print('\033[92m答对了！\33[0m')
        sheet['F'+str(row)].value += 1
        wb.save(filename)

    else:
        print('\033[91m答错了!\33[0m')
        sheet['G'+str(row)].value += 1
        wb.save(filename)

    print(sheet['B'+str(row)].value,',',sheet['C'+str(row)].value,',',sheet['D'+str(row)].value,)
    print(sheet['H'+str(row)].value,',',sheet['I'+str(row)].value,',',sheet['J'+str(row)].value,)
