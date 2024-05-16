import random

import openpyxl

filename = r"/Users/hyeon/Downloads/test1.xlsx"   
wb = openpyxl.load_workbook(filename)
ws = wb.active

excel_to_list_all=[]

for column in ws.columns:
    excel_to_list1=[]

    for cell in column:
        excel_to_list1.append(cell.value)

    excel_to_list_all.append(excel_to_list1)

def divide_list(input_list, num_groups):
    if num_groups <= 0:
        raise ValueError("0이상의 숫자를 입력하셔야 합니다.")
    if num_groups > len(input_list):
        raise ValueError("그룹의 숫자는 리스트의 숫자보다 클수 없습니다.")

    random.shuffle(input_list)

    groups = [[] for _ in range(num_groups)]

    for i, item in enumerate(input_list):
        groups[i % num_groups].append(item)

    return groups

input_list = excel_to_list_all[0]
num_groups = int(input("몇개의 그룹으로 나누시겠습니까?"))

result = divide_list(input_list, num_groups)
print(result)
