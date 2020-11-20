# -
최근 로또 당첨 번호와 상금을 알려주고, 로또 번호들을 추천해주면서 유명한 로또 판매점을 알려주는 프로그램

print('---Welcome to Lotto world!---')
print()
a = input('안녕하세요! 성함이 어떻게 되시나요?: ')
print('반갑습니다.',a,'님. 이제부터 당신의 미래를 위한 로또를 추첨하겠습니다!')
print('최근 추첨결과는 다음과 같습니다.')

import pandas as pd
from IPython.display import display
data1 = pd.read_excel("수정해본거.xlsx")
display(data1) 

print('당신의 거주지에 위치하는 로또판매점을 추천해드리겠습니다.')
print('당신의 현재 거주지를 알려주세요. (ex. 경기 광명시, 경기 고양시 일산동구, 경기 수원시 영통구, 충남 천안시 동남구, 서울 구로구..) ')
b = str(input())

import pandas as pd
data1 = pd.read_excel("온라인복권자동번호당첨판매점현황_200518.xlsx")
data1 

'''if len(b) < 6:
    new = str(input('잘못 입력하셨습니다. 다시 한번 입력해주세요: '))
    b = new'''

import os
path = "./files"
file_list = os.listdir(path)

from openpyxl import load_workbook

c = False
for file_name_raw in file_list:
    if c == False:
        c = True
        continue
    file_name = "./files/" + file_name_raw
    wb = load_workbook(filename=file_name, data_only=True)
    ws = wb.active
    
    result = []
    
    i = 2 
    while i < 309:
        A = 'A' + str(i)
        B = 'B' + str(i)
        C = 'C' + str(i)
        result.append(ws[A].value)
        result.append(ws[B].value)
        result.append(ws[C].value)
        i += 1
        
if len(b) < 6:
    new = str(input('잘못 입력하셨습니다. 다시 한번 입력해주세요: '))
    b = new        
if b not in result:
    print('당신의 거주지에는 유명한 로또판매점이 없습니다. 아쉽습니다.')
    new = str(input('당신의 거주지에서 가까운 곳을 입력해주세요(ex. 경기 수원시 영통구): '))
    b = new
    

num = result.index(b)
fir = result[num-1]
sed = result[num+1]
fir = str(fir)
sed = str(sed)
print()
print('*'+b+'에 위치하는 유명한 로또판매점은 '+fir+'입니다. 이곳에서 1등(자동)이 총 '+sed+'번 당첨되었습니다.')
print()
print('이 판매점으로 가서 로또를 사시길 추천드립니다.')
print()
print('다음은 로또 번호를 추천해드리겠습니다.')

import random
lotto = set()
while True:
    if len(lotto) == 6:
        break;
    lotto.add(random.randint(1,45))
    
lotto = tuple(lotto)
print(lotto)

print()
print('이번주 로또 1등 당첨은 당신의 것입니다!')
print()
print('---이용해주셔서 감사합니다---')
