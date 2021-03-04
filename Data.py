from tkinter.constants import TRUE
from numpy.lib.npyio import load
import tkinter.messagebox
import pandas as pd
from datetime import datetime
from tkinter import filedialog
import sys, numpy as np

# 기술분류 Radiobutton
def Rb1():
    global rb
    rb = 1

def Rb2():
    global rb
    rb =2

def Rb3():
    global rb
    rb =3

def Rb4():
    global rb
    rb =4

# 출원연도 20년
출원연도list = []
for year in range(20):
    years = year + (datetime.today().year - 19)
    출원연도list.append(years)
출원연도list = pd.DataFrame(data={'출원연도' : 출원연도list})

# IP주요국
주요국 = ['KR', 'JP', 'US', 'EP']

# EP국가
EP국가 = ['GR', 'NL', 'DK', 'DE', 'LV', 'RO', 'LU', 'LT', 'BE', 'BG', 'CY', 'SE', 'ES', 'SK', 'SI', 'IE', 'EE', 'GB', 'AT', 'IT', 'CZ', 'PT', 'PL', 'FR', 'FI', 'HU']

# 동적 변수 할당 mod
mod = sys.modules[__name__]

#
def Load():
    tkinter.messagebox.showinfo('messagebox', '전처리가 완료된 엑셀 파일을 불러오시오.')
    loadpath = filedialog.askopenfilename(initialdir="/", title="엑셀 파일 선택",
                                          filetypes=(("Excel files","*.xlsx"),
                                          ("all files", "*.*")))
    return loadpath

#
def Save():
    tkinter.messagebox.showinfo('messagebox', '저장할 위치를 선택하시오.')
    savepath = filedialog.asksaveasfilename(initialdir="/", title="저장 위치 선택",
                                         filetypes=(("Excel files","*.xlsx"),
                                         ("all files", "*.*")))
    return savepath

#
def 전체출원동향():
    global loadpath
    global Rawdata
    loadpath = Load()
    Rawdata = pd.read_excel(loadpath)

    #데이터 정리
    #출원연도list - 출원연도 값 비교하여 통합
    A출원연도counts = pd.DataFrame(Rawdata['출원연도'].value_counts())
    A그래프data = A출원연도counts.reset_index()
    A그래프data.columns = ['출원연도', '출원건수']
    A그래프data = A그래프data.sort_values(by='출원연도', ascending = True)
    A그래프data = pd.merge(출원연도list, A그래프data, on='출원연도', how='left')
    A그래프data = A그래프data.replace(np.nan, 0, regex=True)
    A그래프data['누적건수'] = np.cumsum(A그래프data['출원건수'])

    return A그래프data

#
def 주요국출원동향():

    #데이터정리(한국, 일본, 미국, 유럽만 해당)
    B국가연도data = Rawdata[['출원국가코드', '출원연도']]
    for country in 주요국:
        Bconditoin = (B국가연도data['출원국가코드'] == country)
        B주요국data = B국가연도data[Bconditoin]
        B출원연도counts = B주요국data['출원연도'].value_counts()
        B주요국data2 = B출원연도counts.reset_index()
        B주요국data2.columns = ['출원연도', '출원건수']
        B주요국data2 = B주요국data2.sort_values(by='출원연도', ascending = True)
        B주요국data21 = pd.merge(출원연도list, B주요국data2, on='출원연도', how='left')
        B주요국data22 = B주요국data21.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr1{}'.format(country), B주요국data22)

    B주요국KR = getattr(mod, 'setattr1{}'.format('KR'))
    B주요국JP = getattr(mod, 'setattr1{}'.format('JP'))
    B주요국US = getattr(mod, 'setattr1{}'.format('US'))
    B주요국EP = getattr(mod, 'setattr1{}'.format('EP'))

    B주요국merge1 = pd.merge(B주요국KR, B주요국JP, on = '출원연도', how = 'left')
    B주요국merge2 = pd.merge(B주요국merge1, B주요국US, on = '출원연도', how = 'left')
    B그래프data = pd.merge(B주요국merge2, B주요국EP, on = '출원연도', how = 'left')

    B그래프data.columns = ['출원연도', 'KR', 'JP', 'US', 'EP']
    
    return B그래프data

#
def 상위다출원국가(code):

    #데이터정리
    Ccondition1 = (Rawdata['출원국가코드'] == code)
    C출원국가data = Rawdata[Ccondition1]
    C출원국가코드counts = C출원국가data['출원인국가코드'].value_counts()
    C상위4개국 = C출원국가코드counts.reset_index()
    C상위4개국.columns = ['출원인국가코드', '출원건수']
    # 상위 4개국
    C상위4개국 = C상위4개국.sort_values(by='출원건수', ascending = False).head(4)

    C국가코드list = list(C상위4개국['출원인국가코드'][0:])
    for country2 in C국가코드list:
        Ccondition2 = (C출원국가data['출원인국가코드'] == country2)
        C출원인국가data = C출원국가data[Ccondition2]
        C출원인국가data = C출원인국가data[['출원연도','출원인국가코드']]
        C출원연도counts = C출원인국가data['출원연도'].value_counts()
        C그래프data2 = C출원연도counts.reset_index()
        C그래프data2.columns = ['출원연도','출원건수']
        C그래프data21 = pd.merge(출원연도list, C그래프data2, on='출원연도', how='left')
        C그래프data22 = C그래프data21.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr2{}'.format(country2), C그래프data22)

    C그래프1 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][0]))
    C그래프2 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][1]))
    C그래프3 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][2]))
    C그래프4 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][3]))

    C그래프merge1 = pd.merge(C그래프1, C그래프2, on = '출원연도', how = 'left')
    C그래프merge2 = pd.merge(C그래프merge1, C그래프3, on = '출원연도', how = 'left')
    C상위다출원국가 = pd.merge(C그래프merge2, C그래프4, on = '출원연도', how = 'left')

    C상위다출원국가.columns = ['출원연도', '{}'.format(C상위4개국['출원인국가코드'][0]), '{}'.format(C상위4개국['출원인국가코드'][1]), '{}'.format(C상위4개국['출원인국가코드'][2]), '{}'.format(C상위4개국['출원인국가코드'][3])]
    C그래프data = C상위다출원국가.sort_values(by='출원연도', ascending = True)
    return C그래프data

def 기술분류():
    D기술분류counts = Rawdata['기술분류'].value_counts()
    D기술분류data = D기술분류counts.reset_index()
    D기술분류data.columns = ['기술분류', '출원건수']

    for classification in D기술분류data['기술분류']:
        Dcondition = (Rawdata['기술분류'] == classification)
        D분류data = Rawdata[Dcondition]
        D출원연도counts = D분류data['출원연도'].value_counts()
        D출원연도data = D출원연도counts.reset_index()
        D출원연도data.columns = ['출원연도', '출원건수']
        D그래프data2 = pd.merge(출원연도list, D출원연도data, on='출원연도', how='left')
        D그래프data2 = D그래프data2.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr3{}'.format(classification), D그래프data2)

    if rb == 2:
        D그래프1 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][0]))
        D그래프2 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][1]))
        D그래프data = pd.merge(D그래프1, D그래프2, on = '출원연도', how = 'left')
        D그래프data.columns = ['출원연도', '{}'.format(D기술분류data['기술분류'][0]), '{}'.format(D기술분류data['기술분류'][1])]
        return D그래프data

    elif rb == 3:
        D그래프1 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][0]))
        D그래프2 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][1]))
        D그래프3 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][2])) 
        D그래프merge1 = pd.merge(D그래프1, D그래프2, on = '출원연도', how = 'left')
        D그래프merge2 = pd.merge(D그래프merge1, D그래프3, on = '출원연도', how = 'left')
        D그래프merge2.columns = ['출원연도', '{}'.format(D기술분류data['기술분류'][0]), '{}'.format(D기술분류data['기술분류'][1]), '{}'.format(D기술분류data['기술분류'][2])]
        D그래프data = D그래프merge2.sort_values(by='출원연도', ascending = True)
        return D그래프data

    elif rb == 4:
        D그래프1 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][0]))
        D그래프2 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][1]))
        D그래프3 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][2]))
        D그래프4 = getattr(mod, 'setattr3{}'.format(D기술분류data['기술분류'][3]))
        D그래프merge1 = pd.merge(D그래프1, D그래프2, on = '출원연도', how = 'left')
        D그래프merge2 = pd.merge(D그래프merge1, D그래프3, on = '출원연도', how = 'left')
        D그래프data = pd.merge(D그래프merge2, D그래프4, on = '출원연도', how = 'left')
        D그래프data.columns = ['출원연도', '{}'.format(D기술분류data['기술분류'][0]), '{}'.format(D기술분류data['기술분류'][1]), '{}'.format(D기술분류data['기술분류'][2]), '{}'.format(D기술분류data['기술분류'][3])]
        D그래프data = D그래프data.sort_values(by='출원연도', ascending = True)
        return D그래프data

def 내외국인점유율():
    ERawdata = Rawdata
    for EP수정 in EP국가:
        ERawdata['출원인국가코드'] = np.where(ERawdata['출원인국가코드'] == EP수정, 'EP', Rawdata['출원인국가코드'])

    for country in 주요국:
        Econdition = (ERawdata['출원국가코드'] == country)
        E주요국data = ERawdata[Econdition]
        E주요국data['출원인국가코드'] = np.where(E주요국data['출원인국가코드'] == country, '내국인', '외국인')
        # A value is trying to be set on a copy of a slice from a DataFrame.
        # Try using .loc[row_indexer,col_indexer] = value instead
        # See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
        E내외국인counts = E주요국data['출원인국가코드'].value_counts()
        E내외국인counts = E내외국인counts.reset_index()
        E내외국인counts.columns = ['분류', '{}'.format(country)]
        setattr(mod, 'setattr4{}'.format(country), E내외국인counts)
    
    E그래프1 = getattr(mod, 'setattr4{}'.format(주요국[0]))
    E그래프2 = getattr(mod, 'setattr4{}'.format(주요국[1]))
    E그래프3 = getattr(mod, 'setattr4{}'.format(주요국[2]))
    E그래프4 = getattr(mod, 'setattr4{}'.format(주요국[3]))

    E그래프merge1 = pd.merge(E그래프1, E그래프2, on = '분류', how = 'left')
    E그래프merge2 = pd.merge(E그래프merge1, E그래프3, on = '분류', how = 'left')
    E그래프data1= pd.merge(E그래프merge2, E그래프4, on = '분류', how = 'left')
    return E그래프data1
    
def 외국인점유율():
    ERawdata = Rawdata
    for EP수정 in EP국가:
        ERawdata['출원인국가코드'] = np.where(ERawdata['출원인국가코드'] == EP수정, 'EP', Rawdata['출원인국가코드'])

    for country in 주요국:
        Econdition = (ERawdata['출원국가코드'] == country)
        E주요국data = ERawdata[Econdition]
        E내외국인counts = E주요국data['출원인국가코드'].value_counts()
        E내외국인counts = E내외국인counts.reset_index()
        E내외국인counts.columns = ['{}'.format(country), '출원건수']
        Econdition2 = (E내외국인counts['{}'.format(country)] != country)
        E외국인data = E내외국인counts[Econdition2]
        setattr(mod, 'setattr5{}'.format(country), E외국인data)

    E그래프1 = getattr(mod, 'setattr5{}'.format(주요국[0]))
    E그래프2 = getattr(mod, 'setattr5{}'.format(주요국[1]))
    E그래프3 = getattr(mod, 'setattr5{}'.format(주요국[2]))
    E그래프4 = getattr(mod, 'setattr5{}'.format(주요국[3]))

    E그래프merge1 = pd.merge(E그래프1, E그래프2, how = 'inner', left_index = True, right_index = True)
    E그래프merge2 = pd.merge(E그래프merge1, E그래프3, how = 'inner', left_index = True, right_index = True)
    E그래프data2= pd.merge(E그래프merge2, E그래프4, how = 'inner', left_index = True, right_index = True)
    E그래프data2.columns = ['KR', '출원건수', 'JP', '출원건수', 'US', '출원건수', 'EP', '출원건수']
    return E그래프data2

