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
    tkinter.messagebox.showinfo('messagebox', '전처리된 엑셀 파일을 선택하시오.')
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

def 내외국인점유율():
    DRawdata = Rawdata
    for EP수정 in EP국가:
        DRawdata['출원인국가코드'] = np.where(DRawdata['출원인국가코드'] == EP수정, 'EP', Rawdata['출원인국가코드'])

    for country in 주요국:
        Dcondition = (DRawdata['출원국가코드'] == country)
        D주요국data = DRawdata[Dcondition]
        D주요국data['출원인국가코드'] = np.where(D주요국data['출원인국가코드'] == country, '내국인', '외국인')
        # A value is trying to be set on a copy of a slice from a DataFrame.
        # Try using .loc[row_indexer,col_indexer] = value instead
        # See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
        D내외국인counts = D주요국data['출원인국가코드'].value_counts()
        D내외국인counts = D내외국인counts.reset_index()
        D내외국인counts.columns = ['분류', '{}'.format(country)]
        setattr(mod, 'setattr4{}'.format(country), D내외국인counts)
    
    D그래프1 = getattr(mod, 'setattr4{}'.format(주요국[0]))
    D그래프2 = getattr(mod, 'setattr4{}'.format(주요국[1]))
    D그래프3 = getattr(mod, 'setattr4{}'.format(주요국[2]))
    D그래프4 = getattr(mod, 'setattr4{}'.format(주요국[3]))

    D그래프merge1 = pd.merge(D그래프1, D그래프2, on = '분류', how = 'left')
    D그래프merge2 = pd.merge(D그래프merge1, D그래프3, on = '분류', how = 'left')
    D그래프data1= pd.merge(D그래프merge2, D그래프4, on = '분류', how = 'left')
    return D그래프data1
    
def 외국인점유율():
    DRawdata = Rawdata
    for EP수정 in EP국가:
        DRawdata['출원인국가코드'] = np.where(DRawdata['출원인국가코드'] == EP수정, 'EP', Rawdata['출원인국가코드'])

    for country in 주요국:
        Dcondition = (DRawdata['출원국가코드'] == country)
        D주요국data = DRawdata[Dcondition]
        D내외국인counts = D주요국data['출원인국가코드'].value_counts()
        D내외국인counts = D내외국인counts.reset_index()
        D내외국인counts.columns = ['{}'.format(country), '출원건수']
        Dcondition2 = (D내외국인counts['{}'.format(country)] != country)
        D외국인data = D내외국인counts[Dcondition2]
        setattr(mod, 'setattr5{}'.format(country), D외국인data)

    D그래프1 = getattr(mod, 'setattr5{}'.format(주요국[0]))
    D그래프2 = getattr(mod, 'setattr5{}'.format(주요국[1]))
    D그래프3 = getattr(mod, 'setattr5{}'.format(주요국[2]))
    D그래프4 = getattr(mod, 'setattr5{}'.format(주요국[3]))

    D그래프coacat1 = pd.concat([D그래프1, D그래프2], axis = 1)
    D그래프coacat2 = pd.concat([D그래프coacat1, D그래프3], axis = 1)
    D그래프data2 = pd.concat([D그래프coacat2, D그래프4], axis = 1)
    return D그래프data2

# def 기술성장():





# def 기술분류():
#     F기술분류counts = Rawdata['기술분류'].value_counts()
#     F기술분류data = F기술분류counts.reset_index()
#     F기술분류data.columns = ['기술분류', '출원건수']

#     for classification in F기술분류data['기술분류']:
#         Fcondition = (Rawdata['기술분류'] == classification)
#         F분류data = Rawdata[Dcondition]
#         F출원연도counts = F분류data['출원연도'].value_counts()
#         F출원연도data = F출원연도counts.reset_index()
#         F출원연도data.columns = ['출원연도', '출원건수']
#         F그래프data2 = pd.merge(출원연도list, D출원연도data, on='출원연도', how='left')
#         F그래프data2 = F그래프data2.replace(np.nan, 0, regex=True)
#         setattr(mod, 'setattr3{}'.format(classification), F그래프data2)

#     if rb == 2:
#         F그래프1 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][0]))
#         F그래프2 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][1]))
#         F그래프data = pd.merge(F그래프1, F그래프2, on = '출원연도', how = 'left')
#         F그래프data.columns = ['출원연도', '{}'.format(F기술분류data['기술분류'][0]), '{}'.format(F기술분류data['기술분류'][1])]
#         return F그래프data

#     elif rb == 3:
#         F그래프1 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][0]))
#         F그래프2 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][1]))
#         F그래프3 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][2])) 
#         F그래프merge1 = pd.merge(F그래프1, F그래프2, on = '출원연도', how = 'left')
#         F그래프merge2 = pd.merge(F그래프merge1, F그래프3, on = '출원연도', how = 'left')
#         F그래프merge2.columns = ['출원연도', '{}'.format(F기술분류data['기술분류'][0]), '{}'.format(F기술분류data['기술분류'][1]), '{}'.format(F기술분류data['기술분류'][2])]
#         F그래프data = F그래프merge2.sort_values(by='출원연도', ascending = True)
#         return F그래프data

#     elif rb == 4:
#         F그래프1 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][0]))
#         F그래프2 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][1]))
#         F그래프3 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][2]))
#         F그래프4 = getattr(mod, 'setattr3{}'.format(F기술분류data['기술분류'][3]))
#         F그래프merge1 = pd.merge(F그래프1, F그래프2, on = '출원연도', how = 'left')
#         F그래프merge2 = pd.merge(F그래프merge1, F그래프3, on = '출원연도', how = 'left')
#         F그래프data = pd.merge(F그래프merge2, F그래프4, on = '출원연도', how = 'left')
#         F그래프data.columns = ['출원연도', '{}'.format(F기술분류data['기술분류'][0]), '{}'.format(F기술분류data['기술분류'][1]), '{}'.format(F기술분류data['기술분류'][2]), '{}'.format(F기술분류data['기술분류'][3])]
#         F그래프data = F그래프data.sort_values(by='출원연도', ascending = True)
#         return F그래프data


