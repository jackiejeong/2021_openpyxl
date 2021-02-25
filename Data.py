from numpy.lib.npyio import load
import tkinter.messagebox
import pandas as pd
from datetime import datetime
from tkinter import filedialog
import sys, numpy as np

# 출원연도 20년
출원연도list = []
for year in range(20):
    years = year + (datetime.today().year - 19)
    출원연도list.append(years)
출원연도list = pd.DataFrame(data={'출원연도' : 출원연도list})

# IP주요국
주요국 = ['KR', 'JP', 'US', 'EP']

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
    Rawdata = pd.read_excel(loadpath)

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
def KR상위다출원국가():
    Rawdata = pd.read_excel(loadpath)

    #데이터정리
    Ccondition1 = (Rawdata['출원국가코드'] == 'KR')
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
        C그래프data2 = C그래프data2.sort_values(by='출원연도', ascending = True)
        C그래프data21 = pd.merge(출원연도list, C그래프data2, on='출원연도', how='left')
        C그래프data22 = C그래프data21.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr2{}'.format(country2), C그래프data22)

    C그래프1 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][0]))
    C그래프2 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][1]))
    C그래프3 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][2]))
    C그래프4 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][3]))

    C그래프merge1 = pd.merge(C그래프1, C그래프2, on = '출원연도', how = 'left')
    C그래프merge2 = pd.merge(C그래프merge1, C그래프3, on = '출원연도', how = 'left')
    C상위다출원국가KR = pd.merge(C그래프merge2, C그래프4, on = '출원연도', how = 'left')

    C상위다출원국가KR.columns = ['출원연도', '{}'.format(C상위4개국['출원인국가코드'][0]), '{}'.format(C상위4개국['출원인국가코드'][1]), '{}'.format(C상위4개국['출원인국가코드'][2]), '{}'.format(C상위4개국['출원인국가코드'][3])]

    return C상위다출원국가KR

#
def JP상위다출원국가():
    Rawdata = pd.read_excel(loadpath)

    #데이터정리
    Ccondition1 = (Rawdata['출원국가코드'] == 'JP')
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
        C그래프data2 = C그래프data2.sort_values(by='출원연도', ascending = True)
        C그래프data21 = pd.merge(출원연도list, C그래프data2, on='출원연도', how='left')
        C그래프data22 = C그래프data21.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr2{}'.format(country2), C그래프data22)

    C그래프1 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][0]))
    C그래프2 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][1]))
    C그래프3 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][2]))
    C그래프4 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][3]))

    C그래프merge1 = pd.merge(C그래프1, C그래프2, on = '출원연도', how = 'left')
    C그래프merge2 = pd.merge(C그래프merge1, C그래프3, on = '출원연도', how = 'left')
    C상위다출원국가JP = pd.merge(C그래프merge2, C그래프4, on = '출원연도', how = 'left')

    C상위다출원국가JP.columns = ['출원연도', '{}'.format(C상위4개국['출원인국가코드'][0]), '{}'.format(C상위4개국['출원인국가코드'][1]), '{}'.format(C상위4개국['출원인국가코드'][2]), '{}'.format(C상위4개국['출원인국가코드'][3])]

    return C상위다출원국가JP

#
def US상위다출원국가():
    Rawdata = pd.read_excel(loadpath)

    #데이터정리
    Ccondition1 = (Rawdata['출원국가코드'] == 'US')
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
        C그래프data2 = C그래프data2.sort_values(by='출원연도', ascending = True)
        C그래프data21 = pd.merge(출원연도list, C그래프data2, on='출원연도', how='left')
        C그래프data22 = C그래프data21.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr2{}'.format(country2), C그래프data22)

    C그래프1 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][0]))
    C그래프2 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][1]))
    C그래프3 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][2]))
    C그래프4 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][3]))

    C그래프merge1 = pd.merge(C그래프1, C그래프2, on = '출원연도', how = 'left')
    C그래프merge2 = pd.merge(C그래프merge1, C그래프3, on = '출원연도', how = 'left')
    C상위다출원국가US = pd.merge(C그래프merge2, C그래프4, on = '출원연도', how = 'left')

    C상위다출원국가US.columns = ['출원연도', '{}'.format(C상위4개국['출원인국가코드'][0]), '{}'.format(C상위4개국['출원인국가코드'][1]), '{}'.format(C상위4개국['출원인국가코드'][2]), '{}'.format(C상위4개국['출원인국가코드'][3])]

    return C상위다출원국가US

#
def EP상위다출원국가():
    Rawdata = pd.read_excel(loadpath)

    #데이터정리
    Ccondition1 = (Rawdata['출원국가코드'] == 'EP')
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
        C그래프data2 = C그래프data2.sort_values(by='출원연도', ascending = True)
        C그래프data21 = pd.merge(출원연도list, C그래프data2, on='출원연도', how='left')
        C그래프data22 = C그래프data21.replace(np.nan, 0, regex=True)
        setattr(mod, 'setattr2{}'.format(country2), C그래프data22)

    C그래프1 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][0]))
    C그래프2 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][1]))
    C그래프3 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][2]))
    C그래프4 = getattr(mod, 'setattr2{}'.format(C상위4개국['출원인국가코드'][3]))

    C그래프merge1 = pd.merge(C그래프1, C그래프2, on = '출원연도', how = 'left')
    C그래프merge2 = pd.merge(C그래프merge1, C그래프3, on = '출원연도', how = 'left')
    C상위다출원국가EP = pd.merge(C그래프merge2, C그래프4, on = '출원연도', how = 'left')

    C상위다출원국가EP.columns = ['출원연도', '{}'.format(C상위4개국['출원인국가코드'][0]), '{}'.format(C상위4개국['출원인국가코드'][1]), '{}'.format(C상위4개국['출원인국가코드'][2]), '{}'.format(C상위4개국['출원인국가코드'][3])]

    return C상위다출원국가EP