import pandas as pd
import numpy as np
from tkinter import filedialog
import tkinter.messagebox
from datetime import datetime

def Prep():
    tkinter.messagebox.showinfo('message box', '다운로드 받은 엑셀 파일을 불러오시오.')
    loadpath = filedialog.askopenfilename(initialdir="/", title="엑셀 파일 선택",
                                          filetypes=(("Excel files","*.xlsx"),
                                          ("all files", "*.*")))
    data = pd.read_excel(loadpath, header = 4)
    Rawdata = data[['번호', '명칭', '요약', '출원인', '출원인주소', '출원인국가', '출원번호', '출원일',
                        '공개번호', '공개일', '등록번호', '등록일', '권리 현황', '최종 상태', '자국인용횟수', '자국피인용횟수', 'INPADOC패밀리수']]
    
    # 출원일 필터
    condition = (Rawdata['출원일'] >= str(datetime.today().year - 19))
    Rawdata = Rawdata[condition]
    
    # 출원연도 추출
    출원연도list = []
    for date in Rawdata['출원일']:
        year = date[0:4]
        출원연도list.append(year)
        
    # 출원연도 컬럼 추가
    Rawdata.insert(7,'출원연도',출원연도list)
    
    # NaN을 빈 칸으로 바꾸고 데이터 교체
    Rawdata = Rawdata.replace(np.nan, '', regex=True)
    
    # 출원국가 추출
    출원국가codelist = []
    for number in Rawdata['번호']:
        출원국가code = number[0:2]
        출원국가codelist.append(출원국가code)
        
    # 컬럼 추가(순서, 컬럼이름, 값)
    Rawdata.insert(0,'출원국가코드',출원국가codelist)
    Rawdata.insert(0,'기술분류','')
    
    # 출원인국가 추출 / '-'로 입력된 출원인국가코드 확인하여야 함
    # 배경색 바꾸는거 진행중 # 
    출원인국가codelist = []
    for country in Rawdata['출원인국가']:
        if len(country) >= 2:
            countries = country[0:2]
        else:
            countries = '-'
        출원인국가codelist.append(countries)
        
    # 컬럼 추가
    Rawdata.insert(7,'출원인국가코드', 출원인국가codelist)
    Rawdata.insert(6, '대표명화', '')
    Rawdata.insert(6, '출원인정리', '')

    tkinter.messagebox.showinfo('message box', '전처리가 완료된 파일을 저장할 위치를 선택하시오.')
    savepath = filedialog.asksaveasfilename(initialdir="/", title="전처리 데이터 저장 위치 선택",
                                          filetypes=(("Excel files","*.xlsx"),
                                          ("all files", "*.*")))
    tkinter.messagebox.showinfo('message box', '기술분류, 대표명화 작업하시오.')
    
    
    Rawdata.to_excel('{}.xlsx'.format(savepath))

