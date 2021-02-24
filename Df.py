import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.chart import (LineChart, BarChart, Reference, Series)
from datetime import datetime
from tkinter import filedialog
import sys, numpy as np
import LoadFile

wb = Workbook()

## 첫번째(A), 전체 출원동향(연도별 출원건수 및 누적건수)
def 전체출원동향():

    sheet1 = wb.create_sheet('전체출원동향',0)

    loadpath = LoadFile.Load()
    Rawdata = pd.read_excel(loadpath)

    # x축, 출원연도list
    출원연도list = []
    for year in range(20):
        years = year + (datetime.today().year - 19)
        출원연도list.append(years)

    # DataFrame 생성 시 열 만들고 데이터 입력
    출원연도list = pd.DataFrame(data={'출원연도' : 출원연도list})

    #데이터 정리
    A출원연도counts = pd.DataFrame(Rawdata['출원연도'].value_counts())
    A그래프data = A출원연도counts.reset_index()
    A그래프data.columns = ['출원연도', '출원건수']
    # A그래프data.rename(columns={'index' : '출원연도', '출원연도' : '출원건수'}, inplace = True)
    A그래프data = A그래프data.sort_values(by='출원연도', ascending = True)
    A그래프data['누적건수'] = np.cumsum(A그래프data['출원건수'])

    for r in dataframe_to_rows(A그래프data, index=False, header=True):
        sheet1.append(r)


    savepath = filedialog.asksaveasfilename(initialdir="/", title="엑셀 파일 선택",
                                         filetypes=(("Excel files","*.xlsx"),
                                         ("all files", "*.*")))
    wb.save('{}.xlsx'.format(savepath))




