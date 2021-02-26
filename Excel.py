from tkinter.constants import NONE
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.chart import (LineChart, BarChart, PieChart, Reference, Series, reference)
import tkinter.messagebox
import Data

wb = Workbook()

def ExcelGraph():
    A()
    B()
    C()


def A():
    sheetA = wb.create_sheet('전체 출원동향', 0)
    A그래프data = Data.전체출원동향()

    for r in dataframe_to_rows(A그래프data, index=False, header=True):
        sheetA.append(r)

    sheetA.insert_cols(2)
    for row, cellobj in enumerate(list(sheetA.columns)[1]):
        n = '=right(A%d,2)' % (row+1)
        cellobj.value = n

    chartA1 = BarChart()
    dataA1 = Reference(sheetA, min_col = 3, min_row = 1, max_row = 21)
    catsA1 = Reference(sheetA, min_col = 2, min_row = 2, max_row = 21)
    chartA1.add_data(dataA1, titles_from_data = True)
    chartA1.set_categories(catsA1)
    chartA1.y_axis.majorGridlines = None

    chartA2 = LineChart()
    dataA2 = Reference(sheetA, min_col = 4, min_row = 1, max_row = 21)
    chartA2.add_data(dataA2, titles_from_data = True)
    chartA2.y_axis.axId = 2000

    # y축 위치 변경
    chartA2.y_axis.crosses = 'max'
    chartA1 += chartA2
    chartA1.width = 20
    chartA1.height = 10
    chartA1.legend.position = 't'
    chartA1.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True)) # 테두리 제거
    sheetA.add_chart(chartA1, 'F2')

    global savepath
    savepath = Data.Save()

def B():
    sheetB = wb.create_sheet('주요국가별 출원동향', 1)
    B그래프data = Data.주요국출원동향()

    for r in dataframe_to_rows(B그래프data, index=False, header=True):
        sheetB.append(r)

    sheetB['A22'] = '합계'
    sheetB['B22'] = '=SUM(B2:B21)'
    sheetB['C22'] = '=SUM(C2:C21)'
    sheetB['D22'] = '=SUM(D2:D21)'
    sheetB['E22'] = '=SUM(E2:E21)'

    chartB1 = LineChart()
    dataB1 = Reference(sheetB, min_col = 2, min_row = 1, max_row = 21, max_col = 5)
    catsB1 = Reference(sheetB, min_col = 1, min_row = 2, max_row = 21)
    chartB1.add_data(dataB1, titles_from_data = True)
    chartB1.set_categories(catsB1)
    chartB1.y_axis.majorGridlines = None
    chartB1.width = 25
    chartB1.height = 10
    chartB1.legend.position = 't'
    chartB1.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    sheetB.add_chart(chartB1, 'G2')

    chartB2 = PieChart()
    dataB2 = Reference(sheetB, min_col = 1, min_row = 22, max_col = 5)
    labelsB2 = Reference(sheetB, min_col = 2, min_row = 1, max_col = 5)
    chartB2.add_data(dataB2, from_rows = 22, titles_from_data = False)
    chartB2.set_categories(labelsB2)
    chartB2.width = 5
    chartB2.height =5
    chartB2.legend = None
    chartB2.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))

    sheetB.add_chart(chartB2, 'G20')

def C():
    sheetC = wb.create_sheet('주요국 내 상위다출원국가', 2)
    C상위다출원국가KR = Data.상위다출원국가('KR')
    C상위다출원국가JP = Data.상위다출원국가('JP')
    C상위다출원국가US = Data.상위다출원국가('US')
    C상위다출원국가EP = Data.상위다출원국가('EP')

    for r in dataframe_to_rows(C상위다출원국가KR, index=False, header=True):
        sheetC.append(r)
    
    sheetC['A22'] = '합계'
    sheetC['B22'] = '=SUM(B2:B21)'
    sheetC['C22'] = '=SUM(C2:C21)'
    sheetC['D22'] = '=SUM(D2:D21)'
    sheetC['E22'] = '=SUM(E2:E21)'
    sheetC['A23'] = ' '

    chartC1 = LineChart()
    dataC1 = Reference(sheetC, min_col = 2, min_row = 1, max_col = 5, max_row = 21)
    catsC1 = Reference(sheetC, min_col = 1, min_row = 2, max_row = 21)
    chartC1.add_data(dataC1, titles_from_data = True)
    chartC1.set_categories(catsC1)
    chartC1.y_axis.majorGridlines = None
    chartC1.width = 25
    chartC1.height = 10
    chartC1.legend.position = 't'
    chartC1.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    sheetC.add_chart(chartC1, 'G2')

    chartC2 = PieChart()
    dataC2 = Reference(sheetC, min_col = 1, min_row = 22, max_col = 5)
    labelC2 = Reference(sheetC, min_col = 2, min_row = 1, max_col = 5)
    chartC2.add_data(dataC2, from_rows = 22, titles_from_data = False)
    chartC2.set_categories(labelC2)
    chartC2.width = 5
    chartC2.height =5
    chartC2.legend.position = 't'
    chartC2.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    sheetC.add_chart(chartC2, 'U2')


    for r in dataframe_to_rows(C상위다출원국가JP, index=False, header=True):
        sheetC.append(r)

    sheetC['A45'] = '합계'
    sheetC['B45'] = '=SUM(B25:B44)'
    sheetC['C45'] = '=SUM(C25:C44)'
    sheetC['D45'] = '=SUM(D25:D44)'
    sheetC['E45'] = '=SUM(E25:E44)'
    sheetC['A46'] = ' '

    chart6 = LineChart()
    data6 = Reference(sheetC, min_col = 2, min_row = 24, max_col = 5, max_row = 44)
    cats6 = Reference(sheetC, min_col = 1, min_row = 25, max_row = 44)
    chart6.add_data(data6, titles_from_data = True)
    chart6.set_categories(cats6)
    chart6.y_axis.majorGridlines = None
    chart6.width = 25
    chart6.height = 10
    chart6.legend.position = 't'
    chart6.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    sheetC.add_chart(chart6, 'G25')

    for r in dataframe_to_rows(C상위다출원국가US, index=False, header=True):
        sheetC.append(r)

    sheetC['A68'] = '합계'
    sheetC['B68'] = '=SUM(B48:B67)'
    sheetC['C68'] = '=SUM(C48:C67)'
    sheetC['D68'] = '=SUM(D48:D67)'
    sheetC['E68'] = '=SUM(E48:E67)'
    sheetC['A69'] = ' '

    chart7 = LineChart()
    data7 = Reference(sheetC, min_col = 2, min_row = 47, max_col = 5, max_row = 67)
    cats7 = Reference(sheetC, min_col = 1, min_row = 48, max_row = 67)
    chart7.add_data(data7, titles_from_data = True)
    chart7.set_categories(cats7)
    chart7.y_axis.majorGridlines = None
    chart7.width = 25
    chart7.height = 10
    chart7.legend.position = 't'
    chart7.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    sheetC.add_chart(chart7, 'G48')

    for r in dataframe_to_rows(C상위다출원국가EP, index=False, header=True):
        sheetC.append(r)

    sheetC['A91'] = '합계'
    sheetC['B91'] = '=SUM(B71:B90)'
    sheetC['C91'] = '=SUM(C71:C90)'
    sheetC['D91'] = '=SUM(D71:D90)'
    sheetC['E91'] = '=SUM(E71:E90)'
    sheetC['A92'] = ' '

    chart8 = LineChart()
    data8 = Reference(sheetC, min_col = 2, min_row = 70, max_col = 5, max_row = 90)
    cats8 = Reference(sheetC, min_col = 1, min_row = 71, max_row = 90)
    chart8.add_data(data8, titles_from_data = True)
    chart8.set_categories(cats8)
    chart8.y_axis.majorGridlines = None
    chart8.width = 25
    chart8.height = 10
    chart8.legend.position = 't'
    chart8.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    sheetC.add_chart(chart8, 'G71')

    wb.save('{}.xlsx'.format(savepath))
    tkinter.messagebox.showinfo('messagebox', '그래프 생성 완료')

    
