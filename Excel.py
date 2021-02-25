from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.chart import (LineChart, BarChart, PieChart, Reference, Series, reference)
import Data

wb = Workbook()

def ExcelGraph():
    A()
    B()
    C()


def A():
    sheet1 = wb.create_sheet('전체 출원동향', 0)
    A그래프data = Data.전체출원동향()

    for r in dataframe_to_rows(A그래프data, index=False, header=True):
        sheet1.append(r)

    chart1 = BarChart()
    data1 = Reference(sheet1, min_col = 2, min_row = 1, max_row = 21)
    cats1 = Reference(sheet1, min_col = 1, min_row = 2, max_row = 21)
    chart1.add_data(data1, titles_from_data = True)
    chart1.set_categories(cats1)
    chart1.y_axis.majorGridlines = None

    chart2 = LineChart()
    data2 = Reference(sheet1, min_col = 3, min_row = 1, max_row = 21)
    chart2.add_data(data2, titles_from_data = True)
    chart2.y_axis.axId = 2000

    # y축 위치 변경
    chart2.y_axis.crosses = 'max'
    chart1 += chart2
    sheet1.add_chart(chart1, 'F1')

    global savepath
    savepath = Data.Save()

def B():
    sheet2 = wb.create_sheet('주요국가별 출원동향', 1)
    B그래프data = Data.주요국출원동향()

    for r in dataframe_to_rows(B그래프data, index=False, header=True):
        sheet2.append(r)

    sheet2['A22'] = '합계'
    sheet2['B22'] = '=SUM(B2:B21)'
    sheet2['C22'] = '=SUM(C2:C21)'
    sheet2['D22'] = '=SUM(D2:D21)'
    sheet2['E22'] = '=SUM(E2:E21)'

    chart3 = LineChart()
    data3 = Reference(sheet2, min_col = 2, min_row = 1, max_row = 21, max_col = 5)
    cats3 = Reference(sheet2, min_col = 1, min_row = 2, max_row = 21)
    chart3.add_data(data3, titles_from_data = True)
    chart3.set_categories(cats3)
    chart3.y_axis.majorGridlines = None
    sheet2.add_chart(chart3, 'G1')

    chart4 = PieChart()
    data4 = Reference(sheet2, min_col = 1, min_row = 22, max_col = 5)
    labels4 = Reference(sheet2, min_col = 2, min_row = 1, max_col = 5)
    chart4.add_data(data4, from_rows = 22, titles_from_data = False)
    chart4.set_categories(labels4)
    sheet2.add_chart(chart4, 'G22')

def C():
    sheet3 = wb.create_sheet('주요국 내 상위다출원국가', 2)
    C상위다출원국가KR = Data.KR상위다출원국가()
    C상위다출원국가JP = Data.JP상위다출원국가()
    C상위다출원국가US = Data.US상위다출원국가()
    C상위다출원국가EP = Data.EP상위다출원국가()

    for r in dataframe_to_rows(C상위다출원국가KR, index=False, header=True):
        sheet3.append(r)
    
    sheet3['A22'] = '합계'
    sheet3['B22'] = '=SUM(B2:B21)'
    sheet3['C22'] = '=SUM(C2:C21)'
    sheet3['D22'] = '=SUM(D2:D21)'
    sheet3['E22'] = '=SUM(E2:E21)'
    sheet3['A23'] = ' '

    for r in dataframe_to_rows(C상위다출원국가JP, index=False, header=True):
        sheet3.append(r)

    sheet3['A45'] = '합계'
    sheet3['B45'] = '=SUM(B25:B44)'
    sheet3['C45'] = '=SUM(C25:C44)'
    sheet3['D45'] = '=SUM(D25:D44)'
    sheet3['E45'] = '=SUM(E25:E44)'
    sheet3['A46'] = ' '

    for r in dataframe_to_rows(C상위다출원국가US, index=False, header=True):
        sheet3.append(r)

    sheet3['A68'] = '합계'
    sheet3['B68'] = '=SUM(B48:B67)'
    sheet3['C68'] = '=SUM(C48:C67)'
    sheet3['D68'] = '=SUM(D48:D67)'
    sheet3['E68'] = '=SUM(E48:E67)'
    sheet3['A69'] = ' '

    for r in dataframe_to_rows(C상위다출원국가EP, index=False, header=True):
        sheet3.append(r)

    sheet3['A91'] = '합계'
    sheet3['B91'] = '=SUM(B71:B90)'
    sheet3['C91'] = '=SUM(C71:C90)'
    sheet3['D91'] = '=SUM(D71:D90)'
    sheet3['E91'] = '=SUM(E71:E90)'
    sheet3['A92'] = ' '

    wb.save('{}.xlsx'.format(savepath))

    
