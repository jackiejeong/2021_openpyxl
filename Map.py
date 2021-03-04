import json, folium
import pandas as pd
import tkinter.messagebox
from tkinter import filedialog


def worldmap():
    # 전처리된 엑셀 파일
    tkinter.messagebox.showinfo('messagebox', '전처리된 엑셀 파일을 선택하시오.')
    filepath = filedialog.askopenfilename(initialdir="/", title="엑셀 파일 선택",
                                          filetypes=(("Excel files","*.xlsx"),
                                          ("all files", "*.*")))

    mapdata = pd.read_excel(filepath)
    출원인국가코드counts = mapdata['출원인국가코드'].value_counts()
    출원인국가코드counts = 출원인국가코드counts.reset_index()
    출원인국가코드counts.columns = ['출원인국가코드', '출원건수']

    # 국가코드 파일            
    tkinter.messagebox.showinfo('messagebox', '국가코드 파일을 선택하시오.')
    codepath = filedialog.askopenfilename(initialdir="/", title="엑셀 파일 선택",
                                          filetypes=(("Excel files","*.xlsx"),
                                          ("all files", "*.*")))

    출원인국가코드data = pd.read_excel(codepath)

    # 출원인국가코드counts + 출원인국가코드data
    worlddata = pd.merge(출원인국가코드counts, 출원인국가코드data, how = 'left', on = '출원인국가코드')

    # JSON 파일
    tkinter.messagebox.showinfo('messagebox', 'jSON 파일을 선택하시오.')
    jsonpath = filedialog.askopenfilename(initialdir="/", title="엑셀 파일 선택",
                                          filetypes=(("json files","*.json"),
                                          ("all files", "*.*")))

    jsonfile = json.load(open(jsonpath,encoding = 'utf-8'))
    map = folium.Map(location = [32.3468904,6.1352215], tiles = 'CartoDB dark_matter', zoom_start = 2)

    # 지도 작성
    folium.Choropleth(
        geo_data = jsonfile,
        data = worlddata,
        columns = ['영문명', '출원건수'],
        fill_color = 'YlGn',
        fill_opacity = 0.7,
        line_opacity = 0.5,
        key_on = 'properties.PLACENAME').add_to(map)
    
    folium.GeoJson(jsonfile, style_function = style_function).add_to(map)

    tkinter.messagebox.showinfo('messagebox', '지도를 저장할 위치를 선택하시오.')
    mapsave = filedialog.asksaveasfilename(initialdir="/", title="저장 위치 선택",
                                         filetypes=(("Html files","*.html"),
                                         ("all files", "*.*")))

    map.save('{}.html'.format(mapsave))

def style_function(feature):
    return {
        'opacity' : 0.7,
        'weight' : 0.5, # 선 굵기
        'color' : 'white',
        'fillOpacity' : 0, # 경계선 내 색상
        'dashArray' : '5,5'
    }


