import tkinter as tk
import Preprocessing
import Excel, Data

if __name__ == '__main__':

    root = tk.Tk()
    root.title('정량분석 프로그램')
    root.geometry('250x200+100+100')
    root.resizable(False,False)
    
    전처리bt = tk.Button(root, text = '데이터 전처리', overrelief = 'solid', width = 10, command = Preprocessing.Prep)
    그래프bt = tk.Button(root, text = '그래프 생성', overrelief = 'solid', width = 10, command = Excel.ExcelGraph)
    지도bt = tk.Button(root, text = '지도 생성', overrelief = 'solid', width = 10)
    초기화bt = tk.Button(root, text = '초기화', overrelief = 'solid', width = 10)
    전처리bt.place(x = 10, y = 20)
    그래프bt.place(x = 10, y = 60)
    지도bt.place(x = 10, y = 100)
    초기화bt.place(x= 10, y = 140)

    radio = tk.IntVar()
    기술분류1radio = tk.Radiobutton(root, text = '기술분류 X', variable = radio, value = 1)
    기술분류2radio = tk.Radiobutton(root, text = '기술분류 2개', variable = radio, value = 2)
    기술분류3radio = tk.Radiobutton(root, text = '기술분류 3개', variable = radio, value = 3)
    기술분류4radio = tk.Radiobutton(root, text = '기술분류 4개', variable = radio, value = 4)
    기술분류1radio.place(x = 150, y = 20)
    기술분류2radio.place(x = 150, y = 60)
    기술분류3radio.place(x = 150, y = 100)
    기술분류4radio.place(x = 150, y = 140)

    label = tk.Label(root, text = '만든 이 : Jackie')
    label.place(x = 160, y = 180)

    root.mainloop()