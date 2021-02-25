import tkinter as tk
import Preprocessing
import Excel

if __name__ == '__main__':
    
    root = tk.Tk()
    root.title('정량분석')
    root.geometry('300x300+100+100')
    root.resizable(False,False)
    
    전처리bt = tk.Button(root, text = '데이터 전처리', overrelief = 'solid', width = 10, command = Preprocessing.Prep)
    전처리bt.pack()
    그래프bt = tk.Button(root, text = '그래프 생성', overrelief = 'solid', width = 10, command = Excel.ExcelGraph)
    그래프bt.pack()
    지도bt = tk.Button(root, text = '지도 생성', overrelief = 'solid', width = 10)
    지도bt.pack()
    전처리bt.place(x = 100, y = 20)
    그래프bt.place(x = 100, y = 60)
    지도bt.place(x = 100, y = 100)

    root.mainloop()