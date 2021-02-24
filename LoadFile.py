from numpy.lib.npyio import load
import pandas as pd
from tkinter import filedialog

def Load():
    # 메세지 창 추가 '전처리가 완료된 엑셀 파일을 불러오시오.'
    loadpath = filedialog.askopenfilename(initialdir="/", title="엑셀 파일 선택",
                                          filetypes=(("Excel files","*.xlsx"),
                                          ("all files", "*.*")))
    return loadpath