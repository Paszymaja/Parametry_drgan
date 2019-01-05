import openpyxl
import numpy as np

wb = openpyxl.load_workbook('dane.xlsx')
arkusz = wb['Arkusz1']


def wczytanie_danych(sheet):
    tab = []
    for row_i in range(2, 62):
        for column_i in range(1, 8):
            x1 = float(sheet.cell(row_i, column_i).value)
            tab.append(x1)
    mat = np.array(tab).reshape(60, 7)
    return mat


dane_w_macierzy = wczytanie_danych(arkusz)





