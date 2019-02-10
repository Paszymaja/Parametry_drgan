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

def wczytanie_stanowisk(sheet):
    tab = []
    for row_i in range(2, 5):
        for column_i in range(9, 12):
            x1 = float(sheet.cell(row_i, column_i).value)
            tab.append(x1)
    mat = np.array(tab).reshape(3, 3)
    return mat


def macierz_A(macierz, macierz_stanowisk):
    tab = []
    print(macierz)
    for i in range(60):
        for j in range(4):
            if j == 0:
                tab.append(np.log10(macierz[i][2]))
            elif j == 1:
                if macierz[i][0] == 1:
                    tab.append(np.log10(np.sqrt(pow(macierz[i][3] - macierz_stanowisk[0][1], 2) +
                                                pow(macierz[i][4] - macierz_stanowisk[0][2], 2) +
                                                pow(500, 2))))
                elif macierz[i][0] == 2:
                    tab.append(np.log10(np.sqrt(pow(macierz[i][3] - macierz_stanowisk[1][1], 2) +
                                                pow(macierz[i][4] - macierz_stanowisk[1][2], 2) +
                                                pow(500, 2))))
                elif macierz[i][0] == 3:
                    tab.append(np.log10(np.sqrt(pow(macierz[i][3] - macierz_stanowisk[2][1], 2) +
                                                pow(macierz[i][4] - macierz_stanowisk[2][2], 2) +
                                                pow(500, 2))))
            elif j == 2:
                if macierz[i][0] == 1:
                    tab.append(np.sqrt(pow(macierz[i][3] - macierz_stanowisk[0][1], 2)))
    return tab


macierz_danych = wczytanie_danych(arkusz)
darek = wczytanie_stanowisk(arkusz)
print(macierz_A(macierz_danych, darek))








