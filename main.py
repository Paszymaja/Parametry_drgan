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


def macierz_a(macierz, macierz_stanowisk):
    tab = []
    value = lambda a, b: np.sqrt(pow(macierz[i][3] - macierz_stanowisk[a][b], 2) +
                                 pow(macierz[i][4] - macierz_stanowisk[a][b+1], 2) +
                                 pow(500, 2))
    for i in range(60):
        for j in range(4):
            if j == 0:
                tab.append(np.log10(macierz[i][2]))
            elif j == 1:
                if macierz[i][0] == 1:
                    tab.append(np.log10(value(0, 1)))
                elif macierz[i][0] == 2:
                    tab.append(np.log10(value(1, 1)))
                elif macierz[i][0] == 3:
                    tab.append(np.log10(value(2, 1)))
            elif j == 2:
                if macierz[i][0] == 1:
                    tab.append(value(0, 1))
                elif macierz[i][0] == 2:
                    tab.append(value(1, 1))
                elif macierz[i][0] == 3:
                    tab.append(value(2, 1))
            elif j == 3:
                tab.append(1)
    mat = np.array(tab).reshape(60, 4)
    return mat


def macierz_y(macierz):
    tab = []
    for i in range(60):
        for j in range(1):
            tab.append(np.log10(macierz[i][5]))
    mat = np.array(tab).reshape(60, 1)
    return mat


macierz_danych = wczytanie_danych(arkusz)
stanowiska = wczytanie_stanowisk(arkusz)

macierzA = macierz_a(macierz_danych, stanowiska)
macierzY = macierz_y(macierz_danych)

macierzT = macierzA.transpose()
macierzA = np.dot(macierzT, macierzA)
macierzY = np.dot(macierzT, macierzY)

wynik = np.linalg.solve(macierzA, macierzY)
print(wynik)









