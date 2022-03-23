import xlrd, operator
from xlutils.copy import copy
import os.path

def get_cell_range_values(sheet, row,  col, end):
    data = sheet.col_values(start_rowx=row, colx=col, end_rowx=end)
    return data

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

# Se crea el catalogo
wb = xlrd.open_workbook('catalogo.xls')
sh = wb.sheet_by_index(0)

max = sh.nrows
catalogo = []
cat = []

for i in range(2, max):
    if sh.cell_value(i, 4) != '':
        if sh.cell_value(i, 4) not in cat:
            cat.append(sh.cell_value(i, 4))

for t in cat:
    horas = []
    plaza = []
    for i in range(2, max):
        if sh.cell_value(i, 4) == t:
            if sh.cell_value(i, 2) == 'PLAZA':
                plaza.append(sh.cell_value(i, 0))
            if sh.cell_value(i, 2) == 'HORA':
                horas.append(sh.cell_value(i, 0))

    catalogo.append([t, plaza, horas])

datos = xlrd.open_workbook('distrib.xls')
hoja_datos = datos.sheet_by_index(0)
filas = hoja_datos.nrows-1
columnas = hoja_datos.ncols

plantilla = xlrd.open_workbook('plantilla.xls',formatting_info=True)
hoja_plantilla = plantilla.sheet_by_index(0)
new = copy(plantilla)
hoja_new = new.get_sheet(0) 

col_new = 2
col_ini = 2
for c in catalogo:
    # Plaza
    # print(c[0])
    # print('Plaza')
    listap = c[1]
    # print(listap)
    fil_new = 3
    suma = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for num in range(col_ini, columnas):
        if str(hoja_datos.cell_value(0, num)).upper() in listap:
            tmp = get_cell_range_values(hoja_datos, 1, num, filas)
            suma = list(map(operator.add, suma, tmp))
    # print(suma)
    for i in suma:
        hoja_new.write(fil_new, col_new, i)
        fil_new += 1
    col_new += 1
    # Horas
    # print('Horas)
    listah = c[2]
    # print(listah)
    fil_new = 3
    suma = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for num in range(col_ini, columnas):
        if str(hoja_datos.cell_value(0, num)).upper() in listah:
            tmp = get_cell_range_values(hoja_datos, 1, num, filas)
            suma = list(map(operator.add, suma, tmp))
    # print(suma)
    for i in suma:
        hoja_new.write(fil_new, col_new, i)
        fil_new += 1
    col_new += 1


# hoja_new.write(3, 2, 'asd')
new.save('simple.xls')

print('finalizado')



# Se crea el archivo de salida

