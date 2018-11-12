import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font, NamedStyle
import os

file_location = "//Gcl-file-sr/CCF/03 Planta de Fraccionamiento/Operativo/Producción"
os.chdir(file_location)
reveal = os.getcwd()
workbook = openpyxl.load_workbook("Producción 2018.xlsx")
sheet_prov = int(input("Ingrese mes de producción: "))
sheet_month = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
sheet_1 = sheet_month[sheet_prov - 1]
sheet = workbook[sheet_1]


# Procesamiento de Datos

prod_day = int(input("Ingrese el dia de producción: "))
row = sheet.max_row + 1
max_prod = input("Ingrese número de lotes: ")
cant_prod = input("Ingrese cantidad producida: ")
max_range = row + int(max_prod)


# ----------------------------GENERANDO ESTILOS DE CELDAS----------------------------------------
# Formato de Celdas

format_style = NamedStyle(name="format_style")
thick = Side(border_style="medium", color="000000")
thin = Side(border_style="thin", color="000000")
format_style.font = Font(size=11)
format_style.border = Border(top=thin, left=thin, right=thin, bottom=thin)
format_style.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
# workbook.add_named_style(format_style)

# Estilo del formato inicial para la columna

for j in range(row, max_range):
    for r in range(2, 15):
        sheet.cell(j, r).style = "format_style"
        if j == row:
            sheet.cell(row, r).border = Border(top=thick, left=thick, right=thin, bottom=thin)
        elif j == max_range - 1:
            sheet.cell(j, r).border = Border(top=thin, left=thick, right=thin, bottom=thick)

# Primera Columna
        if r == 2:
            if j == row:
                sheet.cell(row, r).border = Border(top=thick, left=thick, right=thin, bottom=thin)
            elif j == max_range - 1:
                sheet.cell(j, r).border = Border(top=thin, left=thick, right=thin, bottom=thick)
            else:
                sheet.cell(j, r).border = Border(top=thin, left=thick, right=thin, bottom=thin)
# Segunda Columna
        if r == 3:
            if j == row:
                sheet.cell(row, r).border = Border(top=thick, left=thin, right=thick, bottom=thin)
            elif j == max_range - 1:
                sheet.cell(j, r).border = Border(top=thin, left=thin, right=thick, bottom=thick)
            else:
                sheet.cell(j, r).border = Border(top=thin, left=thin, right=thick, bottom=thin)

# Columna Resultados

        for r in range(4, 15):
            if j == row:
                sheet.cell(row, r).border = Border(top=thick, left=thin, right=thin, bottom=thin)
            elif j == max_range - 1:
                sheet.cell(j, r).border = Border(top=thin, left=thin, right=thin, bottom=thick)
            if r == 11:
                if j == row:
                    sheet.cell(row, r).border = Border(top=thick, left=thick, right=thin, bottom=thin)
                elif j == max_range - 1:
                    sheet.cell(j, r).border = Border(top=thin, left=thick, right=thin, bottom=thick)
                else:
                    sheet.cell(j, r).border = Border(top=thin, left=thick, right=thin, bottom=thin)
            if r == 14:
                if j == row:
                    sheet.cell(row, r).border = Border(top=thick, left=thick, right=thick, bottom=thin)
                elif j == max_range - 1:
                    sheet.cell(j, r).border = Border(top=thin, left=thick, right=thick, bottom=thick)
                else:
                    sheet.cell(j, r).border = Border(top=thin, left=thick, right=thick, bottom=thin)


# --------------------------- DATOS CASILLAS EXCEL ----------------------------------------------
# Lote de Procesamiento
    if prod_day < 10:
        sheet.cell(j, 2).value = 'L' + '0' + str(prod_day) + str(sheet_prov) + '18'
    else:
        sheet.cell(j, 2).value = 'L' + str(prod_day) + str(sheet_prov) + '18'
    sheet.cell(j, 3).value = '0' + str(j - (row - 1))

# Carga alimentada + cantidad no alimentada

    sheet.cell(j, 4).value = int(cant_prod)
    sheet.cell(j, 5).value = 0.03

# Valores Fracción Gruesa

    sheet.cell(j, 6).value = float(input("Ingrese producción fracción gruesa {}: ".format(j - (row - 1))))

# Valor Resultados

for co in range(7, 15):
    if co != 8:
        if co == 7:
            sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
            sheet.cell(row, co).value = float(input("Ingrese producción fracción fina: "))
        if co == 9:
            sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
            sheet.cell(row, co).value = '=SUM($F${0}:$G${1})'.format(row, (max_range - 1))
        if co == 10:
            sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
            sheet.cell(row, co).value = '=($D${0}*$C${1})-$I${0}'.format(row, (max_range - 1))
        if co == 11:
            sheet.merge_cells(start_row=row, start_column=co, end_row=(max_range - 1), end_column=13)
            sheet.cell(row, co).value = input("Ingrese observaciones: ")
        if co == 14:
            sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
            sheet.cell(row, co).value = '=SUM($F${0}:$F${1})'.format(row, (max_range - 1))
    else:
        continue


workbook.save('Producción 2018.xlsx')
