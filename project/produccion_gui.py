import tkinter as tk
from tkinter import font, messagebox, Button, ttk
import tkinter.scrolledtext as tsk
import openpyxl
from openpyxl.worksheet.write_only import WriteOnlyCell
from openpyxl.styles import Border, Side, Alignment, Font, NamedStyle
import os
import datetime


class PRD:
    file_location = "//Gcl-file-sr/CCF/03 Planta de Fraccionamiento/Operativo/Producción"
    os.chdir(file_location)

    def produccion(self):
        self.dialogo = tk.Toplevel()
        self.dialogo.geometry('400x300+400+180')
        self.dialogo.resizable(0, 0)
        self.dialogo.title('Producción')
        # Registro Producción Mensual
        customer_prom = tk.LabelFrame(self.dialogo, text='Registros', bg='light steel blue')
        customer_prom.grid(row=0, column=0, padx=10, pady=10)
        etiq_prom = tk.Label(customer_prom, text='Producción Mes Actual')
        etiq_prom.grid(row=0, column=0, stick='e', padx=10, pady=10)
        file_mes = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
        fecha_hoy = datetime.date.today()
        año = fecha_hoy.year
        mes2 = fecha_hoy.month
        file_prod = openpyxl.load_workbook('Producción {0}.xlsx'.format(año), data_only=True)
        sheet3 = file_mes[int(mes2) - 1]
        sheet2 = file_prod[sheet3]
        reg_milavena = sheet2.cell(5, 17).internal_value
        reg_protavena = sheet2.cell(6, 17).internal_value
        print(sheet2, reg_milavena, reg_protavena)
        label_mil = tk.Label(customer_prom, text='Milavena {0:.2f} kg'.format(reg_milavena))
        label_mil.grid(row=1, column=0, stick='w', padx=10)
        label_prot = tk.Label(customer_prom, text='Protavena {0:.2f} kg'.format(reg_protavena))
        label_prot.grid(row=2, column=0, stick='w', padx=10)

        # Registro Producción Mes Anterior
        etiq_prom2 = tk.Label(customer_prom, text='Producción Mes Anterior')
        etiq_prom2.grid(row=0, column=2, stick='e', padx=10, pady=10)
        sheet4 = file_mes[int(mes2) - 2]
        sheet5 = file_prod[sheet4]
        reg_milavena2 = float(sheet5.cell(5, 17).value)
        reg_protavena2 = float(sheet5.cell(6, 17).value)
        print(sheet2, reg_milavena2, reg_protavena2)
        label_mil2 = tk.Label(customer_prom, text='Milavena {0:.2f} kg'.format(reg_milavena2))
        label_mil2.grid(row=1, column=2, stick='w', padx=10)
        label_prot2 = tk.Label(customer_prom, text='Protavena {0:.2f} kg'.format(reg_protavena2))
        label_prot2.grid(row=2, column=2, stick='w', padx=10)

        # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy)
        boton.grid(row=3, column=3)
        botonp = Button(self.dialogo, text='Producción diaria (+)', command=PRD().prod_diaria)
        botonp.grid(row=1, column=0)
        self.dialogo.transient()
        self.dialogo.grab_set()
        self.dialogo.wait_window(self.dialogo)

    def prod_diaria(self):
        dialogo2 = tk.Toplevel()
        dialogo2.geometry('590x400+340+120')
        dialogo2.resizable(0, 0)
        dialogo2.title('Producción diaria')
        # Label producción
        customer_frame = tk.LabelFrame(dialogo2, text='Fecha', bg='light steel blue')
        customer_frame.grid(row=0, column=0, padx=10, pady=20)
        customer_frame2 = tk.LabelFrame(dialogo2, text='Producción', bg='light steel blue')
        customer_frame2.grid(row=0, column=3, padx=10, pady=10)
        customer_frame3 = tk.LabelFrame(dialogo2, text='Observación', bg='light steel blue')
        customer_frame3.grid(row=1, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=5, pady=5)

        # dia de producción
        def check_day(event):
            self.pro_day = day.get()
        etiq2 = tk.Label(customer_frame, text='Dia de Producción:')
        etiq2.grid(row=0, column=0, stick='w', padx=10, pady=10)
        day = ttk.Combobox(customer_frame, values=[i for i in range(1, 32)], state='readonly')
        day.grid(row=1, column=0, padx=10, pady=2)
        day.bind('<<ComboboxSelected>>', check_day)

        # Mes de producción
        def check_prod(event):
            self.values = ('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre')
            self.dia_prod = mes.get()
            self.mes_value = self.values.index(self.dia_prod) + 1
        etiq1 = tk.Label(customer_frame, text='Mes de Producción:')
        etiq1.grid(row=2, column=0, stick='w', padx=10, pady=10)
        mes = ttk.Combobox(customer_frame, values=('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'), state='readonly')
        mes.grid(row=3, column=0, padx=10, pady=2)
        mes.bind("<<ComboboxSelected>>", check_prod)
        # Año de producción
        year_label = tk.Label(customer_frame, text='Año')
        year_label.grid(row=4, column=0, pady=10, padx=10, stick='w')
        self.year_date = tk.Entry(customer_frame, width=5)
        self.year_date.grid(row=4, column=0, stick='e', pady=10, padx=5)

        # Carga Alimentada
        etiqc = tk.Label(customer_frame, text='Carga Alimentada').grid(row=5, column=0, stick='w', padx=10, pady=5)
        self.cantidad_prod = tk.Entry(customer_frame, width=5)
        self.cantidad_prod.grid(row=5, column=0, padx=5, pady=10, stick='e')

        # Número de lotes
        self.max_prod = tk.IntVar()

        def onClick():
            self.max_prod.set(self.max_prod.get() + 1)
            print(self.max_prod.get())
            self.fr_gruesa = []
            for i in range(self.max_prod.get()):
                label_prod = tk.Label(customer_frame2, text='%2d.-' % (i + 1))
                label_prod.grid(row=i + 1, column=2)
                self.c = tk.Entry(customer_frame2, width=10)
                self.c.grid(row=i + 1, column=3)
                self.fr_gruesa.append(self.c)

        def action(event):
            for i in range(self.max_prod.get()):
                self.label_act = tk.Label(customer_frame2, text='✔', bg='light steel blue')
                self.label_act.grid(row=i + 1, column=4)

        def action2(event):
            label_tik = tk.Label(customer_frame2, text='✔', bg='light steel blue')
            label_tik.grid(row=11, column=4)

        def action3(event):
            label_obs = tk.Label(customer_frame3, text='✔', bg='light steel blue')
            label_obs.grid(row=3, column=5)

        def action4(event):
            label_cg = tk.Label(customer_frame, text='✔', bg='light steel blue')
            label_cg.grid(row=5, column=1)

        # Cantidad de produccion + label

        def setText():
            self.entries = []
            print(self.fr_gruesa)
            for self.c in self.fr_gruesa:
                self.entries.append(self.c.get())

        etiq4 = tk.Label(customer_frame2, text='Producción Fracción Gruesa').grid(row=0, column=3, padx=10, pady=10)
        lotes = tk.Button(customer_frame2, text='+', command=onClick, fg='gray24', bg='papaya whip')
        lotes.grid(row=0, column=4)
        boton_max = tk.Button(customer_frame2, text='Set', command=setText, fg='gray24', bg='papaya whip')
        boton_max.grid(row=0, column=5, pady=10)
        boton_max.bind('<Button-1>', action)
        # Cantidad de producción fina
        etiq5 = tk.Label(customer_frame2, text='Producción Fracción Fina').grid(row=10, column=3, padx=10, pady=5)
        self.cant_fina = tk.Entry(customer_frame2, width=10)
        self.cant_fina.grid(row=11, column=3, pady=10)

        # Observaciones
        etiq6 = tk.Label(customer_frame3, text='Observaciones').grid(row=3, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=5, pady=5)
        self.obs = tsk.ScrolledText(master=customer_frame3, wrap=tk.WORD, width=45, height=4)
        self.obs.grid(row=3, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=10, pady=10)
        self.obs.bind('<Return><Return>', action3)

        # Actualizar Base de datos
        boton_act = tk.Button(dialogo2, text='Actualizar BD', command=self._produccion)
        boton_act.grid(row=3, column=5, padx=10)
        # Apertura y guardado de archivo
        self.file_location = "//Gcl-file-sr/CCF/03 Planta de Fraccionamiento/Operativo/Producción"
        os.chdir(self.file_location)
        reveal = os.getcwd()

        def check_file():
            self.year = self.year_date.get()
            if os.path.isfile('Producción {0}.xlsx'.format(self.year)):
                self.workbook = openpyxl.load_workbook("Producción {0}.xlsx".format(self.year))
            else:
                self.workbook = openpyxl.load_workbook('Tipo.xlsx')
                self.format_style = NamedStyle(name="format_style")
                thick = Side(border_style="medium", color="000000")
                thin = Side(border_style="thin", color="000000")
                self.format_style.font = Font(size=11)
                self.format_style.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                self.format_style.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
                self.workbook.add_named_style(self.format_style)
                self.workbook.save('Producción {0}.xlsx'.format(self.year))

        file_boton = tk.Button(customer_frame, text='Revisión', command=check_file)
        file_boton.grid(row=4, column=1, stick='e', padx=5, pady=10)
        # Botón Cerrar y mantener window inicial
        boton = Button(dialogo2, text='Cerrar', command=dialogo2.destroy)
        boton.grid(row=4, column=5, padx=10, pady=10)
        dialogo2.grab_set()
        dialogo2.wait_window(dialogo2)

    def _produccion(self):
        sheet_1 = self.dia_prod
        sheet_2 = self.mes_value
        # Creación de pestaña por mes (si no existe)
        if sheet_1 in self.workbook.sheetnames:
            sheet = self.workbook[sheet_1]

        else:
            sheet = self.workbook.copy_worksheet(self.workbook['Hoja1'])
            sheet.title = sheet_1

            self.workbook.save('Producción {0}.xlsx'.format(self.year))

        # Procesamiento de Datos
        row = sheet.max_row + 1
        prod_day = int(self.pro_day)
        cant_prod = int(self.cantidad_prod.get())
        max_range = row + int(self.max_prod.get())
        # Registro Producción Fino
        sheet.cell(5, 17).value = '=SUM($G${0}:$G${1})'.format(5, row + 1)

        # Registro Producción Grueso
        sheet.cell(6, 17).value = '=SUM($N${0}:$N${1})'.format(5, row + 1)

        # Registro Merma
        sheet.cell(7, 17).value = '=SUM($J${0}:$J${1})'.format(5, row + 1)
        # Formato de Celdas

        thick = Side(border_style="medium", color="000000")
        thin = Side(border_style="thin", color="000000")

        # Estilo del formato inicial para la columna

        for j in range(row, max_range):
            for r in range(2, 15):
                sheet.cell(j, r).style = 'format_style'
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
            if (prod_day < 10 and sheet_2 < 10):
                sheet.cell(j, 2).value = 'L' + '0' + str(prod_day) + '0' + str(sheet_2) + '18'
            elif prod_day < 10:
                sheet.cell(j, 2).value = 'L' + '0' + str(prod_day) + str(sheet_2) + '18'
            elif sheet_2 < 10:
                sheet.cell(j, 2).value = 'L' + str(prod_day) + '0' + str(sheet_2) + '18'
            else:
                sheet.cell(j, 2).value = 'L' + str(prod_day) + str(sheet_2) + '18'
            sheet.cell(j, 3).value = '0' + str(j - (row - 1))

        # Carga alimentada + cantidad no alimentada

            sheet.cell(j, 4).value = int(cant_prod)
            sheet.cell(j, 5).value = 0.03

        # Valores Fracción Gruesa
        # VALOR A CAMBIAR FRACCION GRUESA input("Ingrese producción fracción gruesa {}: "
            sheet.cell(j, 6).value = int(str(self.entries[(j - row)]))

        # Valor Resultados

        for co in range(7, 15):
            if co != 8:
                if co == 7:
                    sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
                    sheet.cell(row, co).value = int(self.cant_fina.get())
                if co == 9:
                    sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
                    sheet.cell(row, co).value = '=SUM($F${0}:$G${1})'.format(row, (max_range - 1))
                if co == 10:
                    sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
                    sheet.cell(row, co).value = '=($D${0}*$C${1})-$I${0}'.format(row, (max_range - 1))
                if co == 11:
                    sheet.merge_cells(start_row=row, start_column=co, end_row=(max_range - 1), end_column=13)
                    sheet.cell(row, co).value = str(self.obs.get('1.0', 'end-2c'))
                if co == 14:
                    sheet.merge_cells(start_column=co, start_row=row, end_row=(max_range - 1), end_column=co)
                    sheet.cell(row, co).value = '=SUM($F${0}:$F${1})'.format(row, (max_range - 1))
            else:
                continue

        self.workbook.save('Producción {0}.xlsx'.format(self.year))
