import tkinter as tk
from tkinter import font, messagebox, Button, ttk
import tkinter.scrolledtext as tsk
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font, NamedStyle
import os


class PRD:

    def produccion(self):
        dialogo = tk.Toplevel()
        dialogo.geometry('500x400+350+100')
        dialogo.resizable(0, 0)
        dialogo.title('Producción')
        # Botón Cerrar y mantener window inicial
        boton = Button(dialogo, text='Cerrar', command=dialogo.destroy)
        boton.grid(row=3, column=3)
        botonp = Button(dialogo, text='Producción diaria (+)', command=self.prod_diaria())
        botonp.grid(row=1, column=0)
        dialogo.transient()
        dialogo.grab_set()
        dialogo.wait_window(dialogo)

    def prod_diaria(self):
        dialogo2 = tk.Toplevel()
        dialogo2.geometry('550x500+450+100')
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
            self.dia_prod = mes.get()
        etiq1 = tk.Label(customer_frame, text='Mes de Producción:')
        etiq1.grid(row=2, column=0, stick='w', padx=10, pady=10)
        mes = ttk.Combobox(customer_frame, values=('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'), state='readonly')
        mes.grid(row=3, column=0, padx=10, pady=2)
        mes.bind("<<ComboboxSelected>>", check_prod)

        # Carga Alimentada
        etiqc = tk.Label(customer_frame, text='Carga Alimentada').grid(row=4, column=0, stick='w', padx=10, pady=5)
        self.cantidad_prod = tk.Entry(customer_frame, width=5)
        self.cantidad_prod.grid(row=5, column=0, pady=10)

        # Número de lotes
        self.max_prod = tk.IntVar()

        def onClick():
            self.max_prod.set(self.max_prod.get() + 1)
            self.lot_boton = tk.Label(customer_frame2, textvariable=self.max_prod).grid(row=1, column=3)
            print(self.max_prod.get())
            self.fr_gruesa = []
            for i in range(self.max_prod.get()):
                label_prod = tk.Label(customer_frame2, text='%2d.-' % (i + 1))
                label_prod.grid(row=i + 4, column=2)
                self.c = tk.Entry(customer_frame2, width=10)
                self.c.grid(row=i + 4, column=3)
                self.fr_gruesa.append(self.c)

        def action(event):
            self.label_act = tk.Label(customer_frame2, text='✔', bg='light steel blue')
            self.label_act.grid(row=i + 4, column=4)

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
            for self.c in self.fr_gruesa:
                self.entries = self.c.get()

        etiq4 = tk.Label(customer_frame2, text='Cantidad Producida').grid(row=3, column=3, padx=10, pady=10)

        etiq3 = tk.Label(customer_frame2, text='Número de Lotes').grid(row=0, column=3, padx=10, pady=10)
        lotes = tk.Button(customer_frame2, text='+', command=onClick, fg='gray24', bg='papaya whip')
        lotes.grid(row=1, column=4, padx=10, pady=10)
        boton_max = tk.Button(customer_frame2, text='Set', command=setText, fg='gray24', bg='papaya whip')
        boton_max.grid(row=1, column=5, pady=10)
        # Cantidad de producción fina
        etiq5 = tk.Label(customer_frame2, text='Producción Fracción Fina').grid(row=10, column=3, padx=10, pady=5)
        self.cant_fina = tk.Entry(customer_frame2, width=10)
        self.cant_fina.grid(row=11, column=3, pady=10)
        self.cant_fina.bind('<Return>', action2)

        # Observaciones
        etiq6 = tk.Label(customer_frame3, text='Observaciones').grid(row=1, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=5, pady=5)
        self.obs = tsk.ScrolledText(master=customer_frame3, wrap=tk.WORD, width=45, height=4)
        self.obs.grid(row=1, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=10, pady=10)
        self.obs.bind('<Return><Return>', action3)

        # Actualizar Base de datos
        boton_act = tk.Button(dialogo2, text='Actualizar BD', command=self._produccion)
        boton_act.grid(row=3, column=5, padx=10)
        # Botón Cerrar y mantener window inicial
        boton = Button(dialogo2, text='Cerrar', command=dialogo2.destroy)
        boton.grid(row=4, column=5, padx=10, pady=10)
        dialogo2.grab_set()
        dialogo2.wait_window(dialogo2)

    def _produccion(self):
        file_location = "//Gcl-file-sr/CCF/03 Planta de Fraccionamiento/Operativo/Producción"
        os.chdir(file_location)
        reveal = os.getcwd()
        workbook = openpyxl.load_workbook("Producción 2018.xlsx")
        sheet_1 = self.dia_prod
        sheet = workbook[sheet_1]

        # Procesamiento de Datos

        prod_day = int(self.pro_day)
        row = sheet.max_row + 1
        cant_prod = int(self.cantidad_prod.get())
        max_range = row + int(self.max_prod.get())

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
                sheet.cell(j, 2).value = 'L' + '0' + str(prod_day) + sheet_1[:3] + '18'
            else:
                sheet.cell(j, 2).value = 'L' + str(prod_day) + str(sheet_prov) + '18'
            sheet.cell(j, 3).value = '0' + str(j - (row - 1))

        # Carga alimentada + cantidad no alimentada

            sheet.cell(j, 4).value = int(cant_prod)
            sheet.cell(j, 5).value = 0.03

        # Valores Fracción Gruesa
        # VALOR A CAMBIAR FRACCION GRUESA input("Ingrese producción fracción gruesa {}: "
            print(self.entries)
            sheet.cell(j, 6).value = int(str(self.entries[{0}]).format(j - row))

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

        workbook.save('Producción 2018.xlsx')
