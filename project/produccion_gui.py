import tkinter as tk
from tkinter import font, messagebox, Button, ttk
import tkinter.scrolledtext as tsk
import Produccion


def produccion(self):
    self.dialogo = tk.Toplevel()
    self.dialogo.geometry('500x400+350+100')
    self.dialogo.resizable(0, 0)
    self.dialogo.title('Producción')
    # Botón Cerrar y mantener window inicial
    boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy)
    boton.grid(row=3, column=3)
    botonp = Button(self.dialogo, text='Producción diaria (+)', command=prod_diaria(self))
    botonp.grid(row=1, column=0)
    self.dialogo.transient()
    self.dialogo.grab_set()
    self.dialogo.wait_window(self.dialogo)


def prod_diaria(self):
    self.dialogo = tk.Toplevel()
    self.dialogo.geometry('550x500+450+100')
    self.dialogo.resizable(0, 0)
    self.dialogo.title('Producción diaria')
    # Label producción
    customer_frame = tk.LabelFrame(self.dialogo, text='Fecha', bg='light steel blue')
    customer_frame.grid(row=0, column=0, padx=10, pady=20)
    customer_frame2 = tk.LabelFrame(self.dialogo, text='Producción', bg='light steel blue')
    customer_frame2.grid(row=0, column=3, padx=10, pady=10)
    customer_frame3 = tk.LabelFrame(self.dialogo, text='Observación', bg='light steel blue')
    customer_frame3.grid(row=1, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=5, pady=5)
    # dia de producción
    pro_day = tk.IntVar()
    etiq2 = tk.Label(customer_frame, text='Dia de Producción:')
    etiq2.grid(row=0, column=0, stick='w', padx=10, pady=10)
    day = ttk.Combobox(customer_frame, values=[i for i in range(1, 32)], textvariable=pro_day)
    day.grid(row=1, column=0, padx=10, pady=2)

    # Mes de producción
    def check_mes(event):
        date_prod = self.mes.get()

    etiq1 = tk.Label(customer_frame, text='Mes de Producción:')
    etiq1.grid(row=2, column=0, stick='w', padx=10, pady=10)
    mes = ttk.Combobox(customer_frame, values=('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'), state='readonly')
    mes.grid(row=3, column=0, padx=10, pady=2)
    mes.bind('<<ComboboxSelected>>', check_mes)

    # Carga Alimentada
    etiqc = tk.Label(customer_frame, text='Carga Alimentada').grid(row=4, column=0, stick='w', padx=10, pady=5)
    cant_carg = tk.Entry(customer_frame, width=5)
    cant_carg.grid(row=5, column=0, pady=10)

    # Número de lotes
    max_prod = tk.IntVar()

    def onClick(event=None):
        max_prod.set(max_prod.get() + 1)

    def action(i, label_act):
        label_act = tk.Label(customer_frame2, text='✔', bg='light steel blue')
        label_act.grid(row=i + 4, column=4)

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
    etiq4 = tk.Label(customer_frame2, text='Cantidad Producida').grid(row=3, column=3, padx=10, pady=10)

    def setText(max_prod):
        entries = []
        for i in range(max_prod.get()):
            label_prod = tk.Label(customer_frame2, text='%2d' % i)
            label_prod.grid(row=i + 4, column=3)
            entries.append(tk.Entry(customer_frame2, width=10))
            entries[i].grid(row=i + 4, column=3)
            entries[i].bind('<Return>', lambda event, r=i, lprod=label_prod: action(r, lprod))
            label_prod.config(text=str(int(label_prod.cget('text')) + 1))

    etiq3 = tk.Label(customer_frame2, text='Número de Lotes').grid(row=0, column=3, padx=10, pady=10)
    lot_boton = tk.Label(customer_frame2, textvariable=max_prod).grid(row=1, column=3)
    lotes = tk.Button(customer_frame2, text='+', command=lambda: [onClick(), setText(max_prod)], fg='gray24', bg='papaya whip')
    lotes.grid(row=1, column=4, padx=10, pady=10)
    max_produccion = max_prod

    # Cantidad de producción fina
    etiq5 = tk.Label(customer_frame2, text='Producción Fracción Fina').grid(row=10, column=3, padx=10, pady=5)
    cant_fina = tk.Entry(customer_frame2, width=10)
    cant_fina.grid(row=11, column=3, pady=10)
    cant_fina.bind('<Return>', action2)

    # Observaciones
    etiq6 = tk.Label(customer_frame3, text='Observaciones').grid(row=1, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=5, pady=5)
    obs = tsk.ScrolledText(master=customer_frame3, wrap=tk.WORD, width=45, height=4)
    obs.grid(row=1, column=0, rowspan=4, columnspan=5, sticky='W' + 'E' + 'N' + 'S', padx=10, pady=10)
    obs.bind('<Return><Return>', action3)

    # Actualizar Base de datos
    boton_act = tk.Button(self.dialogo, text='Actualizar BD', command=Produccion._produccion)
    boton_act.grid(row=3, column=5, padx=10)

    # Botón Cerrar y mantener window inicial
    boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy)
    boton.grid(row=4, column=5, padx=10, pady=10)
    self.dialogo.grab_set()
    self.dialogo.wait_window(self.dialogo)
