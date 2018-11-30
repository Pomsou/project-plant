import tkinter as tk
from tkinter import font, messagebox, Button, ttk


class INV():
    def inventario(self):
        self.dialogo = tk.Toplevel()
        self.dialogo.geometry('420x400+450+80')
        self.dialogo.resizable(0, 0)
        self.dialogo.title('Inventario')

        # Diseño UI
        self.TopFrame = tk.Frame(self.dialogo, bg='#151763', bd=0, width=400, height=100)
        self.TopFrame.grid(column=0, row=0, ipadx=10, ipady=10, stick='we')

        # Frame etiquetas
        custom_frame = tk.LabelFrame(self.dialogo, bg='#5475B1', bd=0)
        custom_frame.grid(row=1, column=0, pady=10)

        # Inventario total
        label_mil = tk.Button(custom_frame, text='Inventario Físico', command=self.inventario_fisico, font='Sunday', height=3, fg='#FFFFFF', bg='#151763', width=20)
        label_mil.grid(row=1, column=0, padx=10, pady=10, stick='w')

        # Inventario contable
        label_prot = tk.Button(custom_frame, text='Inventario Perpetuo', command=self.inventario_perpetuo, font='Sunday', height=3, fg='#FFFFFF', bg='#151763', width=20)
        label_prot.grid(row=2, column=0, padx=10, pady=10, stick='w')

        # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy, bd=0, fg='#FFFFFF', bg='#151763', height=2, width=10)
        boton.grid(row=3, column=0, pady=10)
        self.dialogo.transient()
        self.dialogo.grab_set()
        self.dialogo.wait_window()

    def inventario_fisico(self):
        self.dialogo2 = tk.Toplevel()
        self.dialogo2.geometry('1000x400+300+120')
        self.dialogo2.resizable(0, 0)
        self.dialogo2.title('Inventario Físico')

        # Diseño UI
        customer_frame2 = tk.LabelFrame(self.dialogo2, bg='#FFFFFF', bd=2)
        customer_frame2.grid(row=0, column=0, padx=10, pady=10)
        customer_frame3 = tk.LabelFrame(self.dialogo2, bg='#FFFFFF', bd=2)
        customer_frame3.grid(row=0, column=1, padx=10, pady=10)
        # Box Lotes en planta
        self.button = []
        cont_check = [i for i in range(36)]
        cont = 0
        for i in range(7):
            for j in range(0, 5):
                if i == 4 and j == 2:
                    self.button.append(Button(customer_frame2, text='Planta', command=lambda i=cont + 1: open_this(i)))
                    self.button[cont].grid(row=i, column=j, sticky='e', padx=10, pady=10)
                else:
                    self.button.append(Button(customer_frame2, text='N° ' + str(cont + 1), command=lambda i=cont + 1: open_this(i)))
                    self.button[cont].grid(row=i, column=j, sticky='e', padx=10, pady=10)
                cont = cont + 1

        # Box Revisión Pallet
        def open_this(myNum):
            self.max_valor = tk.IntVar()
            # Box Revisión Pallet

            def onClick():
                self.max_valor.set(self.max_valor.get() + 1)
                self.num_lote = []
                self.cant_tipo = []
                self.tipo_formato = []
                self.cant_form = []
                self.cant_prod = []
                for i in range(self.max_valor.get()):
                    # if myNum in [0, 11]:
                    # Tipo de Producto
                    self.label_prod = ttk.Combobox(customer_frame3, values=['Milavena', 'Protavena', 'Fibravena'], state='readonly')
                    self.label_prod.grid(row=i + 2, column=1)
                    self.cant_tipo.append(self.label_prod)

                    # Número de Lote
                    self.c = tk.Entry(customer_frame3, width=10)
                    self.c.grid(row=i + 2, column=2)
                    self.num_lote.append(self.c)

                    # Tipo de envase
                    self.form_prd = ttk.Combobox(customer_frame3, values=['Saco', 'Bolsa'], state='readonly', width=10)
                    self.form_prd.grid(row=i + 2, column=3)
                    self.tipo_formato.append(self.form_prd)

                    # Formato
                    self.form_env = tk.Entry(customer_frame3, width=10)
                    self.form_env.grid(row=i + 2, column=4)
                    self.cant_form.append(self.form_env)

                    # Cantidad Producida
                    self.max_prod = ttk.Combobox(customer_frame3, values=[x for x in range(30)], state='readonly')
                    self.max_prod.grid(row=i + 2, column=5)
                    self.cant_prod.append(self.max_prod)

            def action(event):
                for i in range(self.max_valor.get()):
                    self.label_act = tk.Label(customer_frame3, text='✔', fg='#151763')
                    self.label_act.grid(row=i + 2, column=6)

            def setText():
                self.entries = []
                self.entries2 = []
                self.entries3 = []
                self.entries4 = []
                self.entries5 = []
                for self.c in self.num_lote:
                    self.entries.append(self.c.get())
                for self.label_prod in self.cant_tipo:
                    self.entries2.append(self.label_prod.get())
                for self.form_prd in self.tipo_formato:
                    self.entries3.append(self.form_prd.get())
                for self.form_env in self.cant_form:
                    self.entries4.append(self.form_env.get())
                for self.max_prod in self.cant_prod:
                    self.entries5.append(self.max_prod.get())

            etiq = tk.Label(customer_frame3, text='Pallet #{0}'.format(myNum), fg='#FFFFFF', bg='#151763', width=25).grid(row=0, column=1, padx=10, pady=10)
            label_tlot = tk.Label(customer_frame3, text='Tipo de Producto', fg='#FFFFFF', bg='#151763', width=15)
            label_tlot.grid(row=1, column=1, padx=5, pady=10)
            label_lot = tk.Label(customer_frame3, text='Lote', fg='#FFFFFF', bg='#151763', width=5)
            label_lot.grid(row=1, column=2, padx=5, pady=10)
            label_env = tk.Label(customer_frame3, text='Tipo de envase', fg='#FFFFFF', bg='#151763', width=15)
            label_env.grid(row=1, column=3, padx=5, pady=10)
            label_form = tk.Label(customer_frame3, text='Formato (kg)', fg='#FFFFFF', bg='#151763', width=10)
            label_form.grid(row=1, column=4, padx=5, pady=10)
            label_cant = tk.Label(customer_frame3, text='Cantidad', fg='#FFFFFF', bg='#151763', width=10)
            label_cant.grid(row=1, column=5, padx=5, pady=10)
            lotes = tk.Button(customer_frame3, text='+', command=onClick, fg='gray24', bg='papaya whip')
            lotes.grid(row=0, column=4, stick='e')
            boton_max = tk.Button(customer_frame3, text='Set', command=setText, fg='gray24', bg='papaya whip')
            boton_max.grid(row=0, column=5, pady=10, padx=10)
            boton_max.bind('<Button-1>', action)

        # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo2, text='Cerrar', command=self.dialogo2.destroy, bd=0, fg='#FFFFFF', bg='#151763', height=2, width=10)
        boton.grid(row=7, column=1, pady=10)
        self.dialogo2.transient()
        self.dialogo2.grab_set()
        self.dialogo2.wait_window()

    def inventario_perpetuo(self):
        self.dialogo.geometry('520x230+350+150')
        self.dialogo.title('Inventario Físico')
