import tkinter as tk
from tkinter import font, messagebox, Button
import etiqueta_milavena
import etiqueta_protavena
import etiqueta_fibravena
import etiqueta_lotes


class ETQ():

    def etiquetado(self):
        self.dialogo = tk.Toplevel()
        self.dialogo.geometry('420x570+450+80')
        self.dialogo.title('Etiquetado')

        # Design UI
        self.TopFrame = tk.Frame(self.dialogo, bg='#151763', bd=0, width=400, height=100)
        self.TopFrame.grid(column=0, row=0, ipadx=10, ipady=10, stick='we')

        # Frame etiquetas
        custom_frame = tk.LabelFrame(self.dialogo, bg='#5475B1', bd=0)
        custom_frame.grid(row=1, column=0, pady=10)

        # Etiquetas Milavena
        label_mil = tk.Button(custom_frame, text='Etiquetas Milavena', command=etiqueta_milavena.MIL().milavena, font='Sunday', height=3, fg='#FFFFFF', bg='#151763', width=20)
        label_mil.grid(row=1, column=0, padx=10, pady=10, stick='w')

        # Etiquetas Protavena
        label_prot = tk.Button(custom_frame, text='Etiquetas Protavena', command=etiqueta_protavena.PROT().protavena, font='Sunday', height=3, fg='#FFFFFF', bg='#151763', width=20)
        label_prot.grid(row=2, column=0, padx=10, pady=10, stick='w')

        # Etiquetas Fibravena
        label_prot = tk.Button(custom_frame, text='Etiquetas Fibravena', command=etiqueta_fibravena.FIB().fibravena, font='Sunday', height=3, fg='#FFFFFF', bg='#151763', width=20)
        label_prot.grid(row=3, column=0, padx=10, pady=10, stick='w')

        # Etiquetas Lotes
        label_lot = tk.Button(custom_frame, text='Etiquetas Lotes', command=etiqueta_lotes.LOT()._lotes, font='Sunday', height=3, fg='#FFFFFF', bg='#151763', width=20)
        label_lot.grid(row=4, column=0, padx=10, pady=10, stick='w')

        # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy, bd=0, fg='#FFFFFF', bg='#151763', height=2, width=10)
        boton.grid(row=5, column=0, pady=10)
        self.dialogo.transient()
        self.dialogo.grab_set()
        self.dialogo.wait_window()
