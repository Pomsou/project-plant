import tkinter as tk
from tkinter import font, messagebox, Button


def inventario(self):
    self.dialogo = tk.Toplevel()
    self.dialogo.geometry('600x450+350+100')
    self.dialogo.resizable(0, 0)
    self.dialogo.title('Inventario')
    # Botón Cerrar y mantener window inicial
    boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy)
    boton.pack(side='bottom', padx=20, pady=20)
    self.dialogo.transient()
    self.dialogo.grab_set()
    self.dialogo.wait_window(self.dialogo)
