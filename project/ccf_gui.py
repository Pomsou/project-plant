import tkinter as tk
import inventario_gui
import produccion_gui
import etiquetado_gui
from tkinter import font, messagebox, Button


class Window():

    def __init__(self, master):
        arial_new = tk.font.Font(family='Arial', size=15)
        # Frame etiquetas
        custom_frame = tk.LabelFrame(settings_windows, bg='#5475B1', bd=0)
        custom_frame.grid(row=1, column=0, pady=15, padx=15)

        # Primera aplicación
        inv_settings = inventario_gui.INV()
        w1 = Button(custom_frame, text='Inventario', command=inv_settings.inventario, height=5, font='Sunday', fg='#FFFFFF', bg='#151763', width=30)
        w1.grid(row=1, column=0, padx=10, pady=10, stick='e')

        # Segunda aplicación
        prod_settings = produccion_gui.PRD()
        w2 = Button(custom_frame, text='Producción', command=prod_settings.produccion, font='Sunday', height=5, fg='#FFFFFF', bg='#151763', width=30)
        w2.grid(row=2, column=0, padx=10, pady=10, stick='e')

        # Tercera aplicación
        etq_settings = etiquetado_gui.ETQ()
        w3 = Button(custom_frame, text='Etiquetado', command=etq_settings.etiquetado, font='Sunday', height=5, fg='#FFFFFF', bg='#151763', width=30)
        w3.grid(row=3, column=0, padx=10, pady=10, stick='e')


if __name__ == '__main__':
    settings_windows = tk.Tk()
    app = Window(settings_windows)
    settings_windows.title("Planta de Fraccionamiento CCF")
    settings_windows.geometry('330x420+500+160')
    settings_windows.resizable(0, 0)
    settings_windows.mainloop()
