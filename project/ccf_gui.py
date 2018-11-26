import tkinter as tk
import inventario_gui
import produccion_gui
import etiquetado_gui
from tkinter import font, messagebox, Button


class Window():

    def __init__(self, master):
        arial_new = tk.font.Font(family='Arial', size=15)
        # Primera aplicación
        w1 = Button(settings_windows, text='Inventario', command=inventario_gui.inventario, height='3', width='15', font=arial_new)
        w1.pack(padx=30, pady=30)

        # Segunda aplicación
        prod_settings = produccion_gui.PRD()
        w2 = Button(settings_windows, text='Producción', command=prod_settings.produccion, height='3', width='15', font=arial_new)
        w2.place(x=65, y=150)

        # Tercera aplicación
        w3 = Button(settings_windows, text='Etiquetado', command=etiquetado_gui.etiquetado, height='3', width='15', font=arial_new)
        w3.place(x=65, y=270)


if __name__ == '__main__':
    settings_windows = tk.Tk()
    app = Window(settings_windows)
    settings_windows.title("Planta de Fraccionamiento CCF")
    settings_windows.geometry('300x400+500+160')
    settings_windows.resizable(0, 0)
    settings_windows.mainloop()
