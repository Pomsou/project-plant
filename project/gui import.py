import tkinter as tk
from tkinter import font, messagebox, Button


class Window():

    def __init__(self):
        self.raiz = tk.Tk()
        self.raiz.title("Planta de Fraccionamiento CCF")
        self.raiz.geometry('300x400+500+160')
        self.raiz.resizable(0, 0)
        arial_new = tk.font.Font(family='Arial', size=15)
        # Primera aplicación
        w1 = Button(self.raiz, text='Inventario', command=self.inventario, height='3', width='15', font=arial_new)
        w1.pack(padx=30, pady=30)

        # Segunda aplicación
        w2 = Button(self.raiz, text='Producción', command=self.produccion, height='3', width='15', font=arial_new)
        w2.place(x=65, y=150)

        # Tercera aplicación
        w3 = Button(self.raiz, text='Etiquetado', command=self.etiquetado, height='3', width='15', font=arial_new)
        w3.place(x=65, y=270)

        self.raiz.mainloop()

    def produccion(self):
        self.dialogo = tk.Toplevel()
        self.dialogo.geometry('600x450+350+100')
        self.dialogo.resizable(0, 0)
        self.dialogo.title('Producción')
    # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo, text='Cerrar', command=self.dialogo.destroy)
        boton.pack(side='bottom', padx=20, pady=20)
        self.dialogo.transient(master=self.raiz)
        self.dialogo.grab_set()
        self.raiz.wait_window(self.dialogo)

    def inventario(self):
        self.dialogo = tk.Toplevel()
        self.dialogo.geometry('600x450+350+100')
        self.dialogo.resizable(0, 0)
        self.dialogo.title('Inventario')
        # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo, text='Cerrar',
                       command=self.dialogo.destroy)
        boton.pack(side='bottom', padx=20, pady=20)
        self.dialogo.transient(master=self.raiz)
        self.dialogo.grab_set()
        self.raiz.wait_window(self.dialogo)

    def etiquetado(self):
        self.dialogo = tk.Toplevel()
        self.dialogo.geometry('600x450+350+100')
        self.dialogo.resizable(0, 0)
        self.dialogo.title('Etiquetado')
        # Botón Cerrar y mantener window inicial
        boton = Button(self.dialogo, text='Cerrar',
                       command=self.dialogo.destroy)
        boton.pack(side='bottom', padx=20, pady=20)
        self.dialogo.transient(master=self.raiz)
        self.dialogo.grab_set()
        self.raiz.wait_window(self.dialogo)


def main():
    mi_app = Window()
    return(0)


if __name__ == '__main__':
    main()
