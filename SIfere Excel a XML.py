#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
from BIN.sifere import Procesar_SIFERE_Excel_to_XML
import os

# Abrir Carpeta de Resultados
def Abrir_Carpeta_Resultados():
    os.system("start Resultados")
    
# Abrir Cafecito
def Cafecito():
    os.system("start https://cafecito.app/abustos")
    
# Abrir Excel de Ejemplo
def Abrir_Excel_ejemplo():
    os.system("start Base.xlsx")


class App:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=250,
            width=325)
        Toplevel_1.iconbitmap("BIN/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.minsize(325, 250)
        Toplevel_1.overrideredirect("False")
        Toplevel_1.resizable(False, False)
        Toplevel_1.title("SIFERE XML")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="BIN/ABP-blanco-sin-fondo.png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            takefocus=False,
            text='Transformar Papel de Tabajo de SIFERE a XML para su importación en AFIP\n',
            wraplength=325)
        Label_1.pack(expand=True, side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand=True, side="top")
        self.Abrir_Excel_ejemplo = ttk.Button(Toplevel_1)
        self.Abrir_Excel_ejemplo.configure(text='Abrir Excel de Ejemplo', command=Abrir_Excel_ejemplo)
        self.Abrir_Excel_ejemplo.pack(expand=True, pady=4, side="top")
        self.Abrir_Excel = ttk.Button(Toplevel_1)
        self.Abrir_Excel.configure(text='Procesar Excel', command=Procesar_SIFERE_Excel_to_XML)
        self.Abrir_Excel.pack(expand=True, pady=4, side="top")
        self.Resultados = ttk.Button(Toplevel_1)
        self.Resultados.configure(text='Abrir Carpeta de Resultados', command=Abrir_Carpeta_Resultados)
        self.Resultados.pack(expand=True, padx=0, pady=4, side="top")
        self.Colaboraciones = ttk.Button(Toplevel_1)
        self.Colaboraciones.configure(text='Colaboraciones', command=Cafecito)
        self.Colaboraciones.pack(expand=True, pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
