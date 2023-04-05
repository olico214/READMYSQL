from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import csv
from openpyxl import Workbook
import openpyxl


class Buscador:
    def __init__(self):
        self.root = Tk()
        self.root.title("ACTUALIZADOR")
        self.root.geometry("720x480")
        
        self.cuerpo= ttk.Frame(self.root)
        self.cuerpo.place(x=0,y=0,height=480, width=720)
        
        
        
        self.historico = ttk.Button(self.root, text="Historico de ventas")
        self.historico.place(x=10,y=60,width=200)
        
        self.historico_90 = ttk.Button(self.root, text="INFORME A 90 DIAS")
        self.historico_90.place(x=230,y=60,width=200)
        
        
        
        
        
        self.root.mainloop()
        
        
        
aplicacion1 = Buscador()