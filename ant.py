from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import organizarExcel as lRow
import tiparExcel as tipo
import tkinter as tk 
import testeArq
import organizarRFID as rfid

def addUnidade():
    frame = tk.Tk() 
    frame.title("Adicionar Unidade") 
    frame.geometry('300x200') 
    def printInput(): 
        global unidadeAdd
        unidadeAdd = inputtxt.get(1.0, "end-1c") 
        lbl.config(text = "Unidade adicionado: "+unidadeAdd) 
        lRow.unidade = unidadeAdd
    inputtxt = tk.Text(frame, 
                    height = 5, 
                    width = 20) 
    inputtxt.pack() 
    printButton = tk.Button(frame, 
                            text = "Adicionar",  
                            command = printInput) 
    printButton.pack() 
    closeButton = tk.Button(frame, 
                            text = "Confirmar",  
                            command = frame.destroy) 
    closeButton.pack() 
    lbl = tk.Label(frame, text = "") 
    lbl.pack() 
    frame.mainloop() 
def addDispositivo():
    frame = tk.Tk() 
    frame.title("Adicionar Equipamento") 
    frame.geometry('300x200') 
    def printInput(): 
        global equipamentoAdd
        equipamentoAdd = inputtxt.get(1.0, "end-1c") 
        lbl.config(text = "Equipamentos adicionados: "+equipamentoAdd) 
        lRow.equipamento = equipamentoAdd
    inputtxt = tk.Text(frame, 
                    height = 5, 
                    width = 20) 
    inputtxt.pack() 
    printButton = tk.Button(frame, 
                            text = "Adicionar",  
                            command = printInput) 
    printButton.pack() 
    closeButton = tk.Button(frame, 
                            text = "Confirmar",  
                            command = frame.destroy) 
    closeButton.pack() 
    lbl = tk.Label(frame, text = "") 
    lbl.pack() 
    frame.mainloop()
try:
    janela = Tk()
    frm = ttk.Frame(janela, padding=10)
    janela.title('Organizador de Excel')
    frm.grid()
    ttk.Label(frm, text="Organizador de Excel").grid(column=2, row=0)
    ttk.Button(frm, text='Local', command=testeArq.localArq).grid(column=1, row=1)
    ttk.Button(frm, text='Unidade', command=addUnidade).grid(column=4, row=1) 
    ttk.Button(frm, text='Equipamento', command=addDispositivo).grid(column=3, row=1)
    ttk.Label(frm, text="Organizar excel do cliente" ).grid(column=0, row=4)
    ttk.Button(frm, text='Organizar', command=lRow.organizadorEx).grid(column=4, row=4)
    ttk.Label(frm, text="Tipar dados para importação" ).grid(column=0, row=6)
    ttk.Button(frm, text='Organizar', command=tipo.tipagemDados).grid(column=4, row=6)
    ttk.Label(frm, text="Organizar RFID" ).grid(column=0, row=8)
    ttk.Button(frm, text='Organizar', command=rfid.organizarRfid).grid(column=4, row=8)
    janela.mainloop()
except:
    pass