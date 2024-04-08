import tkinter as tk
from tkinter import filedialog
from openpyxl import *

def localArq():
    def select_file():
        global file_path
        file_path = filedialog.askopenfilename()
        file_label.config(text="Arquivo selecionado: " + file_path)
    root = tk.Tk()
    root.title("Selecionar arquivo para importar")
    file_button = tk.Button(root, text="Selecionar arquivo", command=(select_file))
    file_button.pack(pady=20)
    close_button = tk.Button(root, text="Confirmar", command=(root.destroy))
    close_button.pack(pady=40)
    file_label = tk.Label(root, text="")
    file_label.pack()
    root.mainloop()
