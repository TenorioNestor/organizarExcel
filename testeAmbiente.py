import pandas as pd
import xlsxwriter
import sys
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
import testeArq



def organizadorEx():
    try:
        workbook = xlsxwriter.Workbook('Ambientes.xlsx')
        worksheet = workbook.add_worksheet()
        local = testeArq.file_path
        dataBl1 = open('AmbientesTeste','rb')
        df_bl1 = pd.read_excel(dataBl1)
        worksheet.set_column('A:A', len(df_bl1.index))
        worksheet.set_column('B:B', len(df_bl1.index))

            #--------------nome-------------------
        try:
            lNome = 1
            worksheet.write(0, 2, 'NOME')
            contNome = 0
            for i in range(len(df_bl1.index)):

                nome = df_bl1['NOME'].loc[i]  
                worksheet.write(lNome, 2, nome)
                lNome = lNome + 1 
                contNome = contNome + 1  
            print('Nome OK')
        except:messagebox.showinfo("Erro", "Erro na coluna NOME, verifique o conteudo")

        #-----------------AMBIENTE-----------------
        try:
            linha1 = 1
            worksheet.write(0, 11, "AMBIENTE")
            df_bl1['AMBIENTE'] = df_bl1['AMBIENTE'].map(str)
            for i in range(len(df_bl1.index)):
                ambiente = df_bl1['AMBIENTE'].loc[i]
                worksheet.write(linha1, 11, ambiente, format1)
                linha1 = linha1 + 1
            print('Ambiente OK')
        except:messagebox.showinfo("Erro", "Erro na coluna AMBIENTE, verifique o conteudo")
   
        except:messagebox.showinfo("Erro", "Erro na coluna UNIDADE, verifique o conteudo")
        messagebox.showinfo("Title", "Arquivo organizado\n Procure o arquivo CadRef")    
        workbook.close()
    except:
        messagebox.showinfo("Title", "Erro\n Verifique o arquivo Excel")