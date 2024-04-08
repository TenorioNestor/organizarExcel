import pandas as pd
import xlsxwriter
import sys
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
import testeArq

global unidade
def tela():
    frame = Tk() 
    frame.title("TextBox Input") 
    frame.geometry('400x200') 
    def printInput(): 
        inp = inputtxt.get(1.0, "end-1c") 
        lbl.config(text = "Provided Input: "+inp) 
    inputtxt = ttk.E(frame, 
                    height = 5, 
                    width = 20) 
    inputtxt.pack() 
    printButton = ttk.Button(frame, 
                            text = "Print",  
                            command = printInput) 
    printButton.pack() 
    lbl = ttk.Label(frame, text = "") 
    lbl.pack() 
def organizarRfid():
    try:
        workbook = xlsxwriter.Workbook('Rfid.xlsx')
        worksheet = workbook.add_worksheet()
        localRfid = testeArq.file_path
        dataBl1 = open(f"{localRfid}",'rb')
        df_bl1 = pd.read_excel(dataBl1)
        worksheet.set_column('A:A', len(df_bl1.index))
        worksheet.set_column('B:B', len(df_bl1.index))
        worksheet.set_column('C:C', len(df_bl1.index))
        worksheet.set_column('D:D', len(df_bl1.index))
        worksheet.set_column('E:E', len(df_bl1.index))

        format1 = workbook.add_format({'num_format': '@'})
        format3 = workbook.add_format({'num_format': '0'})
        print('=============================================================')

        #----------UNIDADE--------------------------
        try:
            lUnidade = 1
            worksheet.write(0, 1, 'UNIDADE')
            for i in range(len(df_bl1.index)):
                worksheet.write(lUnidade, 1, unidade, format3)
                lUnidade = lUnidade + 1
            print('unidade OK')
        except:messagebox.showinfo("Erro", "Erro na coluna UNIDADE, verifique o conteudo")
        #--------------codigo-------------------
        try:
            lcodigo = 1
            worksheet.write(0, 2, 'CODIGO')
            contCodigo = 0
            for i in range(len(df_bl1.index)):
                codigo = df_bl1['CODIGO'].loc[i]  
                if len(codigo) == 0:
                    codigo = 'null'
                    worksheet.write(lcodigo, 2, codigo,format1)
                else:
                    worksheet.write(lcodigo, 2, codigo,format1)
                lcodigo = lcodigo + 1 
                contCodigo = contCodigo + 1  
            print('codigo OK')
        except:messagebox.showinfo("Erro", "Erro na coluna CODIGO, verifique o conteudo")
        #--------------hex-------------------
        try:
            lHex = 1
            worksheet.write(0, 3, 'cod_hexadecimal')
            conthex = 0
            for i in range(len(df_bl1.index)):
                hex = df_bl1['cod_hexadecimal'].loc[i]
                if len(hex) == 0:
                    hex = 'null'  
                    worksheet.write(lHex, 3, hex,format1)
                else:
                    worksheet.write(lHex, 3, hex,format1)
                lHex = lHex + 1 
                conthex = conthex + 1  
            print(hex) 
        except:#messagebox.showinfo("Erro", "Erro na coluna Cod_Hex, verifique o conteudo")
            for i in range(len(df_bl1.index)):
                        hex = 'null'
                        worksheet.write(lHex, 4, hex,format1)
                        lHex = lHex + 1 
                        conthex = conthex + 1  
                        print('hex OK')

        #--------------hex_cal-------------------
        try:
            lCal = 1
            worksheet.write(0, 4, 'hex_cal')
            contCal = 0
            if len(df_bl1['hex_cal']) != 0:
                for i in range(len(df_bl1.index)):
                    calHex = df_bl1['hex_cal'].loc[i]  
                    if len(calHex) == 0:
                        calHex = 'S'
                        worksheet.write(lCal, 4, calHex,format1)
                    else:
                        worksheet.write(lCal, 4, calHex,format1)
                    lCal = lCal + 1 
                    contCal = contCal + 1  
                print('hex OK')
            else:
                for i in range(len(df_bl1.index)):
                    calHex = 'S'
                    worksheet.write(lCal, 4, calHex,format1)
                    lCal = lCal + 1 
                    contCal = contCal + 1  
                print('hex OK')
        except:#messagebox.showinfo("Erro", "Erro na coluna hex_cal, verifique o conteudo")
            for i in range(len(df_bl1.index)):
                        calHex = 'S'
                        worksheet.write(lCal, 4, calHex,format1)
                        lCal = lCal + 1 
                        contCal = contCal + 1  
                        print('hex OK')
        messagebox.showinfo("success", "Arquivo organizado\n Procure o arquivo CadRef")    
        workbook.close()
    except:
        messagebox.showinfo("Erro", "Erro\n Verifique o arquivo Excel")