import pandas as pd #Lib para trabalhar com tabelas
import xlsxwriter #Lib para trabalhar com tabelas
import sys
from openpyxl.utils.dataframe import dataframe_to_rows #Lib para trabalhar com tabelas
from tkinter import * #Lib para parte visual
from tkinter import ttk
import tkinter as tk 
from tkinter import messagebox
import datetime 
import testeArq

global unidade
global equipamento
global bloco
def log():
    root = tk.Tk()
    root.title('Log')
    root.geometry("400x700")
    label = tk.Label(root, text="Prints")
    label.pack()
    sys.stdout.write = lambda x: label.config(text=label.cget("text") + "\n" + x)
    root.mainloop()
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
def organizadorEx():
    try:
        workbook = xlsxwriter.Workbook('CadRef.xlsx')
        worksheet = workbook.add_worksheet()
        local = testeArq.file_path
        dataBl1 = open(f"{local}",'rb')
        df_bl1 = pd.read_excel(dataBl1)
        worksheet.set_column('A:A', len(df_bl1.index))
        worksheet.set_column('B:B', len(df_bl1.index))
        worksheet.set_column('C:C', len(df_bl1.index))
        worksheet.set_column('D:D', len(df_bl1.index))
        worksheet.set_column('E:E', len(df_bl1.index))
        worksheet.set_column('F:F', len(df_bl1.index))
        worksheet.set_column('G:G', len(df_bl1.index))
        worksheet.set_column('H:H', len(df_bl1.index))
        worksheet.set_column('I:I', len(df_bl1.index))
        worksheet.set_column('J:J', len(df_bl1.index))
        worksheet.set_column('K:K', len(df_bl1.index))
        worksheet.set_column('L:L', len(df_bl1.index))
        worksheet.set_column('M:M', len(df_bl1.index))
        worksheet.set_column('N:N', len(df_bl1.index))
        worksheet.set_column('O:O', len(df_bl1.index))
        format1 = workbook.add_format({'num_format': '@'})
        format2 = workbook.add_format({'num_format': 'dd/mm/yy'})
        format3 = workbook.add_format({'num_format': '0'})
        print('=============================================================')
            #----------ANIVERSARIO-------------------
        try:
            lData = 1
            date = datetime.date(1930,1,1) #Define o formato da data
            worksheet.write(0, 0, 'NASCIMENTO')
            for i in range(len(df_bl1.index)):
                try:
                    aniversario = df_bl1['NASCIMENTO'].loc[i]
                    worksheet.write(lData, 0, aniversario,format2)
                except:
                    aniversario = date
                    print('Except')
                    worksheet.write(lData, 0, aniversario,format2)
                lData = lData + 1
            print('Aniversario OK')
        except:messagebox.showinfo("Erro", "Erro na coluna ANIVERSARIO, verifique o conteudo")
        #----------CPF--------------------------
        try:
            lCpf = 1
            try:
                df_bl1['CPF'] = df_bl1['CPF'].fillna(0)
            except:
                pass
            worksheet.write(0, 1, 'CPF')
            for i in range(len(df_bl1.index)):
                try:
                    cpf = df_bl1['CPF'].loc[i]
                except:
                    cpf = 0
                worksheet.write(lCpf, 1, cpf, format3)
                lCpf = lCpf + 1
            print('CPF OK')
            #--------------nome-------------------
            df_bl1['NOME'] = df_bl1['NOME'].str.replace(",", ";")
            df_bl1['NOME'] = df_bl1['NOME'].str.replace("/", ";")
            df_bl1['NOME'] = df_bl1['NOME'].str.replace(" e ", ";")
            df_bl1['NOME'] = df_bl1['NOME'].str.replace(" E ", ";")
            lNome = 1
            worksheet.write(0, 2, 'NOME')
            contNome = 0
            for i in range(len(df_bl1.index)):
                try:
                    nome = df_bl1['NOME'].loc[i]
                    divisao = ';'
                    if divisao in nome:
                        nome = nome[:divisao]
                    else:
                        nome = nome    
                    worksheet.write(lNome, 2, nome,format1)
                    lNome = lNome + 1 
                except:
                    nome = df_bl1['NOME'].loc[i]
                    worksheet.write(lNome, 2, nome,format1)
                    lNome = lNome + 1
                contNome = contNome + 1  
            print('Nome OK')
        except:messagebox.showinfo("Erro", "Erro na coluna NOME, verifique o conteudo")
        #-----------------RG---------------------
        try:
            lRg = 1
            try:
                df_bl1['RG'] = df_bl1['RG'].fillna('null').map(str)
            except:
                rg = 'null'
            worksheet.write(0, 3, "RG")
            for i in range(len(df_bl1.index)):
                try:
                    rg = df_bl1['RG'].loc[i]
                except:
                    rg = 'null'
                worksheet.write(lRg, 3, rg, format1)
                lRg = lRg + 1
            print('RG OK')
        except:messagebox.showinfo("Erro", "Erro na coluna RG, verifique o conteudo")
        #----------TELEFONE--------------------------
        try:  
            lTelefone = 1
            worksheet.write(0, 4, "TELEFONE")
            try:
                df_bl1['TELEFONE'] = df_bl1['TELEFONE'].fillna(0)
            except:
                pass
            for i in range(len(df_bl1.index)):
                try:
                    telefone = df_bl1['TELEFONE'].loc[i]
                except:
                    telefone = 0
                worksheet.write(lTelefone, 4, telefone, format3)
                lTelefone = lTelefone + 1
            print('Telefone OK')
        except:messagebox.showinfo("Erro", "Erro na coluna TELEFONE, verifique o conteudo")
        #------------------EMAIL------------------
        try:
            lEmail = 1
            contEmail = 0
            df_bl1['EMAIL'] = df_bl1['EMAIL'].fillna('null').map(str)
            worksheet.write(0, 5, "EMAIL")
            for i in range(len(df_bl1.index)):
                try:
                    email = df_bl1['EMAIL'].loc[i]  
                    worksheet.write(lEmail, 5, email)
                    divisao = ';'
                    if divisao in email:
                        email = email[:divisao]
                    else:
                        email = email    
                    lEmail = lEmail + 1
                except:
                    email = df_bl1['EMAIL'].loc[i]  
                    worksheet.write(lEmail, 5, email)
                    lEmail = lEmail + 1
                contEmail = contEmail +1
            print('Email OK')
        except:messagebox.showinfo("Erro", "Erro na coluna EMAIL, verifique o conteudo")
        #------------------Celular------------------
        try:
            lCelular = 1
            worksheet.write(0, 6, "CELULAR")
            try:
                df_bl1['CELULAR'] = df_bl1['CELULAR'].fillna(0)
            except:
                pass
            for i in range(len(df_bl1.index)):
                try:
                    celular = df_bl1['CELULAR'].loc[i]
                except:
                    celular = 0
                worksheet.write(lCelular, 6, celular, format3)
                lCelular = lCelular + 1
            print('Celular OK')
        except:messagebox.showinfo("Erro", "Erro na coluna CELULAR, verifique o conteudo")
        #--------------------------Perfil*-----------------------
        try:
            lPerfil = 1
            df_bl1['PERFIL'] = df_bl1['PERFIL'].fillna('2')
            worksheet.write(0, 7, "PERFIL")
            for i in range(len(df_bl1.index)):
                try:
                    Perfil = df_bl1['PERFIL'].loc[i]
                    Prop = 'Propri' #2
                    PROP = 'PROPR' #2
                    MORADOR = 'MORADOR' #2
                    morador = 'morador' #2
                    Morador = 'Morador' #2
                    DEPENDENTE =  'DEPENDENTE'
                    dependente = 'dependente'
                    sindico =  'ndico' #1  
                    SINDICO =  'NDICO' #1
                    filho = 'filh' # 3
                    FILHO = 'FILHO' # 3
                    funcionario = 'funcion' #8
                    FUNCIONARIO = 'FUNCION' #8
                    locatario = 'locat' #23
                    LOCATARIO = 'LOCAT' #23
                    Prestador = 'Prestador'#6
                    PRESTADOR = 'PRESTADOR'#6
                    conjugue = 'espos' #3
                    ESPOS = 'ESPOS' #3
                    mae = 'mae'#9
                    MAE = 'MAE'#9
                    pai = 'pai'#9
                    PAI = 'PAI'#9
                    Zelador = 'Zelador'
                    ZELADOR = 'ZELADOR'
                    Inquilino = 'Inquilino'
                    INQUILINO = 'INQUILINO'
                    inquilino = 'inquilino'

                    if Prop in Perfil or PROP in Perfil or Inquilino in Perfil or INQUILINO in Perfil or MORADOR in Perfil or morador in Perfil or Morador in Perfil or inquilino in Perfil:
                        Perfil = '2'
                        worksheet.write(lPerfil, 7, Perfil)

                    elif filho in Perfil or conjugue in Perfil or FILHO in Perfil or ESPOS in Perfil or DEPENDENTE in Perfil or dependente in Perfil:
                        Perfil = '3'
                        worksheet.write(lPerfil, 7, Perfil)
                    elif pai in Perfil or mae in Perfil or PAI in Perfil or MAE in Perfil:
                        Perfil = '9'
                        worksheet.write(lPerfil, 7, Perfil)
                    elif sindico in Perfil or SINDICO in Perfil:
                        Perfil = '1'
                        worksheet.write(lPerfil,7, Perfil)

                    elif funcionario in Perfil or FUNCIONARIO in Perfil:
                        Perfil = '8'
                        worksheet.write(lPerfil,7, Perfil)
                    elif locatario in Perfil or LOCATARIO in Perfil:
                        Perfil = '23'
                        worksheet.write(lPerfil,7, Perfil)
                    elif Prestador in Perfil or PRESTADOR in Perfil:
                        Perfil = '5'
                        worksheet.write(lPerfil,7, Perfil)
                    elif Zelador in Perfil or ZELADOR in Perfil:
                        Perfil = '12'
                        worksheet.write(lPerfil,7, Perfil)

                    else:
                        Perfil = '6'
                        worksheet.write(lPerfil, 7, Perfil)
                    lPerfil = lPerfil + 1
                except:
                    pass
            print('Perfil OK')
        except:messagebox.showinfo("Erro", "Erro na coluna PERFIL, verifique o conteudo")
        #------------------ENTRADA------------------
        lEntrada = 1
        Entrada = 2
        worksheet.write(0, 8, "ENTRADA")
        for i in range(len(df_bl1.index)):
            worksheet.write(lEntrada, 8, Entrada)
            lEntrada = lEntrada + 1
        print('Entrada OK')
        #------------------OBS------------------
        try:
            lOBS = 1
            try:
                df_bl1['OBS'] = df_bl1['OBS'].fillna('null').map(str)
            except:
                pass
            worksheet.write(0, 9, "OBS")
            for i in range(len(df_bl1.index)):
                try:
                    OBS = df_bl1['OBS'].loc[i]
                except:
                    OBS = 'null'
                worksheet.write(lOBS, 9, OBS, format1)
                lOBS = lOBS + 1
            print('Obs OK')
        except:messagebox.showinfo("Erro", "Erro na coluna OBS, verifique o conteudo")
        #-----------------BLOCO-------------------
        try:
            linha = 1
            worksheet.write(0, 10, "BLOCO")
            df_bl1['BLOCO'] = df_bl1['BLOCO'].map(str)
            for i in range(len(df_bl1.index)):
                bloco = df_bl1['BLOCO'].loc[i]
                worksheet.write(linha, 10, bloco, format1)
                linha = linha + 1
            print('Bloco OK')
        except:messagebox.showinfo("Erro", "Erro na coluna BLOCO, verifique o conteudo")
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
        #------------------CARTAO------------------
        try:
            lCartao = 1
            worksheet.write(0, 13, "CARTAO")
            for i in range(len(df_bl1.index)):
                try:
                    try:
                        cartao = df_bl1['CARTAO'].loc[i]
                        #print(cartao)
                        worksheet.write(lCartao, 13, cartao, format1)
                    except:
                        cartao = 'null'
                        worksheet.write(lCartao, 13, cartao, format1)
                except:
                    cartao = 'null'
                    worksheet.write(lCartao, 13, cartao, format1)
                lCartao = lCartao + 1
            print('CARTAO OK')
        except:messagebox.showinfo("Erro", "Erro na coluna CARTAO, verifique o conteudo")
        #-----------------Unidade--------------------------------
        try:
            lUnidade = 1
            worksheet.write(0, 12, "UNIDADE")
            for i in range(len(df_bl1.index)):
                worksheet.write(lUnidade, 12, unidade, format1)
                lUnidade = lUnidade + 1
            print('Unidade OK')
                #-----------------Equipamentos-----------------
            lEquipamentos = 1
            worksheet.write(0, 14, "EQUIPAMENTOS")
            for i in range(len(df_bl1.index)):
                worksheet.write(lEquipamentos, 14, equipamento, format1)
                lEquipamentos = lEquipamentos + 1
                #   print(equipamento)
            print('Equipamento OK')
        except:messagebox.showinfo("Erro", "Erro na coluna UNIDADE, verifique o conteudo")
        messagebox.showinfo("Title", "Arquivo organizado\n Procure o arquivo CadRef")    
        workbook.close()
    except:
        messagebox.showinfo("Title", "Erro\n Verifique o arquivo Excel")