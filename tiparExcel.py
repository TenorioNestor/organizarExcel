import xlsxwriter
import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

def tipagemDados():
    workbook = xlsxwriter.Workbook('Cadastros.xlsx')
    worksheet = workbook.add_worksheet()
    workbookList = xlsxwriter.Workbook('ListaRepetidos.xlsx')
    worksheetList = workbookList.add_worksheet()
    dataBl1 = open(r"C:\Users\Futturis-05\Documents\leitorPython\leitor\CadRef.xlsx",'rb')
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
    print('=============================================================')
    #--------LINHAS E COLUNAS -------------------
    format1 = workbook.add_format({'num_format': '@'})
    format2 = workbook.add_format({'num_format': 'dd/mm/yy'})
    format3 = workbook.add_format({'num_format': '0'})
    #----------ANIVERSARIO-------------------
    lData = 1
    worksheet.write(0, 0, 'NASCIMENTO')
    for i in range(len(df_bl1.index)):
        data = df_bl1['NASCIMENTO'].loc[i]
        print(data)
        worksheet.write(lData, 0, data,format2 )
        lData = lData + 1
    print('Aniversario OK') 
    #----------CPF--------------------------
    lCpf = 1
    df_bl1['CPF'] = df_bl1['CPF'].fillna(0)
    worksheet.write(0, 1, 'CPF')
    for i in range(len(df_bl1.index)):
        cpf = df_bl1['CPF'].loc[i]
        worksheet.write(lCpf, 1, cpf, format3)
        lCpf = lCpf + 1
    print('CPF OK')
    #------------------NOME------------------
    lNome = 1
    lnomeRep = 1
    worksheet.write(0, 2, 'Nome')
    worksheetList.write(0,1,'Nome')
    for i in range(len(df_bl1.index)):
        j = i-1 # 0
        nome = df_bl1['NOME'].loc[i] #1
        worksheet.write(lNome, 2, nome, format1)
        maxList = len(df_bl1.index)
        if i>0 and i< maxList:
            nomel = df_bl1['NOME'].loc[j]
            if nomel == nome:
                worksheetList.write(lnomeRep, 1, nomel, format1)
                lnomeRep = lnomeRep + 1
        lNome = lNome + 1   
    print('Nome OK')
    #-----------------RG---------------------
    lRg = 1
    df_bl1['RG'] = df_bl1['RG'].fillna('null').map(str)
    worksheet.write(0, 3, "RG")
    for i in range(len(df_bl1.index)):
        rg = df_bl1['RG'].loc[i]
        worksheet.write(lRg, 3, rg, format1)
        lRg = lRg + 1
    print('RG OK')
    #----------TELEFONE--------------------------
    lTelefone = 1
    df_bl1['TELEFONE'] = df_bl1['TELEFONE'].fillna(0)
    worksheet.write(0, 4, "Telefone")
    for i in range(len(df_bl1.index)):
        telefone = df_bl1['TELEFONE'].loc[i]
        worksheet.write(lTelefone, 4, telefone, format3)
        lTelefone = lTelefone + 1
    print('Telefone OK')
    #------------------EMAIL------------------
    lEmail = 1
    df_bl1['EMAIL'] = df_bl1['EMAIL'].fillna('null').map(str)
    worksheet.write(0, 5, "Email")
    worksheetList.write(0,2,'Email')
    for i in range(len(df_bl1.index)):
        email = df_bl1['EMAIL'].loc[i]
        worksheet.write(lEmail, 5, email, format1)
        lEmail = lEmail + 1
    print('Email OK')
    #------------------Celular------------------
    lCelular = 1
    df_bl1['CELULAR'] = df_bl1['CELULAR'].fillna(0)
    worksheet.write(0, 6, "CELULAR")
    for i in range(len(df_bl1.index)):
        celular = df_bl1['CELULAR'].loc[i]
        worksheet.write(lCelular, 6, celular, format3)
        lCelular = lCelular + 1
    print('Celular OK')
    #------------------Perfil------------------
    lPerfil = 1
    df_bl1['PERFIL'] = df_bl1['PERFIL'].fillna(2)
    worksheet.write(0, 7, "PERFIL")
    for i in range(len(df_bl1.index)):
        Perfil = df_bl1['PERFIL'].loc[i]
        worksheet.write(lPerfil, 7, Perfil, format3)
        lPerfil = lPerfil + 1
    print('Perfil OK')
    #------------------ENTRADA------------------
    lEntrada = 1
    df_bl1['ENTRADA'] = df_bl1['ENTRADA'].fillna(2)
    worksheet.write(0, 8, "Entrada")
    for i in range(len(df_bl1.index)):
        Entrada = df_bl1['ENTRADA'].loc[i]
        worksheet.write(lEntrada, 8, Entrada, format3)
        lEntrada = lEntrada + 1
    print('Entrada OK')
    #------------------OBS------------------
    lOBS = 1
    df_bl1['OBS'] = df_bl1['OBS'].fillna('null').map(str)
    worksheet.write(0, 9, "OBS")
    for i in range(len(df_bl1.index)):
        OBS = df_bl1['OBS'].loc[i]
        worksheet.write(lOBS, 9, OBS, format1)
        lOBS = lOBS + 1
    print('Obs OK')
    #-----------------BLOCO-------------------
    linha = 1
    worksheet.write(0, 10, "Bloco")
    df_bl1['BLOCO'] = df_bl1['BLOCO'].map(str)
    for i in range(len(df_bl1.index)):
        bloco = df_bl1['BLOCO'].loc[i]
        worksheet.write(linha, 10, bloco, format1)
        linha = linha + 1
    print('Bloco OK')
    #-----------------AMBIENTE-----------------
    linha1 = 1
    worksheet.write(0, 11, "AMBIENTE")
    df_bl1['AMBIENTE'] = df_bl1['AMBIENTE'].map(str)
    for i in range(len(df_bl1.index)):
        AMBIENTE = df_bl1['AMBIENTE'].loc[i]
        worksheet.write(linha1, 11, AMBIENTE, format1)
        linha1 = linha1 + 1
    print('Ambiente OK')
    #------------------Unidade------------------
    lUnidade = 1
    df_bl1['UNIDADE'] = df_bl1['UNIDADE']
    worksheet.write(0, 12, "Unidade")
    for i in range(len(df_bl1.index)):
        Unidade = df_bl1['UNIDADE'].loc[i]
        worksheet.write(lUnidade, 12, Unidade, format3)
        lUnidade = lUnidade + 1
    print('Unidade OK')
    #------------------CARTAO------------------
    lCartao = 1
    df_bl1['CARTAO'] = df_bl1['CARTAO'].fillna('null')
    worksheet.write(0, 13, "CARTAO")
    for i in range(len(df_bl1.index)):
        Cartao = df_bl1['CARTAO'].loc[i]
        worksheet.write(lCartao, 13, Cartao, format1)
        lCartao = lCartao + 1
    print('CARTAO OK')
    #-----------------Equipamentos-----------------
    lEquipamentos = 1
    worksheet.write(0, 14, "Equipamentos")
    for i in range(len(df_bl1.index)):
        Equipamentos = df_bl1['EQUIPAMENTOS'].loc[i]
        worksheet.write(lEquipamentos, 14, Equipamentos, format1)
        lEquipamentos = lEquipamentos + 1
    print('Equipamento OK')
    workbookList.close()
    workbook.close()
    messagebox.showinfo("Title", "Arquivo tipado, o nome do arquivo é Cadastros e esta na mesma pasta deste software!")