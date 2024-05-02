import pandas as pd
from tkinter import messagebox
import re



cpferro = []
nomeerro = []

dataBl1 = open(r"C:\Users\Futturis-05\Documents\leitorPython\leitor\CadRef.xlsx",'rb')
df_bl1 = pd.read_excel(dataBl1)


for i in range(len(df_bl1.index)):
    df_bl1['NOME'] = df_bl1['NOME'].str.replace(",", ";")
    df_bl1['NOME'] = df_bl1['NOME'].str.replace("/", ";")
    df_bl1['NOME'] = df_bl1['NOME'].str.replace(" e ", ";")
    df_bl1['NOME'] = df_bl1['NOME'].str.replace(" E ", ";")
    lNome = 1
    contNome = 0
    try:
        nome = df_bl1['NOME'].loc[i].rstrip()
        nome = nome.lstrip()
        divisao = ';'
        if divisao in nome:
            nome = nome[:divisao]
        else:
            nome = nome  
        lNome = lNome + 1 
    except:
        nome = df_bl1['NOME'].loc[i]
        lNome = lNome + 1
    contNome = contNome + 1  
    #print('Nome OK')
    lCpf = 1
    try:
        df_bl1['CPF'] = df_bl1['CPF'].fillna(0)
    except:
        pass
    try:
        cpf = df_bl1['CPF'].loc[i]
    except:
        cpf = 0
    lCpf = lCpf + 1
    #print('CPF OK')
    
    def validate_cpf(cpf):

        # Remove any non-digit characters
        cpf = cpf.replace('-', '').replace('.', '')

        # Check if the length is correct (11 digits)
        if len(cpf) != 11:
            itemnome = nome
            nomeerro.append(itemnome)
            itemcpf = cpf  # Get input from the user
            cpferro.append(itemcpf)
            return False

        # Convert the CPF string to a list of integers
        cpf_digits = [int(x) for x in cpf]

        # Calculate the first verification digit (DV1)
        sum_dv1 = 0
        for i in range(0, 9):
            sum_dv1 += (10 - i) * cpf_digits[i]

        dv1 = 11 - (sum_dv1 % 11)
        if dv1 == 10 or dv1 == 0:
            dv1 = 0

        # Calculate the second verification digit (DV2)
        sum_dv2 = 0
        for i in range(0, 10):
            sum_dv2 += (11 - i) * cpf_digits[i]

        dv2 = 11 - (sum_dv2 % 11)
        if dv2 == 10 or dv2 == 0:
            dv2 = 0

        # Check if the calculated DV1 and DV2 match the digits in the CPF
        if dv1 != cpf_digits[9] or dv2 != cpf_digits[10]:
            itemnome = nome
            nomeerro.append(itemnome)
            itemcpf = cpf  # Get input from the user
            cpferro.append(itemcpf)
            return False

        return True

    # Example usage
    cpf_to_validate = str(cpf)
    is_valid = validate_cpf(cpf_to_validate)
    print(f"Nome:{nome} ;{cpf_to_validate}; valid: {is_valid}")

    

print("Lista de erro de cpf:", cpferro)

messagebox.showinfo("Title", "Processo finalizado!")    
