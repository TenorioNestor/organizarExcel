import testeArq
import pandas as pd
from tkinter import messagebox
import datetime
import http.client
import re
import requests



cadastroscomerro = []
erroendpoint = []
errojson = []
dataBl1 = open(r"C:\Users\Futturis-05\Documents\leitorPython\leitor\CadRef.xlsx",'rb')
df_bl1 = pd.read_excel(dataBl1)


for i in range(len(df_bl1.index)):
    conn = http.client.HTTPConnection("localhost:8080")
    lData = 1
    date = datetime.date(1930,1,1) #Define o formato da data
    try:
        aniversario = df_bl1['NASCIMENTO'].loc[i]
        aniversario = str(aniversario[:10])
    except:
        aniversario = date
    lData = lData + 1
    #print('Aniversario OK')
#----------CPF--------------------------
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
    
    #--------------nome-------------------
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
#-----------------RG---------------------

    lRg = 1
    try:
        df_bl1['RG'] = df_bl1['RG'].fillna('null').map(str)
    except:
        rg = 'null'
    try:
        rg = df_bl1['RG'].loc[i]
    except:
        rg = 'null'
    lRg = lRg + 1
    #print('RG OK')
#----------TELEFONE--------------------------

    lTelefone = 1
    try:
        df_bl1['TELEFONE'] = df_bl1['TELEFONE'].fillna(0)
    except:
        pass
    try:
        telefone = df_bl1['TELEFONE'].loc[i]
    except:
        telefone = 0
    lTelefone = lTelefone + 1
    #print('Telefone OK')
#------------------EMAIL------------------

    lEmail = 1
    contEmail = 0
    df_bl1['EMAIL'] = df_bl1['EMAIL'].fillna('null').map(str)
    try:
        email = df_bl1['EMAIL'].loc[i]  
        divisao = ';'
        if divisao in email:
            email = email[:divisao]
        else:
            email = email    
        lEmail = lEmail + 1
    except:
        email = df_bl1['EMAIL'].loc[i]  
        lEmail = lEmail + 1
    contEmail = contEmail +1
    #print('Email OK')
#------------------Celular------------------

    lCelular = 1
    try:
        df_bl1['CELULAR'] = df_bl1['CELULAR'].fillna(0)
    except:
        pass
    try:
        celular = df_bl1['CELULAR'].loc[i]
    except:
        celular = 0
    lCelular = lCelular + 1
    #print('Celular OK')
#--------------------------Perfil*-----------------------

    lPerfil = 1
    df_bl1['PERFIL'] = df_bl1['PERFIL'].fillna('2')
    try:
        Perfil = df_bl1['PERFIL'].loc[i]
        Prop = 'Propri' #2
        PROP = 'PROPR' #2
        MORADOR = 'MORADOR' #2
        morador = 'morador' #2
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
        Inquilino = 'inquilino'
        INQUILINO = 'INQUILINO'
        if Prop in Perfil or PROP in Perfil or Inquilino in Perfil or INQUILINO in Perfil or MORADOR in Perfil or morador in Perfil:
            Perfil = '2'
        elif filho in Perfil or conjugue in Perfil or FILHO in Perfil or ESPOS in Perfil or DEPENDENTE in Perfil or dependente in Perfil:
            Perfil = '3'
        elif pai in Perfil or mae in Perfil or PAI in Perfil or MAE in Perfil:
            Perfil = '9'
        elif sindico in Perfil or SINDICO in Perfil:
            Perfil = '1'
        elif funcionario in Perfil or FUNCIONARIO in Perfil:
            Perfil = '8'
        elif locatario in Perfil or LOCATARIO in Perfil:
            Perfil = '23'
        elif Prestador in Perfil or PRESTADOR in Perfil:
            Perfil = '5'
        elif Zelador in Perfil or ZELADOR in Perfil:
            Perfil = '12'
        else:
            Perfil = '6'
        lPerfil = lPerfil + 1
    except:
        pass
    #print('Perfil OK')
#------------------ENTRADA------------------
    lEntrada = 1
    Entrada = 2
    lEntrada = lEntrada + 1
    #print('Entrada OK')
#------------------OBS------------------

    lOBS = 1
    try:
        df_bl1['OBS'] = df_bl1['OBS'].fillna('null').map(str)
    except:
        pass
    try:
        OBS = df_bl1['OBS'].loc[i]
    except:
        OBS = 'null'
    lOBS = lOBS + 1
    #print('Obs OK')
#-----------------BLOCO-------------------

    linha = 1
    df_bl1['BLOCO'] = df_bl1['BLOCO'].map(str)
    bloco = df_bl1['BLOCO'].loc[i]
    linha = linha + 1
    #print('Bloco OK')
#-----------------AMBIENTE-----------------

    linha1 = 1
    df_bl1['AMBIENTE'] = df_bl1['AMBIENTE'].map(str)
    ambiente = df_bl1['AMBIENTE'].loc[i]
    linha1 = linha1 + 1
    #print('Ambiente OK')
#------------------CARTAO------------------

    lCartao = 1
    try:
        try:
            cartao = df_bl1['CARTAO'].loc[i]
            ##print(cartao)
        except:
            cartao = 'null'
    except:
        cartao = 'null'
    lCartao = lCartao + 1
    #print('CARTAO OK')
#-----------------Unidade--------------------------------

    unidade = 1000
    #print('Unidade OK')
        #-----------------Equipamentos-----------------
    lEquipamentos = 1
    equipamentos = "123,234"



    url = "http://localhost:8080/cadcorr/findnome"
    querystring = {"nome":{nome},"cpf":{cpf},"email":{email}}

    payload = ""
    headers = {"User-Agent": "insomnia/2023.5.8"}

    response = requests.request("GET", url, data=payload, headers=headers, params=querystring)

    #print(response.text)
    #print(response)
    data = response



    if "[]".encode('utf-8') in response.content:
        #print("Aqui não -------------------------------------------")
        conn = http.client.HTTPConnection("localhost:8080")
        payload = f"{{\n\t\"nascimento\":\"1912-12-12\",\n\t\"cpf\":\"{cpf}\",\n\t\"nome\":\"{nome}\",\n\t\"rg\":\"{rg}\",\n\t\"telefone\":\"{telefone}\",\n  \"email\":\"{email}\",\n  \"celular\":{celular},\n  \"perfil\":\"{Perfil}\",\n  \"obs\":\"{OBS}\",\n  \"bloco\":\"{bloco}\",\n  \"ambiente\":\"{ambiente}\",\n  \"cartao\":\"12345678\",\n  \"unidade\":{unidade}\n}}"
        payload = f"{{\n\t\"nascimento\":\"{aniversario}\",\n\t\"cpf\":\"{cpf}\",\n\t\"nome\":\"{nome}\",\n\t\"rg\":\"{rg}\",\n\t\"telefone\":\"{telefone}\",\n  \"email\":\"{email}\",\n  \"celular\":{celular},\n  \"perfil\":\"{Perfil}\",\n  \"obs\":\"{OBS}\",\n  \"{bloco}\":\"2\",\n  \"ambiente\":\"{ambiente}\",\n  \"cartao\":\"12345678\",\n  \"unidade\":{unidade}\n}}"
        #print(payload)
        headers = {
                'Content-Type': "application/json",
                'User-Agent': "insomnia/2023.5.8"
            }

        conn.request("POST", "/cadcorr", payload, headers)

        res = conn.getresponse()
        data = res.read()

        #print(data.decode("utf-8"))
        resposta = '"status":400'
        if resposta in data.decode("utf-8"):
            #print(f"nome:{nome}")
            itemJson = nome  # Get input from the user
            errojson.append(itemJson)
    elif response.status_code == 400:
        #print(f"Unexpected status code: {response.status_code}")
        #print(f"Nome do tantan:{nome}")
        itemEnd = nome  # Get input from the user
        erroendpoint.append(itemEnd)
    else:
        #print(f"Unexpected status code: {response.status_code}")
        #print(f"Nome do tantan:{nome}")
        item = nome  # Get input from the user
        cadastroscomerro.append(item)


print("Lista de cadastros repetido:", cadastroscomerro)
print("Lista de cadastros no envio:", erroendpoint)
print("Lista de cadastros Erro 400:", errojson)
messagebox.showinfo("Title", "Processo finalizado!")    
