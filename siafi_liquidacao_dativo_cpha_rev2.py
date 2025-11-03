# -*- coding: utf-8 -*-
"""
Created on Mon Dec 12 11:24:16 2022

@author: m1379117
release 11
"""



import pyautogui as pi
import time
import subprocess
import os #, sys
import asyncio
#from asgiref.sync import sync_to_async
import pyperclip
#import time
import re
import pandas as pd
import datetime, calendar
import sys


import pygetwindow as gw
tit='pw3270'

def active_maximize(tit):
    print("INICIO da funcao active_maximize")
    time.sleep(1)
    x=gw.getAllTitles()
    print(x)
    if tit in x:
        print(f'OK - Localizada janela com titulo {tit} do SAIFI')
        janela=gw.getWindowsWithTitle(tit)[0]
        janela.activate()
        #time.sleep(2)
        janela.maximize()
        #time.sleep(0.5)
        return "OK"
    else:
        print(f'Nao localizada janela com titulo {tit}')
        return "NOK"



def abre_siafi_run(local):
    
    print("INICIO da funcao  abre_siafi_run Liquidacao")
    
    #import sys
    print(os.getcwd())
    
    
    
    if local=="itautec":
        dd=r"C:\Program Files (x86)\pw3270"
        #"C:\Program Files (x86)\pw3270\pw3270.exe" --host=bhmvsb.prodemge.gov.br
        #b=[r"C:\Program Files (x86)\pw3270\pw3270.exe"]#, "--host=bhmvsb.prodemge.gov.br"]
        b=r'"C:\Program Files (x86)\pw3270\pw3270.exe" --host=bhmvsb.prodemge.gov.br'
        
    if local=="accer_1":
        dd=r"C:\Program Files\pw3270"
        #b=[r"C:\Program Files\pw3270\pw3270.exe"]# "--host=bhmvsb.prodemge.gov.br"]
        b=r'"C:\Program Files (x86)\pw3270\pw3270.exe" --host=bhmvsb.prodemge.gov.br'
        
    if local=="robo-server":
        print("To be installed")
        
    os.chdir(dd)
   
    print(os.getcwd())
   
    subprocess.run(b)
    
                
    msg= "OK from abre_siafi_run"
    print(msg)
    
    return

def read_excel(file, tipo):
    
    # Read excel and returns a data frame)
    # tipo =1 planilha dativos
    # tipo =  2  planilha gerais
    
    df=pd.read_excel(file)
    
    column_names = list(df.columns)
    
    print(column_names)
    
    print(type(column_names))
    
    
    if tipo=="1": # Planilha Dativos
        #returns only the columns index 3 and 4
        #df2 = df.iloc[:,[3,4]] # Select columns by Index 3 and 4
        df2 = df.iloc[:,[4,18]] # Select columns by Index 3 and 18 = cpf e valo liquido na planilha dativos
        
    if tipo=="2": # Planilha Gerais
        # Returns all excel df
        #df2=df
        print("TBD")
        
    # Adiciona uma coluna vazia
        
    print(df2)
    return df2

   

def cria_df3(column_names):
     # retorna df vazio apenas com o nome das colunas a serem utilizadas
    
       
    print(column_names)
    
    df1 = pd.DataFrame(columns = column_names)
    
    return df1





def loga_siafi(login, senha, ue, local):
    print("INICIO da funcao loga_siafi")
    
    msg1="Logon executado com sucesso"
    msg2="Unidade Executora:"
    
    import pyautogui as pi
    import time
    
    
    for i in range(3):
        print(i)
        rr=active_maximize(tit)
        if rr=="OK":
            break
    
    
    #pi.getWindowsWithTitle(resp)[0].maximize()
    time.sleep(2)
    print("1")
    
    pi.hotkey('enter')
    time.sleep(2)
    
    """
    #if in english language
    if (local=="accer_1" or local=="itautec"):
        pi.hotkey('alt','n')
    #if local=="itautec":
        #pi.hotkey('alt', 'r')
    if local=="robo-server":
        print("TBD")
    
    time.sleep(1)
    print("2")
    pi.hotkey('c') # Conectar
    time.sleep(2)
    """
   
    #for i in range(25):
        #pi.hotkey('right')
        #time.sleep(0.1)
    
    print("3")
    msg1="Aplicacao:"
    st=1
    n=5
    rsp=check_current_screen(msg1, n, st)
    if rsp=="OK":
        print("Tela de autenticação aberta com sucesso")
        
    else:
        print("Tela de autenticação indisponivel")
        resp="NOK-0 from loga_siafi autenticacao"
        pi.hotkey('esc')
        time.sleep(1)
        pi.hotkey('enter')
        rsp=check_current_screen(msg1, n, st)
        if rsp=="OK":
            print("Tela de autenticação aberta com sucesso")
        else:
            print("Tela de autenticação indisponivel")
            resp="NOK-0 from loga_siafi autenticacao"
            print(resp)     
            return resp
    
    pi.write('simg')
    pi.hotkey('tab')
    pi.write(login)
    pi.write(senha)
    pi.hotkey('enter')
    time.sleep(1)
    
    msg1="Logon executado com sucesso"
    for i in range(10):
        time.sleep(2)
        pi.hotkey('ctrl' , 'a')
        time.sleep(1)
        pi.hotkey('alt' , 'e')
        time.sleep(1)
        pi.hotkey('enter')
        time.sleep(1)
        #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        time.sleep(1)
        #print(x)
        pos1=x.find(msg1)
        
        #print(pos1)
        #ops=pi.confirm(text=f'Verifique se localizou {msg1}...\npos1 = {pos1}', title='liquidacao', buttons=['OK', 'Cancel'])
        #if ops=='Cancel':
            #sys.exit()
        
        if pos1>=0:
            break
        
        pi.hotkey('enter')
        
    
    time.sleep(2)
    pi.write('simg')
    time.sleep(0.5)
    pi.hotkey('enter') 
    time.sleep(1)
    
    a,b=pi.size()
    #pi.click(x=a/2, y=b/2)
    #time.sleep(1)
    
    for i in range(5):
        time.sleep(0.5)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl' , 'a')
        time.sleep(0.5)
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg2)
        if pos1 >=0:
            break
        pi.hotkey('enter')
        
     
    #pi.hotkey('enter')
    pi.hotkey('esc')
    pi.hotkey('tab')
    
    #time.sleep(1)
    pi.write(ue)#uidade executora
    #time.sleep(1)
    
    """
    from datetime import datetime
    now = datetime.now()
    ano = now.strftime("%Y")
    print(ano)
    pi.write(ano)#ano
    """    
    
    
    #time.sleep(1)
    pi.hotkey('enter')
    time.sleep(2)
    
    msg1="01-Rotina Administrativa"
    n=10
    st=1
    rsp=check_current_screen(msg1, n, st)
    if rsp=="NOK":
        resp="NOK-1 from loga_siafi"
        pi.hotkey('esc')
        time.sleep(2)
        rsp=check_current_screen(msg1, n, st)
        if rsp=="NOK":
            resp="NOK-2 from loga_siafi"
        resp="OK-2 from loga_siafi"
    else:
        resp="OK-1 from loga_siafi-1"

    print(resp)
    
    return resp

def get_first_line(qte, n, st):
   """get n first character in the currente screen"""
    # qte = characters qte to get : int
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
    # Click no centro da tela
   a,b=pi.size()
   pi.click(x=a/2, y=b/2)
   #pi.hotkey('tab')#Opção
      
   for i in range(n):
       time.sleep(st)
       pi.click(x=a/2, y=b/2)
       pi.hotkey('ctrl' , 'a')
       pi.hotkey('alt' , 'e')
       pi.hotkey('enter')
       #pyperclip.copy('The text to be copied to the clipboard.')
       x = pyperclip.paste()
       #print(x)
    
   x=x.strip()
   line=x[:qte]
   
    
   return line 


def check_current_screen(msg1, n, st):
    """try to find msg in the currente screen"""
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
     # Click no centro da tela
    a,b=pi.size()
    pi.click(x=a/2, y=b/2)
    #pi.hotkey('tab')#Opção
    
    
    
    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl' , 'a')
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg1)
        if pos1 >=0:
            print(f' localizado {msg1}')
            resp="OK"
            break
        else:
            resp="NOK"
            print(f' ciclo {i+1} não localizado {msg1}')
        
        i+=1
        if i+1==n:
            resp="NOK"
    
    pi.hotkey('tab')
    print(resp)
    return resp

def retorna_tela_inic_empenho():
    # Retorna telas anteriores clicando duas vezes em PF3
    
    pf=ctl_keys_itautec_17(3)
    # Coordenadas PF3
    x1=pf[0]
    y1=pf[1]
    x1=int(x1)
    y1=int(y1)
    
    # Clica 2 x em PF3 para Voltar
    pi.click(x=x1, y=y1)
    time.sleep(1)
    pi.click(x=x1, y=y1)
    time.sleep(1)
    
    return 


def quit_siafi(local):
    
     print("INICIO da funcao quit_siafi")
     
     import pyautogui as pi
     import time
     
     #if in english
     if (local=="accer_1" or local=="itautec"):
        pi.hotkey('alt','f')
        time.sleep(1)
        pi.hotkey('q')
          
     return
 
async def controle_siafi(login, senha, ue, local,ano):
    print("INICIO da funcao  abre_siafi_async_run")
    
    import asyncio
    import sys
    import time
    
    from asgiref.sync import sync_to_async
    loop = asyncio.get_event_loop()
        
    async_function = sync_to_async(abre_siafi_run, thread_sensitive=False)
    task1 = asyncio.create_task(async_function(local))
    print(f'task1 ====> {task1}')
    
    
    #time.sleep(1)
    
    
    async_function = sync_to_async(loga_siafi, thread_sensitive=False)
    task2=loop.create_task(async_function(login, senha, ue, local))
    print(f'task2 ====> {task2}')
        
    await task2
    
    
    
    ##############################################
    ##### Aqui deverá ter o loop da planilha #####
    # async_function = sync_to_async(controle_empenho, thread_sensitive=False)
    # task3=loop.create_task(async_function(file))
    # print(f'task3 ====> {task3}')
    # await task3
    #########################################
    
    if local=="accer_1":
        resp="pw3270:A - tn3270://bhmvsb.prodemge.gov.br:23"
    
    if local=="itautec":
        resp="pw3270 - bhmvsb.prodemge.gov.br"
    
    print(resp)
    
    
    return resp


def get_val_anula(n, st):
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
    print("INICIO da funcao get_val_anula")
    
    msg1="Anular Saldo Liquidado/Retencao/Multa Contratual"
    
    # Click no centro da tela
    a,b=pi.size()
    pi.click(x=a/2, y=b/2)
    #pi.hotkey('tab')#Opção
    
    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl' , 'a')
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg1)
        if pos1 >=0:
            print(f' localizado {msg1}')
            resp="OK"
            break
        else:
            resp="NOK"
            print(f' ciclo {i+1} não localizado {msg1}')
            return resp
            
        i+=1
        if i+1==n:
            resp="NOK"
            return resp
    
    
    #rp=re.findall("\d+[,]\d+", x)
    rp=re.findall("[\d+.]*[,]\d+", x)
    resp=rp[0]
    
    
    resp=re.sub('[^0-9,]', '', resp)
   
        
    print(f' Valor = {resp}')
    
    #ops=pi.confirm(text=f'Verifique a valor a ser anulado = {resp}', title='', buttons=['OK', 'Cancel'])
    #if ops=='Cancel':
        #sys.exit()
    resp=trata_valor_4(resp)
    
    return resp



    

def get_val_retencao(n, st):
    
    print("INICIO da funcao get_val_retencao")
    
    
   
    """ Busca valor retencao na tela"""
    
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
         #"RETENCAO DE CONTRIBUICOES PREVIDENCIARIAS"
    msg1="RETENCAO DE CONTRIBUICOES PREVIDENCIARIAS"
    
    # Click no centro da tela
    a,b=pi.size()
    pi.click(x=a/2, y=b/2)
    #pi.hotkey('tab')#Opção
    
    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl' , 'a')
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg1)
        if pos1 >=0:
            print(f' localizado {msg1}')
            resp="OK"
            break
        else:
            resp="NOK"
            print(f' ciclo {i+1} não localizado {msg1}')
            return resp
            
        i+=1
        if i+1==n:
            resp="NOK"
            return resp
    
    
    rp=re.findall("[\d+.]*[,]\d+", x)
    resp=rp[0]
    
    
    resp=re.sub('[^0-9,]', '', resp)
   
        
    print(f' Valor = {resp}')
    
    #ops=pi.confirm(text=f'Verifique a valor a ser anulado = {resp}', title='', buttons=['OK', 'Cancel'])
    #if ops=='Cancel':
        #sys.exit()
    resp=trata_valor_4(resp)
    
    return resp


def get_emp_number(n, st):
    
    """ Busca numero do empenho na tela"""
    
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
    
    msg1="REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL"
    
    # Click no centro da tela
    a,b=pi.size()
    pi.click(x=a/2, y=b/2)
    #pi.hotkey('tab')#Opção
    
    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl' , 'a')
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg1)
        if pos1 >=0:
            print(f' localizado {msg1}')
            resp="OK"
            break
        else:
            resp="NOK"
            print(f' ciclo {i+1} não localizado {msg1}')
            return resp
            
        i+=1
        if i+1==n:
            resp="NOK"
            return resp
        
        
    t1="Nr. Documento:"#    6386          
    t2="Evento"
    
    pos1=x.find(t1)
    pos2=x.find(t2)
    
    emp=x[pos1+14:pos2]
    emp=emp.strip()
    
    resp=emp
    
    print(f'Numero do Empenho = {resp}')
    return resp




def get_num_anula(n, st):
         
        
    """ Busca numero de anulacao"""
    
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
    
    msg1="ANULACAO EFETUADA"
    
    # Click no centro da tela
    a,b=pi.size()
    pi.click(x=a/2, y=b/2)
    #pi.hotkey('tab')#Opção
    

    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl' , 'a')
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg1)
        if pos1 >=0:
            print(f' localizado {msg1}')
            resp="OK"
            break
        else:
            resp="NOK"
            print(f' ciclo {i+1} não localizado {msg1}')
            return resp
            
        i+=1
        if i+1==n:
            resp="NOK"
            return resp
        
        
      
    #                          Nr.Anulacao: 0000087                              
    # Unid. Executora:
        
        
    t1="Nr.Anulacao:"#    6386          
    t2="Unid. Executora:"
    
    pos1=x.find(t1)
    pos2=x.find(t2)
    
    anul=x[pos1+len(t1):pos2]
    anul=anul.strip()
    
    resp=anul
    
    print(f'Numero do Empenho = {resp}')
    return resp
        




def ajusta(valor):
    print("INICIO da funcao trata_valor")

    try:
        valor = float(valor)
    except:
        valor = valor.replace(",", ".")
        valor = float(valor)

    valor = round(valor, 2)
    valor = str(valor)
    valor = valor.strip()

    # valor=str(valor)
    print(valor)
    pos21 = valor.find(".")
    pos22 = valor.find(",")
    print(pos21)
    print(pos22)

    if pos21 < 0 and pos22 < 0:
        valor = f'{valor}00'

    print(len(valor))
    ll1 = len(valor)-pos21
    if ll1 == 2:
        valor = f'{valor}0'

    print(valor)
    valor = valor.replace(".", "")
    valor = valor.replace(",", "")
    valor = valor.strip()

    print(f'valor corrigido = {valor}')

    return valor




def trata_sei_cpha(sei):
    
    print("INICIO da funcao trata_sei")
    
    sei=str(sei)
    sei=sei.strip()
    sei=sei.replace("'","")
    
    pos1=sei.find(".")
    if pos1>=0:
        sei=sei[:pos1]
        sei=sei.strip()
        
    sei1=re.sub("\D+", "",sei)
    x=len(sei1)
    if sei1[:4]!="1080":#cpha
        if x<20:
            f=20-x
            for j in range(f):
               sei=f'0{sei}'
               print(sei)
    
    print(f'sei corrigido ==> {sei}')
    
    
    return sei

def trata_cpf_cnpj(cpf_cnpj):
    print("INICIO da funcao trata_cpf_cnpj")
    
    cpf_cnpj=str(cpf_cnpj)
    cpf_cnpj= cpf_cnpj.strip()
    cpf_cnpj= cpf_cnpj.replace("'","")
    
    tipo="cpf"
    if "/" in cpf_cnpj: #cnpj
        tipo="cnpj"
        print(tipo)
        cpf_cnpj=re.sub("\D+", "",cpf_cnpj)
        print(cpf_cnpj)
        x=len(cpf_cnpj)
        
        # Completa ate atingir 11 digitos para CNPJ
        if x<14:
           f=14-x
           for j in range(f):
               cpf_cnpj=f'0{cpf_cnpj}'
               print(cpf_cnpj)
        
        
    else: # CPF
        print(tipo)
        cpf_cnpj=re.sub("\D+", "",cpf_cnpj)
        print(cpf_cnpj)
        x=len(cpf_cnpj)
        
       # Completa ate atingir 11 digitos para CPF
        if x<11:
           f=11-x
           for j in range(f):
               cpf_cnpj=f'0{cpf_cnpj}'
               print(cpf_cnpj)
    
    print(f'cpf_cnpj corrigido ==> {cpf_cnpj}')
    
    
    return cpf_cnpj

def trata_processo(processo):
    print("INICIO da funcao trata_processo")
    
    processo=str(processo)
    processo=processo.strip()
    processo=processo.replace("'","")
    
    pos1=processo.find(".")
    if pos1>=0:
        processo=processo[:pos1]
        processo=processo.strip()
    
        
    processo1=re.sub("\D+", "",processo)
    x=len(processo1)
    if x<20:
        f=20-x
        for j in range(f):
           processo=f'0{processo}'
           print(processo)
    
    print(f'processo corrigido ==> {processo}')
    
    
    return processo

def trata_valor(valor):
    
    print("INICIO da funca trata_valor")
    
    try:
        valor=float(valor)
    except:
        valor=valor.replace(",",".")
        valor=float(valor)
    valor = round(valor,2)
    valor=str(valor)
    valor=valor.strip()
    
     #valor=str(valor)
    print(valor)
    pos21=valor.find(".")
    pos22=valor.find(",")
    print(pos21)
    print(pos22)
    
    if pos21<0 and pos22<0:
        valor=f'{valor}00'
        
    print(len(valor))
    ll1=len(valor)-pos21
    if ll1==2:
       valor=f'{valor}0'
        
    print(valor)    
    valor=valor.replace(".","")
    valor=valor.replace(",","")
    valor=valor.strip()
    
    print(f'valor corrigido = {valor}')
        
    return valor

def trata_valor_2(vv):
    vv=str(vv)
    print(vv)

    pos21=vv.find('.')
    pos22=vv.find(',')
    
    if pos21<0 and pos22<0: # 0 0 
        print('00')
        print(len(vv), pos21, pos22)
        if len(vv)<=2:
            vv=vv+'00'
        
        return vv
    
    if pos21>0 and pos22<0: # 1 0
        print('10')
        vv=vv.replace(".", "")
        print(len(vv), pos21, pos22)
        print(vv)
        if len(vv)<=3:
            if len(vv)-pos21==1:
                vv='00'+vv
            if len(vv)-pos21==2:
                vv='0'+vv

        
        return vv
    
    if pos21<0 and pos22>0: # 0 1
        print('01')
        vv=vv.replace(",", "")
        print(vv)
        print(len(vv), pos22, pos21)
        if len(vv)<=3:
            if len(vv)-pos22==1:
                vv='00'+vv
            if len(vv)-pos22==2:
                vv='0'+vv
            #if len(vv)-pos22==3:
                #vv='0'+vv
         
        return vv
     
    
    if pos21>0 and pos22>0: # 1 1
        print('11')
        print(len(vv), pos22, pos21)
        if len(vv)<=4:
            if len(vv)-pos22==2:
                vv='00'+vv
            if len(vv)-pos22==3:
                vv='0'+vv
        vv=vv.replace(",", "")       
        vv=vv.replace(".", "")
        
        return vv


def trata_valor_4(vv):
    print("INICIO da funcao trata_valor_4")
    vv1=vv
    vv=str(vv)
    pos21=vv.find('.')
    pos22=vv.find(',')
    if pos21>0 and pos22>0: # 1 1
        print('11')
        print(len(vv), pos22, pos21)
        if len(vv)<=4:
            if len(vv)-pos22==2:
                vv='00'+vv
            if len(vv)-pos22==3:
                vv='0'+vv
        vv=vv.replace(",", "")       
        vv=vv.replace(".", "")
        resp=vv
    else:
        try:
            vv=vv.replace(",",".")
        except:
            pass
        vv=float(vv)
        resp="%0.2f" % (vv,)
        resp=resp.replace(".","")
    print(f'{vv1} ==> {resp}')
    
    return resp

def liquida_dativo(emp,uo,tipo1,valor,od,inss):
#def liquida_dativo(cpf_cnpj, valor, uo, od, p_trab, nat_desp, item, fonte, tipo, sei):
    print("==================================================")
    print("INICIO da funcao liquida_dativo")
    # emp=numero do empenho
    # uo = unidade orcamentaria =  "1081"
    # tipo1 = "inicial" ou "não inicial"
    # valor = valor tipo 2000 equivale a 20 reais = confirmado OK
    # od = ordenador de despesa = "10606119"
    # valor do desconto inss
     
    
    pi.PAUSE=0.4
    
    
    pi.write(uo)
    pi.write(emp)
    
    
    #ev=pi.confirm(text='Verifique se Uo e Empenho foram inseridos corretamente', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    pi.hotkey('enter')
    time.sleep(1)
    
    
    msg="VALORES PARA LIQUIDACAO INEXISTENTE"
    st=0.5
    n=1
    rsp=check_current_screen(msg, n, st)
    if rsp=="OK":
        
        pi.hotkey('esc')
        #pi.hotkey('up')
        resp=msg
        print(resp)
        #pi.PAUSE=0.3
        return resp
    
    
    #ev=pi.confirm(text=f'Proximas teclas previstas S e N', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    inss=ajusta(inss)
    
    #ev=pi.confirm(text=f'Verifique o valor de inss = {inss}\nSe igual a zero, nao precisa fazer retencao\nRefatorar Linha 895 do prgrama de Liquidacao.', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
    # RETENCAO PREVIDENCIARIA
    if inss=="000" or inss=="00" or inss=="0" or inss==None:
        pi.write("N")
    else:
        pi.write('S')
    
    # MULTA COINTRAUAL
    pi.write('N')
    
    
    # TRATA DATAS
    x = datetime.datetime.now()
    print(x)
    dia=x.strftime("%x")
    print(f'today ==> {dia}')
    yy=x.year
    mm=x.month
    #dd=x.day
    
    # MES CORRENTE
    res = calendar.monthrange(yy, mm)
    dd1 = res[1]
    print(dd1)
    
    dd1=str(dd1)
    mm=str(mm)
    if len(mm)<2:
        mm=f'0{mm}'
    yy=str(yy)
    
    
    #ev=pi.confirm(text=f'Verifique dia  mes e ano:\n dia={dd1}, mes={mm} , ano = {yy}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    pi.write(dd1)
    pi.write(mm)
    pi.write(yy)
    
    pi.write(mm)
    
    #ev=pi.confirm(text='Verifique se dia  mes e ano inseridos corretamente', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    # VALOR
    pi.write(valor)
    pi.hotkey('tab')
    
    #ORDENADOR DE DESPESA
    pi.write(od)
    pi.hotkey('enter')
    
    #time.sleep(0.5)
    pi.hotkey('esc')
    line=get_first_line(65, 1, 1)
    print(f' 1a linha da tela ==> {line}')
    #pi.hotkey('esc')
    #pi.hotkey('home')
    
    #ev=pi.confirm(text=f'Verifique 1a linha capturada\n{line}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    msg16='VALOR DA LIQUIDACAO MAIOR QUE SALDO DE EMPENHO'
    if  msg16 in line:
        resp=f'LIQUIDACAO NAO REALIZADA, {msg16}, EMPENHO-{emp}'
        pi.hotkey('f3')
        return resp
    
    
        
    if inss!="000" and inss!="00" and inss!="0" and inss!=None:
        # TELINHA EVENTOS DE RETENCAO
        time.sleep(1)
               
        pi.hotkey('esc')
        pi.hotkey('home')
        
        #ev=pi.confirm(text=f'Verifique proxima acoes\nPrevisto down e 01', title='', buttons=['OK', 'Cancel'])
        #if ev=='Cancel':
            #sys.exit()
        
        pi.hotkey('down')
        pi.write('01')
        pi.hotkey('f5')
                
        # TELINHA detalhamento da liquidação RETENCAO
        pi.write('29979036000140')  ## CNPJ do inss
        #time.sleep(0.5)
        pi.write(mm)
        #time.sleep(0.5)
        pi.write(yy)
        #time.sleep(0.5)
        pi.write(inss)
        #time.sleep(0.5)
        
        
        pi.hotkey('f5')
        #time.sleep(1)
    
    
        n=1
        st=0.5    
        msg="CREDOR INEXISTENTE"
        rsp=check_current_screen(msg, n, st)
        if rsp=="OK":
            resp=f'LIQUIDACAO NAO REALIZADA,CREDOR DE RETENCAO INEXISTENTE, EMPENHO-{emp}'
            print(resp)
            pi.hotkey('f2')
            time.sleep(0.5)
            pi.hotkey('f3')
            time.sleep(0.5)
            #pi.PAUSE=0.3
            return resp
        
        # TELINHA HISTORICO DE LIQUIDACAO
        pi.hotkey('up')
        pi.hotkey('up')
     
     
    time.sleep(0.5)    
    pi.hotkey('home')

    #ev=pi.confirm(text='Verifique se o cursor esta na posição correta p inserir mensagem no histórico', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    time.sleep(0.5)    
    pi.write(f'VALOR LIQUIDADO CONFORME EMPENHO NUMERO {emp}')
    time.sleep(2)
    pi.hotkey('f5')
    
    
    # TECLE PF5 PARA CONFIRMAR OU PF2 PARA ANULAR 
    pi.hotkey('f5')
    
   
    #ev=pi.confirm(text=f'Registro Efetuado com suceso, Verifique numero do empenho\nProxima tecla prevista é F5', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    #REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL 
    msg1="REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL"
    n=2
    st=0.5
    rp=check_current_screen(msg1, n, st)
    
    if rp=="OK":
        resp=f'LIQUIDACAO REALIZADA COM SUCESSO, EMPENHO-{emp}'
    if rp=="NOK":
        resp=f'LIQUIDACAO NAO REALIZADA, EMPENHO-{emp}'
        
    
    #  REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL
    
    print(resp)
    
    #ev=pi.confirm(text=f'Registro Efetuado com suceso, Verifique numero do empenho\nProxima tecla prevista é F5', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    pi.hotkey('enter')
    
    #pi.PAUSE=0.3
    
    return resp

    
    
    
def anula_liquidacao(emp, uo, valor, od, tipo):
    
    print("INICIO da funcao Anula_Liquidacao")
                        
       
    # Telinha  Ano Exercicio

    if tipo != "inicial":
         pi.write('05')#Opção
         pi.write('1')#Navegação     
         pi.hotkey('enter')

    
        
    pi.write(uo)
    pi.write(emp)
    emp=emp.strip()
    if len(emp)<7:
        pi.hotkey('tab')
    pi.write(od)
    pi.hotkey('enter')
    
    #ev=pi.confirm(text=f'Verifique emp={emp}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    # Tela ==> Anular Saldo Liquidado/Retencao/Multa Contratual
    
# =============================================================================
#     msg1="SEM SALDO A SER ANULADO"
    msg1='SALDO LIQUIDADO INEXISTENTE'
    n=1
    st=0.5
    rp=check_current_screen(msg1, n, st)
    pi.hotkey('home')
    
    print(rp)
#     
   
#     
    if rp=="OK":
          
          #ev=pi.confirm(text='Verifique proximos passos, F3 ou F9\n 05 -> 1\n OU uo -> emp', title='', buttons=['OK', 'Cancel'])
          #if ev=='Cancel':
              #sys.exit()
        
          pi.hotkey('f3')
          pi.hotkey('home')
          pi.write('05')#Opção
          pi.write('1')#Navegação     
          pi.hotkey('enter')
   
          return msg1
#         pi.hotkey('tab')
#         pi.write('51')
#         pi.hotkey('enter')
#         return resp


#           SALDO LIQUIDADO INEXISTENTE
# =============================================================================
    
    
    #CASO 1 CANCELA VALOR E RETENCAO
    #CASO 2 CANCELA SOMENTE VALOR
    #CASO 3 CANCELA SOMENTE RETENCAO
    
    
    msg1='RETENCAO DE CONTRIBUICOES PREVIDENCIARIAS - OUTRAS DESPESAS CORRENTES'
    n=1
    st=0.5
    rp=check_current_screen(msg1, n, st)
    print(rp)
    
    #ev=pi.confirm(text=f'Verifique rp={rp} e se \ncursosr se encontra em HOME', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    # GET VALOR DA LIQUIDACAO
    n=2
    st=1
    valor1=get_val_anula(n, st)
    print(valor1)
    
    if valor1=="NOK":
        resp=valor1
          # Retorna Tela Inicial
        pi.hotkey('f9')
        pi.hotkey('tab')
        pi.write('51')
        pi.hotkey('enter')
        return f"ANULACAO DE LIQUIDACAO NAO REALIZADA, EMPENHO-{emp}"   
       
        
    time.sleep(1)
    #ev=pi.confirm(text=f'Verifique caso de uso 1, 2 ou 3\n {rp} ==> {valor1}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()

    

    #CASO 1 CANCELA VALOR E RETENCAO
    if valor1!="000" and rp=="OK":
        pi.hotkey('home')
        pi.write('x')
        pi.write('x')
        pi.hotkey('enter')
        
        #CANCELANDO VALOR        
        pi.write('x')
        pi.write(valor1)
        pi.hotkey('f5')
     
        #ev=pi.confirm(text=f'Verifique proximas etapas 1', title='', buttons=['OK', 'Cancel'])
        #if ev=='Cancel':
            #sys.exit()
    
        # GET VALOR DA RETENCAO
        n=2
        st=0.5
        valor1=get_val_retencao(n, st)
        
        #CANCELANDO RETENCAO
        pi.hotkey('home')
        pi.write(valor1)
        pi.hotkey('f5')
        
        
        
    #CASO 2 CANCELA SOMENTE RETENCAO
    if valor1=="000" and rp=="OK":
        #ev=pi.confirm(text=f'VERIFIQUE PROCEDIMENTO - CASO 3', title='', buttons=['OK', 'Cancel'])
        #if ev=='Cancel':
            #sys.exit()
            
        pi.hotkey('home')
        pi.write('x')   
        pi.hotkey('enter')
            
        n=2
        st=1
        valor1=get_val_retencao(n, st)
            
        
        #CANCELANDO RETENCAO
        pi.hotkey('home')
        pi.write(valor1)
        pi.hotkey('f5')
         
         
# =============================================================================
#    #CASO 3 CANCELA SOMENTE VALOR
#     if valor1!="000" and rp=="NOK":
#         pi.hotkey('home')
#         pi.write('x')
#         
#         pi.write(valor1)
#        
#         # TELINHA  Registro Dt.Previsao Competencia
#         pi.hotkey('enter')    
#         pi.hotkey('f5')
# =============================================================================
        
    #CASO 3 CANCELA SOMENTE VALOR
    if valor1!="000" and rp=="NOK":
        pi.hotkey('home')
        
        pi.write('x')
        pi.hotkey('enter')
        time.sleep(1)
      
        pi.write('x')
        
        n=2
        st=1
        valor1=get_val_anula(n, st)
        
        pi.hotkey('home') 
        pi.hotkey('tab') 
        
        pi.write(valor1)
        
        #ev=pi.confirm(text=f'VERIFIQUE PROCEDIMENTO - CASO 3\n posicao do cursor e se valor digitado com sucesso\n{valor1}', title='', buttons=['OK', 'Cancel'])
        #if ev=='Cancel':
            #sys.exit()
       
        # TELINHA  Registro Dt.Previsao Competencia
        pi.hotkey('enter')    
        pi.hotkey('f5')
       
    # TELA MENSAGEM
    time.sleep(0.5)
    pi.write(f'REGISTRA ANULACAO DE SALDO LIQUIDADO DO EMPENHO {emp}')
    time.sleep(0.5)
    pi.hotkey('f5')
    pi.hotkey('f5')
    
    
    #ev=pi.confirm(text=f'Anote o numero da Anulacao de valor {valor1}\n\nProxima acao será F9', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    # TELA 0047-ANULACAO EFETUADA. 
    
    anul_num=get_num_anula(n,st)
       
    resp=f"ANULACAO DE LIQUIDACAO REALIZADO COM SUCESSO, EMPENHO-{emp} , anulacao_liquidacao : {anul_num}"
    
    print(resp)
    
    # Retorna Tela Inicial
    #pi.hotkey('f3')
    #pi.hotkey('f3')
    
    
    # Retorna Tela Inicial
    pi.hotkey('f9')
    pi.hotkey('tab')
     
    # Preparo inicial
    pi.write('51')
    pi.hotkey('enter')
    
    #ev=pi.confirm(text='Verifique a proxima acao\nProxima acao será 5 + 1 para Anulacao + registro ??', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    return resp

def preparo_inicial(oper):
        print('INICIO da funcao Preparo Inicial em siafi liquidacao')
    
        
        # Click no centro da tela
        a,b=pi.size()
        pi.click(x=a/2, y=b/2)
        #time.sleep(0.5)
        
        #pi.hotkey('tab')#Opção
        #time.sleep(2)
        #pi.hotkey('tab')#Opção
        #time.sleep(2)
        
        pi.hotkey('tab')#Opção
        #time.sleep(1)
        pi.write('51')#Navegação
        #time.sleep(1)
        pi.hotkey('enter')
        #time.sleep(1)
        
        if oper=="Liquidacao":
            pi.write('04')#Opção
            
        if oper=="Anula Liquidacao":
            pi.write('05')#Opção
            
        pi.write('1')#Acão    
        pi.hotkey('enter')
          
        time.sleep(0.5)
        
        return


    
def ctl():    
    time.sleep(3)
    
    pi.FAILSAFE = True
    
    pi.PAUSE=0.4
    
    
    oper="Registra Liquidacao"
    #oper="Anula Liquidacao"
    
    
    
    if oper=="Registra Liquidacao":
        #file=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_2\registra_liquidacao.xlsx"
        file=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\03_2024\PT_liquidacao_siafi_1_a_2000\registra_liquidacao.xlsx" #Maquina Lewnovo
        file_out=file
        #file_out=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_2\resposta_registra_liquidacao.xlsx"
    
    if oper=="Anula Liquidacao":
        file=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2024\dativos 3a listagem\lista_2\anula_liquidacao.xlsx"
        file_out=file
        #file_out=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_2\resposta_anula_liquidacao.xlsx"
    
    
    
    
    #local="accer_1"
    local="itautec"
    #local="robo-server"
    
    #list_wins()
    #list_class(name)
    #abre_siafi_run(local)
    import sys
    
    import nest_asyncio
    nest_asyncio.apply()
    
    
    df=pd.read_excel(file)
    
    #df['Resultado'] = pd.Series(dtype='str')

    df=df.fillna("X")
    df=df.round(decimals=2)    
    
    print(df)
    
    # TRATA NOME DAS COLUNAS
    column_names = list(df.columns)
    
    s=column_names[1] # sei / cpha
    p=column_names[2] # processo
    c=column_names[4] # cpf
    n=column_names[5] # nit
    b=column_names[6] # banco
    a=column_names[7] # agencia
    dva=column_names[8] # dv agencia
    ct=column_names[9] # conta
    dvct=column_names[10] # dv conta
    v=column_names[11] # valor arbitrado
    r=column_names[13] # valor devido INSS
    o=column_names[23] #observação
    rr=column_names[24] # resultado
    
    print(c, b, a, dva, ct, dvct, v, r, o, rr)
    
    types={s:str , p:str,  c:str,  n:str, b:str, a:str, dva:str, ct:str, dvct:str, v:str, r:str, o:str, rr:str}
    
   
    df2=df.astype(types)
    print(df2.dtypes)  
   
   
    ue="1080012" # A unidade executora para o Dativo Administrativo é 1080012.
    
    login="m1379117"
    senha="xx2024"
    
    if login=="xxxx" or senha=="xxxx":
        #ev=pi.confirm(text='Registe seu login e senha do SAIFI,\nnas linhas 1161 e 1162 do programa,\nantes de prosseguir', title='', buttons=['OK', 'Cancel'])
        ev=pi.alert(text='Registe seu login e senha do SAIFI,\nnas linhas 1161 e 1162 do programa,\nantes de prosseguir', title='')
        if ev=='Cancel':
            sys.exit()
    
    local="itautec"
    
    #controle_siafi contains abre_siafi e loga_siafi
    resp=asyncio.run(controle_siafi(login, senha, ue, local)) 
    
    print("de volta a ===>  __name__ == '__main__':")
        
     ### LIQUIDACAO ###
    # Unidadae executora
    #ue="1080012" # A unidade executora para o Dativo Administrativo é 1080012.
    #unidade orcamentaria = uo
    uo="1081"
    #ordenador de despesa = od
    #od="10606119" # Dr fabio da PT
    od="1120530" # Dra Karem PT
    
    # Prog trabalho 
    #p_trab ='10140001' 
    
    # #Nat. Despesa
    #nat_desp ='339036'
    
    # Item
    #item = '21'
    
    # Fonte. Proc/IPE
    #fonte = '1090'
    
    preparo_inicial()
    
    print("INICIO do loop")
    x=len(df)
    print(x)
    #k=0
    excessao="N"
    
    for i in range(x):
        
        print(f'INICIO DA LINHA {i}')
           
        # df.iloc[linha,coluna]
        cpf_cnpj=df2.iloc[i,4]
        cpf_cnpj=trata_cpf_cnpj(cpf_cnpj)
        
        valor=df2.iloc[i,11]#valor arbitrado
        valor=trata_valor_4(valor)
        
        inss=df2.iloc[i,13]#valor devido inss
        inss=float(inss)
        inss = round(inss,2)
        inss=str(inss)
        inss=inss.strip()
        
        
        sei=df2.iloc[i,1]
        sei=trata_sei_cpha(sei)
        
        processo=df2.iloc[i,2] # processo
        processo=trata_processo(processo)
              
       
        print(f'cpj_cnpj ==> {cpf_cnpj} , {type(cpf_cnpj)}')
        
        print(f'valor ==> {valor}  , {type(valor)}')
        
        #sys.exit()
        
        print(f'linha ==> {i}')
        
       
        if i==0:
            tipo="inicial"
        else:
            tipo="outros"
            
        if excessao=="SS":
            tipo="inicial"
            
            
        ct2=""    
        if oper=="Anula Liquidacao":
            ######################    
            #CANCELA LIQUIDACAO
            ######################
            obs=df2.iloc[i,23]
            obs=obs.strip()
            pos23=obs.find("-")
            pos24=obs.find("NAO")
            pos25=obs.find("NOK")
            if (pos23>=0 and pos24<0 and pos25<0):
                liq_num=obs[pos23+1:]
                liq_num=liq_num.strip()
                emp=liq_num
                liq=anula_liquidacao(emp,uo,tipo,valor,od)   
            else:
                liq=obs
                ct2="ja executado"
                #liq=f"NAO FOI POSSIVEL CANCELAR LIQUIDACAO: {obs}"
                if tipo=="inicial":
                    print("inicial")
                    #excessao="SS"
                if tipo=="outros":
                    print("outros")
            
            
            
        if oper=="Registra Liquidacao":
            ######################    
            #EXECUTA LIQUIDACAO
            ######################
            
            obs=df2.iloc[i,23]
            obs=obs.strip()
            
            
            pos24=obs.find("NAO")
            pos25=obs.find("NOK")
            pos26=obs.find("LIQUIDACAO")
            
            pos27=obs.find("EMPENHO DA DESPESA REALIZADO")
            pos23=obs.find("-")
            pos28=obs.find("SUCESSO")
            
            print(obs)
            
            #ev=pi.confirm(text='verifique campop obs', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            
            ob1="LIQUIDACAO NAO REALIZADA, EMPENHO"
            #or (pos27>=0 and pos23>=0 and pos28>=0 and pos24<0 and pos25<0)
            if (ob1 in obs) or (pos27>=0 and pos23>=0 and pos28>=0 and pos24<0 and pos25<0) or (pos27>=0 and pos23>=0 and pos28>=0 and pos24<0 and pos25<0) or(pos26>=0 and pos24>=0) or (pos26>=0 and pos25>0) or obs=="nan" or obs=="X" or obs=="x" or ("REPETIR" in obs):
                num_emp=obs[pos23+1:]
                num_emp=num_emp.strip()
                liq=liquida_dativo(num_emp,uo,tipo,valor,od, inss)
                if (ob1 in liq) and i>6:
                   obb1=df2.iloc[i-1,23]
                   obb2=df2.iloc[i-2,23]
                   #obb3=df2.iloc[i-3,23]
                   if (ob1 in obb1) and (ob1 in obb2):# and (ob1 in obb3):
                        print("RESTART")
                        #Fecha janela SIAfI
                        pi.hotkey('alt','f4')
                        return 'restart'
                        #sys.exit()
            else:
                liq=obs
                ct2="ja executado"
                #liq=f"NAO FOI POSSIVEL REALIZAR LIQUIDACAO: {obs}"
                
                #ev=pi.confirm(text='Faltou numero do empenho em obs\nVerifique qual tecla deve ser clicada em seguida', title='', buttons=['OK', 'Cancel'])
                #if ev=='Cancel':
                   # sys.exit()
               
                #pi.hotkey('f3')
            
        
        # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
        # SE OPERACAO NÃO REALIZADA, NÃO ATUALIZA DF2, NEM SALVA EXCEL
        
        if ct2!="ja executado":
          
        
            print("=========> Change a cell value in dataframe")
            df2[s][i] = f"'{sei}"# coluna sei / cpha
            df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
            df2[p][i] = f"'{processo}" # processo
            #df2[v][i] = valor # coluna valor
            df2[o][i] = liq # coluna observacoes
            
            if "SUCESSO" in liq:
                df2[rr][i]="LIQUIDACAO REALIZADA COM SUCESSO"
            else:
                if "LIQUIDACAO NAO REALIZADA" in liq:
                    liq="LIQUIDACAO NAO REALIZADA"
                df2[rr][i]=liq
            
            #print(df2)
            #print todas linhas e colunas c,v,o only
            print(df2.loc[:,[c,v,o]])
            
            """
            k+=1
            if k==10:
                k=0
                print(df2)
            """
            
            df2.to_excel(file_out , index=False)
        
        #if i==2:
            #break
        
    return "OK"
    #pi.alert(text='---> FIM  LIQUIDACAO <---', title='', button='OK')

       
if __name__ == '__main__':
    
    x=''
    for i in range(10):
        if i==0:
            resp=ctl()
        if resp=='restart':
            resp=ctl()
        if resp=="OK":
            print('FIM')
            break
            
    