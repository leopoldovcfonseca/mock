# -*- coding: utf-8 -*-
"""
Created on Tue Dec 13 10:46:50 2022

@author: m1379117
"""



import pyautogui as pi
import time
import subprocess
import os
import sys
import asyncio
from asgiref.sync import sync_to_async
import pyperclip
import time
import re
import pandas as pd
import datetime, calendar


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



#############
pi.PAUSE=0.4
pi.FAILSAFE=True
#############


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
        #df2 = df.iloc[:,[3,4]] # Select columns by Index 3 and 
        df2 = df.iloc[:,[4,18]] # Select columns by Index 4 and 19 = cpf e valo liquido na planilha dativos
        
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



def abre_siafi_run(local):
    
    print("INICIO da funcao  abre_siafi_run Pagamento")
    
    import sys
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
    
    #mm="UNABLE TO ESTABLISH SESSION"
    #n=2
    #st=0.5
    #verif=check_current_screen(mm, n, st)
    #pi.autogui('home')
    
    #if verif=="OK":
        #return mm

    msg= "OK from abre_siafi_run"
    print(msg)
    
    return

def loga_siafi(login, senha, ue, local):
    print("INICIO da funcao loga_siafi")
    
    msg1="Logon executado com sucesso"
    msg2="Unidade Executora:"
    
    import pyautogui as pi
    import time
  
    
    #time.sleep(1)
    
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
        time.sleep(1.5)
        pi.hotkey('ctrl' , 'a')
        pi.hotkey('alt' , 'e')
        pi.hotkey('enter')
          #pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        #print(x)
        pos1=x.find(msg1)
        if pos1 >=0:
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
    
    n=5
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
        time.sleep(1)
        rsp=check_current_screen(msg1, n, st)
        if rsp=="NOK":
            resp="NOK-2 from loga_siafi"
        resp="OK-2 from loga_siafi"
    else:
        resp="OK-1 from loga_siafi-1"

    print(resp)
    
    
    
    
    return resp



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


def check_current_screen_banco(msg1, n, st):
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
            pos2=x.find("-")
            resp1=x[pos2+1:pos1+11]
            resp=["OK" , resp1]         
            break
        else:
            resp=["NOK" , "X"]
            print(f' ciclo {i+1} não localizado {msg1}')
        
        i+=1
        if i+1==n:
            resp="NOK"
    
    pi.hotkey('tab')
    print(resp)
    return resp


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
 
async def controle_siafi(login, senha, ue, local, ano):
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
    
    #https://superfastpython.com/asyncio-task-done-callback-functions/
    
    #def verifica_unable(task1):
        #if task1=="UNABLE TO ESTABLISH SESSION":
             #pi.confirm(text='UNABLE TO ESTABLISH SESSION\n O SAIFI se encontra congestionado. Agurado alguns minutos e tente novamente.\nO prgrama será encerrado.', title='controle_siafi', buttons=['OK', 'Cancel'])
             #print(unable_to_establish_session)
    
    #task1.add_done_callback(verifica_unable)
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


def ajusta_contas(vv):
    
    print('INICIO da funca ajusta_contas')
    print(f'entrada={vv}')
       
    vv=vv.strip()
    pos22=vv.find(".")
    if pos22>=0:
        vv=vv[:pos22]
        vv=vv.strip()
        
    pos22=vv.find(",")
    if pos22>=0:
        vv=vv[:pos22]
        vv=vv.strip()
    
   
    print(f'saida={vv}')
    return vv
    

def ajusta(valor):
    
    print('INICIO da funca ajusta')
    print(f'entrada={valor}')
    
    valor=str(valor)
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
    
    print(f'saida={valor}')
    
    return valor

def data_prevista():
    """"retorna utimo dia do mes"""
    # TRATA DATAS
    x = datetime.datetime.now()
    print(x)
    dia=x.strftime("%x")
    print(f'today ==> {dia}')
    yy=x.year
    mm=x.month
    dd=x.day
    
    # MES CORRENTE
    res = calendar.monthrange(yy, mm)
    dd1 = res[1]
    print(dd1)
    
    dd1=str(dd1)
    mm=str(mm)
    yy=str(yy)
    
    resp=[dd1,mm,yy]
    return resp


def get_doc_number(n, st, cpf_cnpj):
    
    """ Busca numero do documento pagamento na tela"""
    
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
    cpf=cpf_cnpj
    
    
    
    msg1="INCLUSAO EFETUADA"
    
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
    
    y=x
    
    """ GET DOC NUMBER ->  Nr. OP """
    cpf=cpf[:9]
    print(f'cpf==> {cpf}')
    
    # Nr. OP    Credor OP     Vl. Ordem Pagto   Vl. IRRF   IRRF 
    # 37891    091624936-04      178,00                     588 
        
    #IRRF 
    #37891    091624936
      
    #Nr. OP    Credor OP          Vl. Ordem Pagto           Vl. IRRF              
    #47096    856793306-44                   1.038,63                         
    
    
    t2=cpf
    
    pos2=x.find(t2)
    
    x1=x
    
    x=x[pos2+10:]
    
    t1="IRRF"
    
    pos1=x1.find(t1)
    pos2=x1.find(t2)
   
    
    if pos1<0 or pos2<0:
        resp="nao localizado"
        print(f'Numero dos Documentos = {resp}')
        return resp
    
    x1=x1[pos1+5:pos2]
    x1=x1.strip()
    #IRRF 
    #37891  
       
    #x1=x1[4:]
    x1=x1.strip()
    #37891  
       
    op="OP-"+x
    print(op)
    
    
    x=y
    
    """# GET GLOBAL NUMBER """
    #Nr. Documento Global: 6092                Ano Empenho: 2022
    #Nr. Documento Global: 37119            Credor Empenho:
    
    t1="Nr. Documento Global:"
    t2="Credor Empenho:"
    pos1=x.find(t1)
    print(pos1)
    pos2=x.find(t2)
    print(pos2)
    docg=x[pos1+21:pos2]
    
    print(docg)
    
    docg=docg.strip()
    
    resp=[]
    resp.append(op) # OP number
    resp.append(docg) # global number
    
    print(f'Numeros dos Documentos = {resp}')
    return resp


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
    
    print("INICIO da funcao trata_valor")
    
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
        #vv=item
        try:
            vv=vv.replace(",",".")
        except:
            pass
        vv=float(vv)
        resp="%0.2f" % (vv,)
        resp=resp.replace(".","")
    print(f'{vv1} ==> {resp}')
    
    return resp

def trata_irr(irr):
    
    print("INICIO da funcao trata_irr")
    print(irr, type(irr))
    
    irr=irr.strip()
    if irr=="" or irr=="x" :
        irr="0"
    
    try:
        irr=float(irr)
    except:
        irr=irr.replace(",",".")
        irr=float(irr)    
   
    irr = round(irr,2)
    irr=str(irr)
    irr=irr.strip()
    
     #valor=str(valor)
    print(irr)
    pos21=irr.find(".")
    pos22=irr.find(",")
    print(pos21)
    print(pos22)
    
    if pos21<0 and pos22<0:
        irr=f'{irr}00'
        
    print(len(irr))
    ll1=len(irr)-pos21
    if ll1==2:
       irr=f'{irr}0'
        
    print(irr)    
    irr=irr.replace(".","")
    irr=irr.replace(",","")
    irr=irr.strip()
    
    print(f'irr corrigido = {irr}')
        
    return irr

def preparo_inicial(oper):
    print('INICIO da funcao Preparo Inicial em siafi pagamento')
    
        
    # Click no centro da tela
    a,b=pi.size()
    pi.click(x=a/2, y=b/2)
    ##time.sleep(0.5)
    
    #pi.hotkey('tab')#Opção
    #time.sleep(2)
    #pi.hotkey('tab')#Opção
    #time.sleep(2)
    #time.sleep(1)
    pi.hotkey('tab')#Opção
    #time.sleep(1)
    pi.write('71')#Navegação
    #time.sleep(1)
    pi.hotkey('enter')
    #time.sleep(1)
    
    
    #print(xxxx)
    
    pi.hotkey('enter')
    #time.sleep(1)
        
    pi.write('01')#Opção
    #time.sleep(0.5)
    pi.write('1')#Acão
    time.sleep(0.5)
    
    pi.hotkey('enter')
    
        
    
    #
    
    mm="TRANSMISSAO AOS BANCOS EM PROCESSAMENTO"
    n=1
    st=0.5
    verif=check_current_screen(mm, n, st)
    pi.hotkey('home')
    #pi.hotkey('shift','tab')
    #pi.hotkey('shift','tab')
    #pi.hotkey('esc')
    
    
    
    #ev=pi.confirm(text='Verifique posicao do cursor e msg TRANSMISSAO AOS BANCOS EM PROCESSAMENTO', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    if verif=="OK": # localizado TRANSMISSAO AOS BANCOS EM PROCESSAMENTO      
      
       
        resp=mm
        
        print(resp)
        
        """
        
        # Volta tela anterior
        pi.hotkey('f3')
        time.sleep(0.5)
        pi.write('01')#Opção
        time.sleep(0.5)
        pi.write('1')#Acão
        time.sleep(0.5)
        pi.hotkey('enter')
       """ 

        return resp 
    
    else:
        pi.PAUSE=0.4
        return "OK"

def paga_dativo(emp, uo, tipo1, cpf_cnpj, banco, agencia, dv_a,  conta,  dv_c, valor, valor_bruto, od, inss, irr, sei, processo, i, ano):
    
    #def paga_dativo(cpf_cnpj, valor, uo, od, p_trab, nat_desp, item, fonte, tipo, sei):
    print("INICIO da funcao paga_dativo")
    # emp=numero do empenho
    # uo = unidade orcamentaria =  "1081"
    # tipo1 = "inicial" ou "não inicial"
    # valor = valor tipo 2000 equivale a 20 reais = confirmado OK
    # od = ordenador de despesa = "1120530" Dra KAren
    # valor do desconto inss
    
    
    #pi.PAUSE=0.4
    #valor=str(valor)
    #print(valor)
    #rint(emp)
    
    print("==================================================================================")
    
    #ev=pi.confirm(text='Verifique melhor procedimento utilizando home shift-tab', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    time.sleep(0.5)
    pi.hotkey('shift','tab')
    #time.sleep(0.5)
    pi.hotkey('shift','tab')
    #time.sleep(0.5)
    pi.write('X')
    #time.sleep(0.5)
    
    #=pi.confirm(text=f'Verifique se os x foi acionado corretamente. Proxima tecla será ENTER', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    pi.hotkey('enter')
    time.sleep(0.5)
       
    
    # TELINHA
    #pi.hotkey('tab')
    #time.sleep(0.5)
    #pi.hotkey('tab')
    #time.sleep(0.5)
    pi.hotkey('shift','tab')
    #time.sleep(0.5)
    pi.write('X')
    #time.sleep(0.5)
    pi.hotkey('enter')
    time.sleep(0.5)    
    
    # TELA  Registrar Ordem de Pagamento Bancaria
    
    # TRATA DATAS
    x = datetime.datetime.now()
    print(x)
    dia=x.strftime("%x")
    print(f'today ==> {dia}')
    mm=x.month
    yy=x.year
    mm=str(mm)
    yy=str(yy)
    print(mm)
    print(yy)
    print(f'Ano = {ano}')
    
    
    if ano!=yy:
        #yy=ano
        print("OPeracao Restos a pagar")
    else:
        print("OPeracao Pagamento")
    
    
    # ano
    pi.write(ano)
    #time.sleep(0.5)
    pi.write(emp)
    ##time.sleep(0.5)
    pi.hotkey('tab')
    #time.sleep(0.5)
    pi.write(uo)
    #time.sleep(0.5)
    
    pi.write('x')
    #time.sleep(0.5)
    
    pi.hotkey('down')
    #time.sleep(0.5)
    
    pi.hotkey('tab')
    #time.sleep(0.5)
    
    
     # BANCO
    banco=ajusta_contas(banco)
    banco=banco.strip()
    if len(banco)<3:
        for j in range(3-len(banco)):
            banco=f'0{banco}'
    pi.write(banco)
    
    print(f'banco ===> {banco}')
           
    
    # AGENCIA
    agencia=ajusta_contas(agencia)
    print(f'agencia ==> {agencia}')
    
    
    #f banco=="001" or banco=="237" or banco=="104":
    agg=agencia
    #dva_a=dva
    
    #if banco=="033":
        #agg=f'{agg}{dv_a}'
    
    pi.write(agg)
    if len(agg)<5:
        pi.hotkey('tab') #
    
        
     # DIGITO VERIFICADOR AGENCIA
    dv_a=ajusta_contas(dv_a)
    print(f'Digito verificador ==> {dv_a}')

    #ev=pi.confirm(text=f'Verifique valor de dv_a => {dv_a}' , title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()    
    
    
    # NAO INFORMAR DV
    if banco=="033" or banco=="389" or banco=="341" or banco=="260" or banco=="77" or banco=="756" or banco=="077" or banco=="336" or banco=="623" or banco=="748":
        pi.hotkey('tab') # sem dv_a
    # NFORMAR DV
    else:  #if banco=="104" or banco="389":
        if dv_a=="x": # Veio vazio no original por isto está com x letra minuscula
            pi.hotkey('tab') # sem dv_a
        else:
            pi.write(dv_a)
    
    dv_a=dv_a.strip()   
    
   
        
    if dv_a=='x':
        dv_a = None
        
    #ev=pi.confirm(text=f'Verifique NOVO valor de dv_a => {dv_a}' , title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
    # CONTA
    conta=ajusta_contas(conta)
    conta=conta.strip()
    if len(conta)<12:
        for j in range(12-len(conta)):
            conta=f'0{conta}'
    pi.write(conta)
    #pi.hotkey('tab')
    
        
    # DIGITO VERIFICADOR CONTA
    # dv conta
    dv_c=ajusta_contas(dv_c)
    if dv_c=="nan":
        #pi.hotkey('tab')
        dv_c="0"
    #time.sleep(0.5)
    pi.write(dv_c)  
    
    print("ano, banco, agencia, dv_agencia, conta, dv_c")
    print(yy, banco, agg, dv_a, conta, dv_c )
    print(type(yy), type(agg), type(dv_a), type(conta), type(dv_c) )
   
    
    bac=[banco,agencia,conta]
    import re
    for vv in bac:
        ld1=re.findall("\d",vv)
        if len(ld1)==0:
            resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, BANCO, AGENCIA OU CONTA INEXISTENTE {banco} / {agg} / {dv_a} / {conta} / {dv_c}'
            print(resp)
            # Volta tela anterior
            pi.hotkey('f3')
            #.sleep(0.5)
            pi.write('01')#Opção
            #time.sleep(0.5)
            pi.write('1')#Acão
            #time.sleep(0.5)
            pi.hotkey('enter')
            return resp
    
    
    """CREDOR DA OP = CPF """
    print(cpf_cnpj)
    x=len(cpf_cnpj)
    
    #ev=pi.confirm(text=f'Verifique cpf_cnpj ==> {cpf_cnpj} {x} digitos \n linha da planilha = {i+2}' , title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    
    if x==10:
        cpf_cnpj=f'0{cpf_cnpj}'
        x=len(cpf_cnpj)
    if x==11:
        tipo="cpf"
    if x==13:
        cpf_cnpj=f'0{cpf_cnpj}'
        x=len(cpf_cnpj)
    if x==14:
        tipo="cnpj"
        
    if x!=11 and x!=14:
        resp=f"cpf_cnpj invalido ==> {cpf_cnpj}"
        # Volta tela anterior
        pi.hotkey('f3')
        time.sleep(0.5)
        pi.write('01')#Opção
        time.sleep(0.5)
        pi.write('1')#Acão
        time.sleep(0.5)
        pi.hotkey('enter')
        return resp 
    
    if tipo=="cpf":
        pi.write(cpf_cnpj)
        pi.hotkey('tab')
    if tipo=="cnpj":
        pi.write(cpf_cnpj)
        
    
    """ PROCESSO """
    processo = re.sub(r'[^0-9\s]', '', processo) # RETIRA . - etc.../
    pi.write(processo)
    
    pi.write('1')
    pi.hotkey('tab')
    
    pi.write(valor)
    pi.hotkey('tab')
    
    
    """ IRRF (S/N) """
    if len(irr)==0 or irr=="0" or irr==None or irr=="000" or irr=="x" or irr=="X":
        pi.write('N')
        irr_check="N"
    else:
        pi.write('S')
        irr_check="S"
    
    #if len(inss)==0 or inss=="0" or irr==None or inss=="000" or inss=="x" or inss=="X":
    #    inss_check="N"
    #else:
        
    """INSS = 01"""
    #inss_check="S"
    pi.write('01')
    
    #pi.write('588')
    #pi.write('01')
    #time.sleep(1)
    
    #ev=pi.confirm(text='Testar IRR = S e IRR = N\n Testar qte contrib 01 e vazio', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
    pi.hotkey('home')
    pi.hotkey('shift','tab')
   
    """ ORDENADOR DE DESPESAS """
    time.sleep(0.5)
    pi.write(od)
   
    
    #ev=pi.confirm(text='Verificar se down funcionou e OD', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
        
    pi.hotkey('enter')
    
    mm="NAO DISPONIVEL PARA PAGAMENTO"
    n=1
    st=0.5
    verif=check_current_screen(mm, n, st)
    pi.hotkey('tab')
    pi.hotkey('tab')
    pi.hotkey('esc')
    if verif=="OK": # localizado inexistente
        resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, RESTO A PAGAR/SALDO/EMPENHO NAO DISPONIVEL PARA PAGAMENTO'
        print(resp)
        # Volta tela anterior
        pi.hotkey('f3')
        pi.write('01')#Opção
        #time.sleep(0.5)
        pi.write('1')#Acão
        #time.sleep(0.5)
        pi.hotkey('enter')
        return resp   
            
    
    #pi.hotkey('home')
    #time.sleep(0.5)
    #pi.hotkey('shift','tab')
    #time.sleep(1)
    
    
    #ev=pi.confirm(text=f'Verifique tela com IRRF-1', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    #     Ordem de Pagamento / Quitacao
    mm='Ordem de Pagamento / Quitacao'
    n=1
    st=0.5
    verif=check_current_screen_banco(mm, n, st)
    pi.hotkey('esc')
    
    
    
    if verif[0]!="OK":
        
        # Aqui Verificar 0101- AGENCIA A CREDITAR INEXISTENTE(S).
        # Aqui Verificar 0101- BANCO A CREDITAR INEXISTENTE
        # Aqui Verificar 0101- CONTA A CREDITAR INEXISTENTE
        # Aqui Verificar 0119-TOTAL A PAGAR. MAIOR QUE O SALDO DO EMPENHO
        
        
        
        # # NAO TEM IRRF NEM INSS
        # mm="INFORME A QUANTIDADE DE CONTRIBUINTES"
        # n=1
        # st=0.5
        # verif=check_current_screen(mm, n, st)
        # pi.hotkey('shift','tab')
        # pi.hotkey('esc')
        # if verif[0]=="OK": # localizado inexistente
        #     add=verif[1]
        #     add=add.strip()
        #     resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, {mm}'
        #     print(resp)
        #     # Volta tela anterior
        #     pi.hotkey('f3')
        #     time.sleep(0.5)
        #     pi.write('01')#Opção
        #     time.sleep(0.5)
        #     pi.write('1')#Acão
        #     time.sleep(0.5)
        #     pi.hotkey('enter')
        #     return resp
        
        # Aqui Verificar TOTAL A PAGAR. MAIOR QUE O SALDO DO EMPENHO
        mm="TOTAL A PAGAR. MAIOR QUE O SALDO DO EMPENHO"
        n=1
        st=0.5
        verif=check_current_screen_banco(mm, n, st)
        pi.hotkey('shift','tab')
        pi.hotkey('esc')
        if verif[0]=="OK": # localizado inexistente
            add=verif[1]
            add=add.strip()
            resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, {mm}'
            print(resp)
            # Volta tela anterior
            pi.hotkey('f3')
            time.sleep(0.5)
            pi.write('01')#Opção
            time.sleep(0.5)
            pi.write('1')#Acão
            time.sleep(0.5)
            pi.hotkey('enter')
            return resp
        
        # Aqui Verificar INEXISTENTE
        mm="INEXISTENTE"
        n=1
        st=0.5
        verif=check_current_screen_banco(mm, n, st)
        pi.hotkey('shift','tab')
        pi.hotkey('esc')
        if verif[0]=="OK": # localizado inexistente
            add=verif[1]
            add=add.strip()
            resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, {add} , {banco} / {agg} / {dv_a} / {conta} / {dv_c}'
            print(resp)
            # Volta tela anterior
            pi.hotkey('f3')
            time.sleep(0.5)
            pi.write('01')#Opção
            time.sleep(0.5)
            pi.write('1')#Acão
            time.sleep(0.5)
            pi.hotkey('enter')
            return resp 
        
        mm="BLOQUEADO"
        n=1
        st=0.5
        verif=check_current_screen(mm, n, st)
        pi.hotkey('shift','tab')
        pi.hotkey('esc')
        if verif=="OK": # localizado inexistente
            resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, AGENCIA OU CONTA BLOQUEADA {banco} / {agg} / {dv_a} / {conta} / {dv_c}'
            print(resp)
            # Volta tela anterior
            pi.hotkey('f3')
            pi.write('01')#Opção
            #time.sleep(0.5)
            pi.write('1')#Acão
            #time.sleep(0.5)
            pi.hotkey('enter')
            return resp 
        
        
        
        mm="AGENCIA A CREDITAR"
        n=1
        st=0.5
        verif=check_current_screen(mm, n, st)
        pi.hotkey('shift','tab')
        pi.hotkey('esc')
        
        if verif=="OK": # localizado inexistente
            resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, ERRO NA AGENCIA OU DV-AGENCIA {banco} / {agg} / {dv_a} / {conta} / {dv_c}'
            print(resp)
            # Volta tela anterior
            pi.hotkey('f3')
            pi.write('01')#Opção
            #time.sleep(0.5)
            pi.write('1')#Acão
            #time.sleep(0.5)
            pi.hotkey('enter')
            return resp
        
        mm="Input value for a numeric field is not numeric"
        n=1
        st=0.5
        verif=check_current_screen(mm, n, st)
        pi.hotkey('tab')
        pi.hotkey('tab')
        pi.hotkey('esc')
        if verif=="OK": # localizado inexistente
            #resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp}, RESTO A PAGAR NAO DISPONIVEL PARA PAGAMENTO'
            resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp} , Perda de Sincronismo_1'
            print(resp)
            
            #ev=pi.confirm(text='Verifique porque pagamento não realizado,\n linha 1290 \n porque não localizou -PAGAMENTO EFETUADO-\nfaltou DOC number\n ????', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            
            # Volta tela anterior
            pi.hotkey('f3')
            pi.write('01')#Opção
            #time.sleep(0.5)
            pi.write('1')#Acão
            #time.sleep(0.5)
            pi.hotkey('enter')
            return resp
            
    
    # # INSS(N) IRR(S)
    # if irr_check=="S" and inss_check=="N":
    #      resp="INFORME A QUANTIDADE DE CONTRIBUINTES, sem IRR nem INSS"
    #     # Volta tela anterior
    #     pi.hotkey('f3')
    #     pi.write('01')#Opção
    #     #time.sleep(0.5)
    #     pi.write('1')#Acão
    #     #time.sleep(0.5)
    #     pi.hotkey('enter')
    #     return resp
    
    
    # # INSS(N) IRR(N) ambos nulos:    
    # if irr_check=="N" and inss_check=="N":
    #     pi.hotkey('enter')
    
    
    # INSS (S) e IRR (S/N)
    #if inss_check=="S": #irr_check=="S" or irr_check=="N" and 
    print(f"Tem irr = {irr} e inss ={inss}")
    #ev=pi.confirm(text=f'Verifique tela com IRRF-2', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    pi.hotkey('home')
    
    if irr_check=="S":
        pi.write(irr)
        pi.hotkey('enter')
    
    pi.write(cpf_cnpj)
    pi.hotkey('tab')
    
    pi.write('588')
    pi.hotkey('tab')
    
    if irr_check=="S":
        pi.write(irr)
        pi.hotkey('tab')
    
    #pi.write(valor) #2394,26
    pi.write(valor_bruto)
    pi.hotkey('tab')
    
    
    pi.write(inss) # INSS devido 26337
    pi.hotkey('tab')
    
    #pi.write('000') # pensao alimenticia
    pi.hotkey('tab')
    
    #pi.write('000') # dependentes
    pi.hotkey('tab')
              
    pi.hotkey('tab') #valor nao tributado
    
            
    
    #ev=pi.confirm(text='IRR e INSS\nVerifique tela esta correta\nInformacoes Adicionais para DIRF', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
    pi.hotkey('enter')
     
    #ev=pi.confirm(text='FAVOR CANCELAR PARA EVITAR PAGAMENTO POR F5', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    
    pi.hotkey('f5')
    time.sleep(0.5)
    
    
    #valores=[cpf_cnpj , '588' , valor_bruto, inss]
    #print(valores)
    #ev=pi.confirm(text='Verifique valores inseridos na tela;\\nProxima tecla será previsto ENTER', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
         #sys.exit()
    
    
    # SEM ALTERAR DATA PREVISTA DE PAGAMENTO
    pi.hotkey('enter')
        

    
    
    """ TELA DE MENSAGEM FINAL """
    
    #ev=pi.confirm(text='Verifique se abriu tela de mensagem', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    inss_msg=inss[:-2],inss[-2:]
    print(inss_msg)
    
    irr_msg=irr[:-2],irr[-2:]
    print(irr_msg)
    
    #ev=pi.confirm(text='Verifique o valor de:\nINSS R$ {inss_msg}\nIRR R$ {irr_msg}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    msg1='PAGAMENTO REFERENTE A ADVOGADO DATIVO ADMINISTRATIVO'
    msg2=f'RETIDA CONTRIBUICAO PREVIDENCIARIA DE R$ {inss_msg}'
    msg25=f'RETIDA CIMPOSTO DE RENDA DE R$ {irr_msg}'
    if '1080' in sei:
        msg3=f'SEI: {sei}'
    else:
        msg3=f'CPHA : {sei}'
    msg4=f'PROCESSO : {processo}'
    
    #ev=pi.confirm(text=f'Verifique se tela registro msgs\n previsto escrita de msg.' , title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
        
   
    # TELINHA HISTORICO DE LIQUIDACAO
    pi.write(msg1)
    pi.hotkey('tab')
    pi.write(msg2)
    pi.hotkey('tab')
    pi.write(msg25)
    pi.hotkey('tab')
    pi.write(msg3)
    pi.hotkey('tab')
    pi.write(msg4)
    
    time.sleep(2)
    
    #ev=pi.confirm(text=f'Verifique msg OK ;\Qual tecla deve ser clicada \n previsto F5', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
        
    pi.hotkey('f5')
        
    # TECLE PF5 PARA CONFIRMAR OU PF2 PARA ANULAR 
    # GERAR OP
    pi.hotkey('f5')
    
    
    #=============================================================================
    # VERIFICA SE registro efetuado e busca doc number e global number
    n=1
    st=0.5
    doc=get_doc_number(n, st, cpf_cnpj)
    pi.hotkey('esc')
    
    #ev=pi.confirm(text=f'Verifique numero doc na tela = {doc}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
    
    if doc=="NOK": # Registro não efetuado
        resp=f'PAGAMENTO NAO REALIZADO, EMPENHO-{emp} , Perda de Sincronismo_2'
        print(resp)
        doc=get_doc_number(n, st, cpf_cnpj)
        pi.hotkey('esc')
        if doc=="NOK": # localizou
            print(resp)
            #ev=pi.confirm(text='Verifique porque pagamento não realizado,\n linha 1388 \n porqueue não localizou -PAGAMENTO EFETUADO-\nfaltou DOC number\n ????', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            
            return resp
    
    #else    
    if doc!="NOK":    
    
        doc1=doc[0] # doc number
        doc2=doc[1] # # global number
        resp=f'PAGAMENTO REALIZADO COM SUCESSO, EMPENHO-{emp}, Nr.{doc1}, DOCUMENTO_GLOBAL-{doc2}'
    
    #======================================================================
    
    #ev=pi.confirm(text=f'Verifique valores capurados para empenho e variavel global:/n==> {doc1} e   {doc2} ', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
       #sys.exit()

    #resp=f'PAGAMENTO REALIZADO COM SUCESSO, EMPENHO-{emp}, DOCUMENTO-{doc1}, GLOBAL-{doc2}'
    #resp=f'PAGAMENTO REALIZADO COM SUCESSO, EMPENHO-{emp}'
    #print(resp)
    
    #ev=pi.confirm(text='Verifique se Proxima tecla prevista será F3', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
       #sys.exit()
       
       
    # Volta tela anterior
    pi.hotkey('f3')
    
    pi.write('01')#Opção
    #time.sleep(0.5)
    pi.write('1')#Acão
    #time.sleep(0.5)
    pi.hotkey('enter')
      
    time.sleep(0.5)
    
    
    return resp 
    
    
    
   
def anula_pagamento():
    
    
    
    return

if __name__ == '__main__':
    
    time.sleep(3)
    
    pi.FAILSAFE = True
        
    pi.PAUSE=0.3
    
    oper="Registra Pagamento"
    #oper="Anula Pagamento"
    
    if oper=="Registra Pagamento":
        #file=r"C:\Users\m1371121\Desktop\siafi\registra_pagamento.xlsx"
        #file=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\03_2023\PT pagamento SIAFI\registra_pagamento.xlsx"
        file=r'C:\Users\m1379117\Desktop\registra.xlsx'
        file_out=file
        #file_out=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_1\resposta_registra_pagamento.xlsx"
    
    if oper=="Anula Pagamento":
        #file=r"C:\Users\m1371121\Desktop\siafi\anula_pagamento.xlsx"
        file=r'C:\Users\m1379117\Desktop\registra.xlsx'
        file_out=file
        #file_out=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_1\resposta_anula_pagamento.xlsx"
    
    
    
    #local="accer_1"
    local="itautec"
    #local="robo-server"
    
    #list_wins()
    #list_class(name)
    #abre_siafi_run(local)
    #import sys
    
    import nest_asyncio
    nest_asyncio.apply()
    
    
    df=pd.read_excel(file)
    df=df.fillna("X")
    df=df.round(decimals=2)
    
    #df['Resultado'] = pd.Series(dtype='str')    
    
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
    ir=column_names[17] # valor irrf
    o=column_names[23] #observação
    rr=column_names[24] # resultado
    
    print(c, b, a, dva, ct, dvct, v, r, o, rr)
    
    types={s:str , p:str,  c:str,  n:str, b:str, a:str, dva:str, ct:str, dvct:str, v:str, r:str, o:str, rr:str}
    
    df2=df.astype(types)
    print(df2.dtypes) 
   
    
    ue="1080012" # A unidade executora para o Dativo Administrativo é 1080012.
    
    # Yan
    #login="m1371121"
    #senha="ym2023"
    
    # Leopoldo
    login="m1367336"
    senha="ps22024"
    
    
    if login=="xxxx" or senha=="xxxx":
        #ev=pi.confirm(text='Registe seu login e senha do SAIFI,\nnas linhas 1161 e 1162 do programa,\nantes de prosseguir', title='', buttons=['OK', 'Cancel'])
        ev=pi.alert(text='Registe seu login e senha do SAIFI,\nnas linhas 1116 e 1117 do programa,\nantes de prosseguir', title='')
        sys.exit()
    
    local="itautec"
    
    #controle_siafi contains abre_siafi e loga_siafi
    ano='2023'
    resp=asyncio.run(controle_siafi(login, senha, ue, local, ano)) 
    
    print("de volta a ===>  __name__ == '__main__':")
    
    
        
     ### LIQUIDACAO ###
    # Unidadae executora
    #ue="1080012" # A unidade executora para o Dativo Administrativo é 1080012.
    #unidade orcamentaria = uo
    uo="1081"
    #ordenador de despesa = od
    od="1120530" # Karen PTPT
    # Prog trabalho 
    p_trab ='10140001' 
    # #Nat. Despesa
    nat_desp ='339036'
    # Item
    item = '21'
    # Fonte. Proc/IPE
    fonte = '1010'
    
    st22=preparo_inicial()
    if st22=="TRANSMISSAO AOS BANCOS EM PROCESSAMENTO":
        pi.confirm(text='TRANSMISSAO AOS BANCOS EM PROCESSAMENTO\nAguarde alguns minutos antes de continuar\nFalta automatizar o reinicio deste script após 5 minutos.\nO programa será encerrado', title='', buttons=['OK', 'Cancel'])
        sys.exit()
    
    print("INICIO do loop")
    x=len(df)
    print(x)
    k=0
   
    for i in range(x):
        
        
        print(f'====================================Inicio da linha da planilha ==> {i+2} ================================================')
           
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
        
        
        sei=df2.iloc[i,1]# sei
        sei=trata_sei_cpha(sei)
        
        
        banco=df2.iloc[i,6]# banco
        banco=str(banco)
              
        agencia=df2.iloc[i,7]# agencia
        agencia=str(agencia)
       
        dv_a=df2.iloc[i,8]# dv agencia
        dv_a=str(dv_a)
       
        conta=df2.iloc[i,9]# conta
        conta=str(conta)
       
        dv_c=df2.iloc[i,10]# dv conta
        dv_c=str(dv_c)
        
        irr=df2.iloc[i,17]# irr
        irr=trata_irr(irr)
        irr=trata_valor_4(irr)
       
        obs=df2.iloc[i,23]# observacao
        obs=obs.strip()
       
       
        #pg=paga_dativo(emp, uo, tipo1, cpf_cnpj, banco, agencia, dv_a,  conta,  dv_c, valor, od, inss, sei):
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
            
        print(tipo)
        print(f'obs ==> {obs}')
        #print(xxxxx)
        
        
        ct2=""
        if oper=="Anula Pagamento":
            ######################    
            #CANCELA LIQUIDACAO
            ######################
            
            pos23=obs.find("-")
            pos24=obs.find("NAO")
            pos25=obs.find("NOK")
          
            pos26=obs.find("PAGAMENTO")
            
            pos27=obs.find("LIQUIDACAO REALIZADA COM SUCESSO")
            pos23=obs.find("-")
            pos28=obs.find("SUCESSO")
            
            ob1="PAGAMENTO NAO REALIZADO, EMPENHO"
            
            if (pos23>=0 and pos24<0 and pos25<0) or obs=="nan" or obs=="X" or obs=="x" or obs=="REPETIR":
                liq_num=obs[pos23+1:]
                liq_num=liq_num.strip()
                emp=liq_num
                pg=anula_pagamento(emp,uo,tipo,valor,od)   
            else:
                pg=obs
                ct2="ja executado"
                #pg=f"NAO FOI POSSIVEL CANCELAR PAGAMENTO: {obs}"
                pi.hotkey('f3')
            
        
        if oper=="Registra Pagamento":
            ######################    
            #EXECUTA LIQUIDACAO
            ######################
            print('Pagamento...')
            
            pos23=obs.find("-")
            pos24=obs.find("NAO")
            pos25=obs.find("NOK")
            
            
            pos26=obs.find("PAGAMENTO")
            
            pos27=obs.find("LIQUIDACAO REALIZADA COM SUCESSO")
          
            pos28=obs.find("SUCESSO")
            
            pos29=obs.find("INEXISTENTE")
            pos32=obs.find("FALTA")
            if pos29>0 or pos32>0:
                pos29=1
            
            pos30=obs.find("BANCO")
            
            pos31=obs.find("RESTO A PAGAR NAO DISPONIVEL")
            
            
            
            ob1="PAGAMENTO NAO REALIZADO, EMPENHO"
            
            
            #PAGAMENTO NAO REALIZADO, EMPENHO-15852, BANCO, AGENCIA OU CONTA INEXISTENTE 341 / 876 / 4 / 000000000507 / 3
            
            aa=bb=cc=dd=ee=ff=gg=hh=ii=jj="False"
            
            # Tem "PAGAMENTO NAO REALIZADO, EMPENHO"
            # Não tem "INEXISTENTE"
            if (ob1 in obs and pos29<0):
                if pos31>0:
                    print('Resto a pagar nao disponivel')
                else:
                    print("Executado pelo criterio A")
                    aa="True"
                
            # Não tem "PAGAMENTO", Não tem "SUCESSO", Não tem "INEXISTENTE"
            # Tem "LIQUIDACAO REALIZADA COM SUCESSO"
            if (pos26<0 and pos28<0 and pos29<0 and pos27>=0): 
                print("Executado pelo criterio B")
                bb="True"
            
                # Não tem "NAO" , Não tem "NOK",  Não tem "INEXISTENTE",
                # Tem "LIQUIDACAO REALIZADA COM SUCESSO", Tem "-"
            if (pos27>=0 and pos23>=0  and pos24<0 and pos25<0 and pos29<0):
                print("Executado pelo criterio C")
                cc="True"
            
            # Não tem "INEXISTENTE", 
            # Tem "LIQUIDACAO REALIZADA COM SUCESSO" , Tem "NAO"
            if (pos27>=0 and pos24>=0 and pos29<0):
                print("Executado pelo criterio D")
                dd="True"
            
            # Não tem "INEXISTENTE" 
            # Tem "LIQUIDACAO REALIZADA COM SUCESSO", Tem "NOK"
            if (pos27>=0 and pos25>0 and pos29<0):
                print("Executado pelo criterio E")
                ee="True"
                
            ########################################################
            if obs=="nan":
                print("Executado pelo criterio F")
                ff="True"
                ff="False"
                
            if obs=="X": 
                print("Executado pelo criterio G")
                gg="True"
                gg="False"
                
            if obs=="x":
                print("Executado pelo criterio H")
                hh="True"
                hh="False"
            
            if obs=="REPETIR":
                print("Executado pelo criterio J")
                ii="True"
                ii="False"
            ################################################################
            
            
            # EXECUTA PAGAMENTO
            print(aa,bb,cc,dd,ee,ff,gg,hh,ii) 
            if aa=="True" or bb=="True" or cc=="True" or dd=="True" or ee=="True" or ff=="True" or gg=="True" or hh=="True" or ii=="True":
            #if (ob1 in obs and pos29<0) or (pos26<0 and pos28<0 and pos29<0) or (pos27>=0 and pos23>=0  and pos24<0 and pos25<0 and pos29<0) or (pos27>=0 and pos24>=0 and pos29<0) or (pos27>=0 and pos25>0 and pos29<0) or obs=="nan" or obs=="X" or obs=="x" or obs=="REPETIR":
                
                # PAGAMENTO NAO REALIZADO, EMPENHO-11352 , Perda de Sincronismo
                # PAGAMENTO NAO REALIZADO, EMPENHO-11432 , fgadsfgadfgdafga
                # LIQUIDACAO REALIZADA COM SUCESSO EMPENHO-9950
                # PAGAMENTO NAO REALIZADO, EMPENHO-9950
                # PAGAMENTO NAO REALIZADO, EMPENHO-7749, BANCO, AGENCIA OU CONTA INEXISTENTE 001 / 607 / 6 / 000000001733 / 8
                #else:
                    
                num_emp=obs[pos23+1:]
                if "," in num_emp:
                    pos233=num_emp.find(",")
                    num_emp=num_emp[:pos233]
                
                num_emp=num_emp.strip()
                print(num_emp)
                
                #ev=pi.confirm(text=f'Verifique se numero do empenho esta correto\nEmpenho = {num_emp}\nlinha da planilha = {i+2}' , title='', buttons=['OK', 'Cancel'])
                #if ev=='Cancel':
                    #sys.exit()
                
                #liq=paga_dativo(num_emp,uo,tipo,valor,od, inss)
                pg=paga_dativo(num_emp, uo, tipo, cpf_cnpj, banco, agencia, dv_a,  conta,  dv_c, valor, od, inss, irr, sei, processo, i)

            else:
                print("Nao Executado por nao atender nenhum criterio")
                pg=obs
                ct2="ja executado"
                #pg=f"NAO FOI POSSIVEL REALIZAR PAGAMENTO :{obs}"
                #pi.hotkey('f3')
         
            
         
        # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
        # SE OPERACAO NÃO REALIZADA, NÃO ATUALIZA DF2, NEM SALVA EXCEL
        
        if ct2!="ja executado": 
         
        
            print("Change a cell value in dataframe")
            """
            df2.loc[s:s,i:i]= f"'{sei}"# coluna sei / cpha
            df2.loc[c:c,i:i]= f"'{cpf_cnpj}" # coluna cpf
            df2.loc[p:p,i:i]= f"'{processo}" # processo
            df2.loc[o:o,i:i)]= pg
            
            """
            df2[s][i] = f"'{sei}"# coluna sei / cpha
            df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
            df2[p][i] = f"'{processo}" # processo
            df2[o][i] = pg # coluna observacoes
            
            
            if "SUCESSO" in pg:
                df2[rr][i]="OPERACAO REALIZADA COM SUCESSO"
            else:
                if "PAGAMENTO NAO REALIZADO" in pg:
                    rr2="PAGAMENTO NAO REALIZADO"
                    df2[rr][i]=rr2
            
            #print(df2)
            #print todas linhas e colunas c,v,o only
            
            
            #print(df2.loc[:,[c,v,o]])
            
            """
            k+=1
            if k==10:
                k=0
                print(df2)
            """
            
            df2.to_excel(file_out , index=False)
            
            
            #ev=pi.confirm(text=f'linha 1652 do programa:\nVerifique mensagem na planilha, linha da planilha {i+2}\n{pg}' , title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
        
    
    pi.alert(text='---> FIM  PAGAMENTO  <---', title='', button='OK')
