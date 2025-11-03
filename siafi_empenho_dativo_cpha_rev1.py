# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 09:40:43 2022

@author: Usuario m1379117
release 11
"""


"""EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO - CPHA"""


# from asgiref.sync import sync_to_async


import asyncio
import os
import re
import subprocess
import sys
import time
import pandas as pd
import pyautogui as pi
import pygetwindow as gw
import pyperclip
tit = 'pw3270'


def active_maximize(tit):
    print("INICIO da funcao active_maximize")
    time.sleep(1)
    x = gw.getAllTitles()
    print(x)
    if tit in x:
        print(f'OK - Localizada janela com titulo {tit} do SAIFI')
        janela = gw.getWindowsWithTitle(tit)[0]
        janela.activate()
        # time.sleep(2)
        janela.maximize()
        # time.sleep(0.5)
        return "OK"
    else:
        print(f'Nao localizada janela com titulo {tit}')
        return "NOK"


def read_excel(file, tipo):

    # Read excel and returns a data frame)
    # tipo =1 planilha dativos
    # tipo =  2  planilha gerais

    df = pd.read_excel(file)

    column_names = list(df.columns)

    print(column_names)

    print(type(column_names))

    if tipo == "1":  # Planilha Dativos
        # returns only the columns index 3 and 4
        # df2 = df.iloc[:,[3,4]] # Select columns by Index 3 and
        # Select columns by Index 4 and 19 = cpf e valo liquido na planilha dativos
        df2 = df.iloc[:, [4, 18]]

    if tipo == "2":  # Planilha Gerais
        # Returns all excel df
        # df2=df
        print("TBD")

    # Adiciona uma coluna vazia

    print(df2)
    return df2


def cria_df3(column_names):
    # retorna df vazio apenas com o nome das colunas a serem utilizadas

    print(column_names)

    df1 = pd.DataFrame(columns=column_names)

    return df1


def abre_siafi_run(local):

    print("INICIO da funcao  abre_siafi_run Empenho")

    # import sys
    print(os.getcwd())

    if local == "itautec":
        dd = r"C:\Program Files (x86)\pw3270"
        # "C:\Program Files (x86)\pw3270\pw3270.exe" --host=bhmvsb.prodemge.gov.br
        # b=[r"C:\Program Files (x86)\pw3270\pw3270.exe"]#, "--host=bhmvsb.prodemge.gov.br"]
        # b=[r'"C:\Program Files (x86)\pw3270\pw3270.exe", --host=bhmvsb.prodemge.gov.br']
        b = r'"C:\Program Files (x86)\pw3270\pw3270.exe" --host=bhmvsb.prodemge.gov.br'

    if local == "accer_1":
        dd = r"C:\Program Files\pw3270"
        # b=[r"C:\Program Files\pw3270\pw3270.exe"]# "--host=bhmvsb.prodemge.gov.br"]
        # b=[r'"C:\Program Files (x86)\pw3270\pw3270.exe", --host=bhmvsb.prodemge.gov.br']
        b = r'"C:\Program Files (x86)\pw3270\pw3270.exe" --host=bhmvsb.prodemge.gov.br'

    if local == "robo-server":
        print("To be installed")

    os.chdir(dd)

    print(os.getcwd())

    subprocess.run(b)

    # time.sleep(1)

    msg = "OK from abre_siafi_run"
    print(msg)

    return


def loga_siafi(login, senha, ue, local, ano):
    print("INICIO da funcao loga_siafi")

    msg1 = "Logon executado com sucesso"
    msg2 = "Unidade Executora:"

    import time

    import pyautogui as pi

    # time.sleep(1)
    for i in range(3):
        rr = active_maximize(tit)
        # time.sleep(1)
        if rr == "OK":
            break

    # pi.getWindowsWithTitle(resp)[0].maximize()

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

    # for i in range(25):
    # pi.hotkey('right')
    # time.sleep(0.1)

    print("3")
    msg1 = "Aplicacao:"
    st = 1
    n = 5
    rsp = check_current_screen(msg1, n, st)
    if rsp == "OK":
        print("Tela de autenticação aberta com sucesso")

    else:
        print("Tela de autenticação indisponivel")
        resp = "NOK-0 from loga_siafi autenticacao"
        print(resp)

        return resp

    pi.write('simg')
    pi.hotkey('tab')
    pi.write(login)
    pi.write(senha)
    pi.hotkey('enter')
    time.sleep(1)

    msg1 = "Logon executado com sucesso"
    for i in range(10):
        time.sleep(1.5)
        pi.hotkey('ctrl', 'a')
        pi.hotkey('alt', 'e')
        pi.hotkey('enter')
        # pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        # print(x)
        pos1 = x.find(msg1)
        if pos1 >= 0:
            break
        pi.hotkey('enter')

    time.sleep(2)
    pi.write('simg')
    time.sleep(0.5)
    pi.hotkey('enter')
    time.sleep(1)

    a, b = pi.size()
    # pi.click(x=a/2, y=b/2)
    # time.sleep(1)

    for i in range(5):
        time.sleep(0.5)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pi.hotkey('alt', 'e')
        pi.hotkey('enter')
        # pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        # print(x)
        pos1 = x.find(msg2)
        if pos1 >= 0:
            break
        pi.hotkey('enter')

    # pi.hotkey('enter')
    pi.hotkey('esc')
    pi.hotkey('tab')

    # time.sleep(1)
    pi.write(ue)  # uidade executora 10080012
    time.sleep(0.5)
    pi.write(ano)

    """
    from datetime import datetime
    now = datetime.now()
    ano = now.strftime("%Y")
    print(ano)
    pi.write(ano)#ano
    """

    # time.sleep(1)
    pi.hotkey('enter')
    time.sleep(2)

    msg1 = "01-Rotina Administrativa"
    n = 10
    st = 1
    rsp = check_current_screen(msg1, n, st)
    if rsp == "NOK":
        resp = "NOK-1 from loga_siafi"
        pi.hotkey('esc')
        time.sleep(1)
        rsp = check_current_screen(msg1, n, st)
        if rsp == "NOK":
            resp = "NOK-2 from loga_siafi"
        resp = "OK-2 from loga_siafi"
    else:
        resp = "OK-1 from loga_siafi-1"

    print(resp)

    return resp


def get_emp_number(n, st):
    """ Busca numero do empenho na tela"""

    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int

    time.sleep(1)
    msg1 = "REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL"

    # Click no centro da tela
    a, b = pi.size()
    pi.click(x=a/2, y=b/2)
    # pi.hotkey('tab')#Opção

    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl', 'a')
        pi.hotkey('alt', 'e')
        pi.hotkey('enter')
        # pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        # print(x)
        pos1 = x.find(msg1)
        if pos1 >= 0:
            print(f' localizado {msg1}')
            resp = "OK"
            break
        else:
            resp = "NOK"
            print(f' ciclo {i+1} não localizado {msg1}')
            return resp

        i += 1
        if i+1 == n:
            resp = "NOK"
            return resp

    t1 = "Nr. Documento:"  # 6386
    t2 = "Evento"

    pos1 = x.find(t1)
    pos2 = x.find(t2)

    emp = x[pos1+14:pos2]
    emp = emp.strip()

    resp = emp

    print(f'Numero do Empenho = {resp}')
    return resp


def get_first_line(qte, n, st):
    """get n first character in the currente screen"""
    # qte = characters qte to get : int
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int

    # Click no centro da tela
    a, b = pi.size()
    pi.click(x=a/2, y=b/2)
    # pi.hotkey('tab')#Opção

    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl', 'a')
        pi.hotkey('alt', 'e')
        pi.hotkey('enter')
        # pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        # print(x)

    x = x.strip()
    line = x[:qte]

    return line


def check_current_screen(msg1, n, st):
    """try to find msg in the currente screen"""
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int

    # Click no centro da tela
    a, b = pi.size()
    pi.click(x=a/2, y=b/2)
    # pi.hotkey('tab')#Opção

    for i in range(n):
        time.sleep(st)
        pi.click(x=a/2, y=b/2)
        pi.hotkey('ctrl', 'a')
        pi.hotkey('alt', 'e')
        pi.hotkey('enter')
        # pyperclip.copy('The text to be copied to the clipboard.')
        x = pyperclip.paste()
        # print(x)
        pos1 = x.find(msg1)
        if pos1 >= 0:
            print(f' localizado {msg1}')
            resp = "OK"
            break
        else:
            resp = "NOK"
            print(f' ciclo {i+1} não localizado {msg1}')

        i += 1
        if i+1 == n:
            resp = "NOK"

    # pi.hotkey('esc') #### Linha em teste
    pi.hotkey('tab')
    print(resp)
    return resp


def mouse_position():
    import time

    import pyautogui as pi

    time.sleep(3)

    x = pi.position()

    print(x)

    return


def ctl_keys_itautec_17(pf):

    # pf = 1,2,3,4,5,6,7,8,9,10,11,12

    arr = ['Point(x=1259, y=129)',
           'Point(x=1302, y=128)',
           'Point(x=1347, y=125)',
           'Point(x=1256, y=167)',
           'Point(x=1305, y=176)',
           'Point(x=1336, y=171)',
           'Point(x=1262, y=216)',
           'Point(x=1306, y=215)',
           'Point(x=1345, y=214)',
           'Point(x=1261, y=259)',
           'Point(x=1304, y=262)',
           'Point(x=1341, y=262)']

    tecla = arr[pf-1]
    tecla = tecla.strip()
    print(tecla)

    x = tecla[8:12]
    y = tecla[16:-1]

    resp = [x, y]
    print(resp)

    return resp


def trata_sei_cpha(sei):

    print("INICIO da funcao trata_sei")

    sei = str(sei)
    sei = sei.strip()
    sei = sei.replace("'", "")

    pos1 = sei.find(".")
    if pos1 >= 0:
        sei = sei[:pos1]
        sei = sei.strip()

    sei1 = re.sub("\D+", "", sei)
    x = len(sei1)
    if sei1[:4] != "1080":  # cpha
        if x < 20:
            f = 20-x
            for j in range(f):
                sei = f'0{sei}'
                print(sei)

    print(f'sei corrigido ==> {sei}')

    return sei


def trata_cpf_cnpj(cpf_cnpj):

    cpf_cnpj = str(cpf_cnpj)
    cpf_cnpj = cpf_cnpj.strip()
    cpf_cnpj = cpf_cnpj.replace("'", "")

    tipo = "cpf"
    if "/" in cpf_cnpj:  # cnpj
        tipo = "cnpj"
        print(tipo)
        cpf_cnpj = re.sub("\D+", "", cpf_cnpj)
        print(cpf_cnpj)
        x = len(cpf_cnpj)

        # Completa ate atingir 11 digitos para CNPJ
        if x < 14:
            f = 14-x
            for j in range(f):
                cpf_cnpj = f'0{cpf_cnpj}'
                print(cpf_cnpj)

    else:  # CPF
        print(tipo)
        cpf_cnpj = re.sub("\D+", "", cpf_cnpj)
        print(cpf_cnpj)
        x = len(cpf_cnpj)

        # Completa ate atingir 11 digitos para CPF
        if x < 11:
            f = 11-x
            for j in range(f):
                cpf_cnpj = f'0{cpf_cnpj}'
                print(cpf_cnpj)

    print(f'cpf_cnpj corrigido ==> {cpf_cnpj}')

    return cpf_cnpj


def trata_processo(processo):
    print("INICIO da funcao trata_processo")

    processo = str(processo)
    processo = processo.strip()
    processo = processo.replace("'", "")

    pos1 = processo.find(".")
    if pos1 >= 0:
        processo = processo[:pos1]
        processo = processo.strip()

    processo1 = re.sub("\D+", "", processo)
    x = len(processo1)
    if x < 20:
        f = 20-x
        for j in range(f):
            processo = f'0{processo}'
            print(processo)

    print(f'processo corrigido ==> {processo}')

    return processo


def trata_valor(valor):
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


def get_val_anula(n, st):
    # msg1 = searched text on the screen: str
    # n = number tries: int
    # sleep time betwwen tries (cicles) : int
    
    print("INICIO da funcao get_val_anula")
    
    msg1="Saldo Empenho:"
    
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
            resp="FALHA NO SINCRONISMO"
            print(f' ciclo {i+1} não localizado {msg1}')
            return resp
            
        i+=1
        if i+1==n:
            resp="FALHA NO SINCRONISMO"
            return resp
    
    
    #rp=re.findall("\d+[,]\d+", x)
    rp=re.findall("[\d+.]*[,]\d+", x)
    print(rp)
    try:
        pos=len(rp)
        resp=rp[pos-1]
    except:
        resp='0.00'
        
    resp=re.sub('[^0-9,]', '', resp)
   
        
    print(f' Valor = {resp}')
    
    #ops=pi.confirm(text=f'Verifique a valor a ser anulado = {resp}', title='', buttons=['OK', 'Cancel'])
    #if ops=='Cancel':
        #sys.exit()
    resp=trata_valor_4(resp)
    
    return resp





def anula_empenho(uo, emp, tipo1, valor, od):#, valor, od):
    print("INICIO da funcao anula_empenho")

    #print(uo, emp, tipo1, valor, od)
 

    pi.write(uo)
    pi.write(emp)
    pi.hotkey('enter')
    time.sleep(1)
    print(f'valor==>{valor}')

    pi.PAUSE = 0.5
    
    n=2
    st=0.5
    valor=get_val_anula(n, st)
    pi.hotkey('home')
    
    if valor == "FALHA NO SINCRONISMO":
        return f"FALHA NO SINCRONISMO-{emp}"
    if valor == "000":
        
        pi.hotkey('f3')
        pi.write('02')
        pi.write('1')
        pi.hotkey('enter')        
        
        return f"SALDO INEXISTENTE PARA ANULACAO DE EMPENHO-{emp}"
    
    pi.write(valor) # VErificar se é possivel fazer  get value na tela
    pi.hotkey('tab')
    pi.write(od)
    
    #ev=pi.confirm(text=f'Verifique se é conseguiu fazer get_value = {valor}', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
    
    pi.hotkey('enter')

    msg1 = "VALOR ANULACAO MAIOR QUE SALDO EMPENHO"
    mm = check_current_screen(msg1, 2, 0.5)
    if mm == "OK":
        resp = f"VALOR ANULACAO EMPENHO MAIOR QUE SALDO-{emp}"
        print(resp)

        # Retorna Tela Inicial
        pi.hotkey('f3')
        pi.write('02')  # Opção
        # time.sleep(0.5)
        pi.write('1')  # Acão
        # time.sleep(0.5)
        pi.hotkey('enter')
        time.sleep(0.5)

        return resp

    # ev=pi.confirm(text='Verifique qual tecla deve ser clicada \n previsto down (x2), X e F5', title='', buttons=['OK', 'Cancel'])
    # if ev=='Cancel':
        # sys.exit()

    # pi.hotkey('down')
    # pi.hotkey('down')
    pi.write('X')
    pi.hotkey('f5')

    pi.PAUSE = 1

    pi.write(f'REGISTRA ANULACAO EMPENHO-{emp}')
    
    #ev=pi.confirm(text='Verifique tecla F5 apos anulacao/n para inicio de novo ciclo/N/N Incluir check_screen para validar Sucesso', title='', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
        #sys.exit()
        
    pi.hotkey('f5')
    pi.hotkey('f5')
    
    
    n=2
    st=0.5
    msg1='REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL'
    ver=check_current_screen(msg1, n, st)
    
    if ver=="OK":
        resp = f"ANULACAO DE EMPENHO REALIZADO COM SUCESSO-{emp}"
        # Retorna Tela Inicial
        pi.hotkey('enter')
        pi.PAUSE = 0.4
    else:
        resp=f'ANULACAO DE EMPENHO NAO REALIZADA-{emp}'
        
    print(resp)
    
    return resp

def registra_empenho(cpf_cnpj, valor, uo, od, ptrab, nat_desp, item, fonte, tipo1, sei, nome, nit):

    print("INICIO da funcao rgistra_empenho")
    # cpf_cnpj=cpf_cnpj do credor - somente numeros oui nao
    # valor = valor tipo 2000 equivale a 20 reais = confirmado OK
    # uo = unidade orcamentaria =  "1081"
    # od = ordenador de despesa = "10606119"
    # ptrab = Prog trabalho  = '10140001'
    # nat_desp = Natureza Despesa '339036'
    # item = Item = '16'
    # fonte = Fonte. Proc/IPE = '1090'
    # tipo = inicial, outros
    # sei = numero do SEI

    at = ""

    x = len(cpf_cnpj)
    if x == 11:
        tipo = "cpf"
    if x == 14:
        tipo = "cnpj"
    if x != 11 and x != 14:
        at = f"cpf_cnpj invalido ==> {cpf_cnpj}"
        # return resp

    # Verifica se 2a telinha aberta com sucesso
    time.sleep(2)
    msg = "Programacao de Empenho"
    n = 3
    st = 1
    rsp = check_current_screen(msg, 1, 0.5)
    if rsp == "NOK":
        resp = "REPETIR"
        print("repetir empenho")
        pi.hotkey('f3')
        # clica_pf(3)
        # time.sleep(1)
        return resp

    pi.hotkey('shift', 'tab')
    time.sleep(1) 	
    pi.hotkey('shift', 'tab')

    # ev=pi.confirm(text='Verifique qual tecla deve ser clicada para 2a telinha', title='', buttons=['OK', 'Cancel'])
    # if ev=='Cancel':
    # sys.exit()

    # print(xxx)

    # 2a TELINHA
    time.sleep(1)
    pi.write(ptrab)  # Prog trabalho
    time.sleep(1)
    pi.write(nat_desp)  # Nat. Despesa
    time.sleep(.5)
    pi.write(item)  # Item
    time.sleep(.5)
    pi.write(fonte)  # Fonte. Proc/IPE
    time.sleep(.5)

    if "invalido" in at:
        resp = at
        print(resp)
        return resp

    pi.write(cpf_cnpj)  # CPF CAP Credor
    pi.write('E')  # Tipo empenho
    pi.hotkey('down')  # Num Especificação
    pi.hotkey('down')  # Valor

    # Historico

    pi.hotkey('f5')
    pi.hotkey('f5')
    # clica_pf(5)
    # clica_pf(5)

    """ Registrar Apropriacao de Empenho  """
    # valor=float(valor)
    pi.write(valor)  # valor
    time.sleep(0.5)
    pi.hotkey('tab')  # valor
    pi.write(od)  # ordenador de despesas
    # time.sleep(2)

    pi.hotkey('enter')

    time.sleep(1)

    # ev=pi.confirm(text=' Verifique se abriu tela Historico - APROPRIACAO DE EMPENHO DA DESPESA', title='', buttons=['OK', 'Cancel'])
    # if ev=='Cancel':
    # sys.exit()

    """Historico - APROPRIACAO DE EMPENHO DA DESPESA """

    msg = "Historico - APROPRIACAO DE EMPENHO DA DESPESA"
    ck = check_current_screen(msg, 1, 0.5)
    if ck != "OK":

        # VERIFICA SE HA SALDO
        msg = "SALDO DE COTA DESCENT.INSUFICIENTE"
        ck = check_current_screen(msg, 1, 0.5)
        if ck == "OK":
            resp = msg
            pi.hotkey('f3')
            return resp
        

        # VERIFICA SE CREDOR CADASTRADO
        
        msg="CREDOR INEXISTENTE NO SIAD"
        ck = check_current_screen(msg, 1, 0.5)
        if ck == "OK":
            resp = msg
            pi.hotkey('f3')
            return resp

        msg = "INFORME CORRETAMENTE CREDOR OU TECLE PF4 PARA INCLUIR NOVO CREDOR"
        ck = check_current_screen(msg, 1, 1)
        if ck == "OK":
            resp = msg
            ###########################
            pi.hotkey('f4')

            # print(xxxx)
            pi.write('13')
            # inclue_credor(nome)
            # pi.hotkey('tab')
            pi.write(nome)
            pi.hotkey('tab')
            pi.hotkey('tab')
            pi.hotkey('tab')

            # pi.write(cpf_cnpj)
            pi.write('BELO HORIZONTE')
            pi.hotkey('tab')
            pi.write('MG')
            pi.hotkey('f6')
            pi.hotkey('down')
            pi.hotkey('down')
            pi.hotkey('down')
            pi.write('x')

            # ev=pi.confirm(text='Verifique valores inseridos na tela de inclusao novo credor / cpf;\n\nProxima tecla será previsto ENTER', title='', buttons=['OK', 'Cancel'])
            # if ev=='Cancel':
            # sys.exit()

            pi.hotkey('enter')
            pi.hotkey('f5')
            pi.hotkey('f5')
            # checar  0012-INCLUSAO EFETUADA.

            line = get_first_line(36, 1, 1)

            if "CPF" in line:
                resp = "CPF invalido"
                pi.hotkey('f3')
                pi.hotkey('f3')

                return resp

            if "JA CADASTRADO NA TABELA" in line:
                resp = "CREDOR JA CADASTRADO NA TABELA"
                pi.hotkey('f3')
                pi.hotkey('f3')

                return resp

            if ("FIELD" in line) or ("field" in line):
                resp = "CREDOR invalido"
                pi.hotkey('f3')
                pi.hotkey('f3')

                return resp

            # ck=check_current_screen(msg, 1, 0.5)
            # if ck=="OK":

            # ev=pi.confirm(text='Verifique tela atual após inclusao de novo credor/cpf com sucesso;\n\nProxima tecla será previsto F3', title='', buttons=['OK', 'Cancel'])
            # if ev=='Cancel':
                # sys.exit()

            pi.hotkey('f3')
            pi.write(nit)
            pi.hotkey('enter')

            ###############################
            # ev=pi.confirm(text='Verifique valores inseridos na tela de inclusao novo credor / cpf;\n\nProxima tecla será previsto ENTER', title='', buttons=['OK', 'Cancel'])
            # if ev=='Cancel':
            # sys.exit()
            ##################################

            msg = "DIGITO INVALIDO DO NIT/PIS/PASEP"
            ck = check_current_screen(msg, 1, 0.5)
            if ck == "OK":
                resp = msg
                pi.hotkey('f3')
                # pi.hotkey('f3')

                return resp

            pi.hotkey('down')
            pi.hotkey('down')
            pi.write('x')  # tela de selecao
            pi.hotkey('f5')
            if "1080" in sei:
                msg5 = 'VALOR DE EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO'
                msg6 = f'SEI {sei}'
            else:
                msg5 = 'VALOR DE EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO'
                msg6 = f'CPHA {sei}'
            time.sleep(0.5)
            pi.write(msg5)
            time.sleep(0.5)
            pi.hotkey('tab')
            time.sleep(0.5)
            pi.write(msg6)
            time.sleep(0.5)
            pi.hotkey('f5')
            pi.hotkey('f5')
            # Busca numero do bloqueio

            emp1 = get_emp_number(2, 1)
            print(emp1)

            # Tecle Enter para enviar para assinatura Digital
            pi.hotkey('enter')
            time.sleep(0.5)
            pi.hotkey('f3')

            if emp1 == "NOK":
                resp = f'EMPENHO DA DESPESA NAO REALIZADO-{emp1}'
            else:
                resp = f'EMPENHO DA DESPESA REALIZADO COM SUCESSO-{emp1}'
                # Retorna telas anteriores
                # pi.hotkey('f3')
                # Retorna resposta
            print(resp)
            return resp
            ##########################
            # RETORNA TELA INICIAL
            # pi.hotkey('f3')
            # pi.hotkey('f3')
            # return resp

        # VERIFICA NECESSIDADE DE NIT
        msg = "PREENCHER NIT OU PIS/PASEP"
        ck = check_current_screen(msg, 1, 0.5)
        if ck == "OK":
            resp = msg
            pi.write(nit)

            pi.hotkey('enter')

            msg = "NIT/PIS/PASEP JA INFORMADO PARA CREDOR"
            ck = check_current_screen(msg, 1, 0.5)
            if ck == "OK":
                pi.hotkey('f3')
                resp = msg
                return resp

            msg = "DIGITO INVALIDO DO NIT/PIS/PASEP"
            ck = check_current_screen(msg, 1, 0.5)
            if ck == "OK":
                resp = msg
                pi.hotkey('f3')

                return resp

            pi.hotkey('down')
            pi.hotkey('down')
            pi.write('x')
            pi.hotkey('f5')
            if "1080" in sei:
                msg5 = 'VALOR DE EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO'
                msg6 = f'SEI {sei}'
            else:
                msg5 = 'VALOR DE EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO'
                msg6 = f'CPHA {sei}'
            time.sleep(0.5)
            pi.write(msg5)
            time.sleep(0.5)
            pi.hotkey('tab')
            time.sleep(0.5)
            pi.write(msg6)
            time.sleep(0.5)
            pi.hotkey('f5')
            pi.hotkey('f5')

            # Busca numero do bloqueio
            emp1 = get_emp_number(2, 1)
            print(emp1)

            # Tecle Enter para enviar para assinatura Digital
            pi.hotkey('enter')
            time.sleep(0.5)
            resp = f'EMPENHO DA DESPESA REALIZADO COM SUCESSO-{emp1}'
            # Retorna telas anteriores
            pi.hotkey('f3')
            # pi.hotkey('f3')
            # Retorna resposta
            print(resp)
            return resp
            ##########################

            # RETORNA TELA INICIAL
            # pi.hotkey('f3')
            # pi.hotkey('f3')
            # return resp

        msg = "Input value for a numeric field is not numeric"
        ck = check_current_screen(msg, 1, 0.5)
        if ck == "OK":
            time.sleep(1)
            pi.hotkey('esc')
            for i in range(20):
                pi.hotkey('del')

            ev = pi.confirm(text='Identificado -Input value for a numeric field is not numeric-\nVerifique tela atual\n Proxima tecla Previsto F3\n linha 965',
                            title='', buttons=['OK', 'Cancel'])
            if ev == 'Cancel':
                sys.exit()

            pi.hotkey('f3')
            print(msg)
            resp = 'Perda sincronismo-REPETIR'
            return resp

        # AJUSTE DE POSICAO PARA ATINGIR ==> "APROPRIACAO EMPENHO - OUTRAS DESPESAS CORRENTES "
        time.sleep(0.5)
        pi.hotkey('up')
        pi.hotkey('up')
        pi.hotkey('up')

    pi.write('X')  # Seleciona
    time.sleep(0.5)

    # ev=pi.confirm(text=' Verifique se foi selecionado\nAPROPRIACAO EMPENHO - OUTRAS DESPESAS CORRENTES', title='', buttons=['OK', 'Cancel'])
    # if ev=='Cancel':
    # sys.exit()

    pi.hotkey('f5')
    # clica_pf(5)

    print(sei, type(sei))
    if "1080" in sei:
        msg5 = 'VALOR DE EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO'
        msg6 = f'SEI {sei}'
    else:
        msg5 = 'VALOR DE EMPENHO PARA PAGAMENTO DE DATIVO ADMINISTRATIVO'
        msg6 = f'CPHA {sei}'
    time.sleep(0.5)
    pi.write(msg5)
    time.sleep(0.5)
    pi.hotkey('tab')
    time.sleep(0.5)
    pi.write(msg6)
    time.sleep(2)
    pi.hotkey('f5')
    pi.hotkey('f5')
    # clica_pf(5)
    # clica_pf(5)

    # Busca numero do bloqueio
    emp1 = get_emp_number(2, 1)
    print(emp1)

    # REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL
    msg1 = "REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL"
    n = 2
    st = 0.5
    rp = check_current_screen(msg1, n, st)

    if rp == "OK":
        resp = f'EMPENHO DA DESPESA REALIZADO COM SUCESSO-{emp1}'

    if rp == "NOK":
        resp = f'EMPENHO DA DESPESA NAO REALIZADO-{emp1}'

    #  REGISTRO EFETUADO, TECLE ENTER PARA ENVIAR PARA A ASSINATURA DIGITAL

    print(resp)

    # Tecle Enter para enviar para assinatura Digital
    pi.hotkey('enter')
    time.sleep(0.5)

    # ev=pi.confirm(text=f'Numero do empenho realizado = {emp1}\n linha 1176 ; valor = {valor}', title='', buttons=['OK', 'Cancel'])
    # if ev=='Cancel':
    #      sys.exit()

    # Retorna telas anteriores
    # retorna_tela_inic_empenho()
    pi.hotkey('f3')

    # print(xxxx)

    # pi.hotkey('f3')
    # clica_pf(3)
    # clica_pf(3)

    # print(resp)
    return resp


def inclue_credor(nome):
    """ESTADO INICIAL = TELA DE INCLUSAO CREDOR APOS TABLA , CREDOR , ETC..."""

    msg = "Incluir Credor/Devedor"
    ck = check_current_screen(msg, 1, 0.5)
    if ck == "OK":

        pi.PAUSE = 0.5
        pi.write('16')  # BENEFICIARIO DE EMPRESTIMO/FINANCIAMENTO-BDMG
        pi.write(nome)
        pi.hotkey('tab')
        pi.hotkey('tab')
        pi.hotkey('tab')
        pi.write('BELO HORIZONTE')
        pi.hotkey('tab')
        pi.write('MG')
        pi.hotkey('f5')
        pi.hotkey('f5')
        pi.hotkey('f3')  # volta tela de cadastro empenho
        resp = "INCLUSAO CREDOR EFETUADA COM SUCESSO"

    else:
        resp = "FALHA NA INCLUSAO CREDOR"

    print(resp)

    return resp


def retorna_tela_inic_empenho():
    # Retorna telas anteriores clicando duas vezes em PF3

    pf = ctl_keys_itautec_17(3)
    # Coordenadas PF3
    x1 = pf[0]
    y1 = pf[1]
    x1 = int(x1)
    y1 = int(y1)

    # Clica 2 x em PF3 para Voltar
    pi.click(x=x1, y=y1)
    time.sleep(1)
    pi.click(x=x1, y=y1)
    time.sleep(1)

    return


def clica_pf(x):

    # clica em PFn um avez (PF1, PF2, PF3,PF4,PF5 ....)

    # pi.press("f5") x 2
    pf = ctl_keys_itautec_17(x)
    # Coordenadas PF5
    x1 = pf[0]
    y1 = pf[1]
    x1 = int(x1)
    y1 = int(y1)
    # Clica 2 x em PF5
    pi.click(x=x1, y=y1)
    time.sleep(1)
    return


def quit_siafi(local):

    print("INICIO da funcao quit_siafi")

    import time

    import pyautogui as pi

    # if in english
    if (local == "accer_1" or local == "itautec"):
        pi.hotkey('alt', 'f')
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

    # time.sleep(1)

    async_function = sync_to_async(loga_siafi, thread_sensitive=False)
    task2 = loop.create_task(async_function(login, senha, ue, local, ano))
    print(f'task2 ====> {task2}')

    await task2

    ##############################################
    ##### Aqui deverá ter o loop da planilha #####
    # async_function = sync_to_async(controle_empenho, thread_sensitive=False)
    # task3=loop.create_task(async_function(file))
    # print(f'task3 ====> {task3}')
    # await task3
    #########################################

    if local == "accer_1":
        resp = "pw3270:A - tn3270://bhmvsb.prodemge.gov.br:23"

    if local == "itautec":
        resp = "pw3270 - bhmvsb.prodemge.gov.br"

    print(resp)

    return resp


def list_wins():

    import pyautogui as pi
    x = pi.getAllWindows()

    print(f"lista de janelas =====\/\/\/\/\/\/")
    for k, v in enumerate(x):
        print(k, v.title)

    return


def list_class(name):

    # name is the class name
    x = dir(name)
    for k, v in enumerate(x):
        # if "Win" in v:
        print(k, v)

    return


def preparo_inicial(oper):
    
    if oper=="Anula Empenho":
        preparo_inicial_anula()  # Opção Anula Empenho
        return

    tipo = "cpf"

    # Click no centro da tela
    a, b = pi.size()
    pi.click(x=a/2, y=b/2)

    pi.hotkey('tab')  # Opção
    # time.sleep(1)
    pi.write('51')  # Navegação 5=Movimentacao despesa - 1=execução
    # time.sleep(1)
    pi.hotkey('enter')
    # time.sleep(1)

        
    pi.write('01')  # Opção Apropriação de Empenho
    # time.sleep(0.5)
    
    pi.write('1')  # OPção Registrar
    # time.sleep(0.5)
    pi.hotkey('enter')

    time.sleep(0.5)

    # ev=pi.confirm(text=f'1a telinha do Empenho linha {i+1}...', title='', buttons=['OK', 'Cancel'])
  # if ev=='Cancel':
    # sys.exit()

    # 1a TELINHA
    # check_current_screen(msg1, n, st)

    pi.write('1081')  # Unidade Orcamentaria
    # time.sleep(0.5)
    pi.write('X')  # Empenho programado
    # time.sleep(0.5)
    pi.hotkey('down')  # Empenho
    # time.sleep(0.5)
    pi.hotkey('down')
    # time.sleep(0.5)
    if tipo == "cpf":
        pi.write('X')  # CPF/CAPF
        # time.sleep(0.5)
        pi.hotkey('down')  # CNPJ/CAPJ
    if tipo == "cnpj":
        pi.write('down')  # CPF/CAPF
        # time.sleep(0.5)
        pi.hotkey('X')  # CNPJ/CAPJ
    # time.sleep(0.5)
    pi.hotkey('down')  # Blank
    # time.sleep(1)
    pi.write('N')  # Adiantamento (S/N)
    # time.sleep(0.5)
    pi.write('N')  # Ressarciemnto (S/N)
    # time.sleep(0.5)
    pi.write('N')  # Contrato/Convenio
    # time.sleep(0.5)
    pi.write('N')  # Contrato saida
    # time.sleep(0.5)
    pi.write('N')  # Cadastro de Obras
    print("1a etapa concluida")
    #

    time.sleep(0.5)

    pi.hotkey('f6')
    pi.hotkey('f6')
    # clica_pf(6)
    # clica_pf(6)

    return


def preparo_inicial_anula():

    tipo = "cpf"

    # Click no centro da tela
    a, b = pi.size()
    pi.click(x=a/2, y=b/2)

    pi.hotkey('tab')  # Opção
    # time.sleep(1)
    pi.write('51')  # Navegação 5=Movimentacao despesa - 1=execução
    # time.sleep(1)
    pi.hotkey('enter')
    # time.sleep(1)

    pi.write('02')  # Opção Anula Empenho
    # time.sleep(0.5)
    pi.write('1')  # OPção Registrar
    # time.sleep(0.5)
    pi.hotkey('enter')

    time.sleep(0.5)

    #ev = pi.confirm(text='1a Telinha Anula Empenho Preenchida corretamente ?' , title='', buttons=['OK', 'Cancel'])
    #if ev == 'Cancel':
        #sys.exit()

    return


if __name__ == '__main__':

    time.sleep(3)

    # oper="Anula Empenho"
    oper = "Registra Empenho"

    if oper == "Registra Empenho":
        # file=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_3\registra_empenho.xlsx"
        # file=r'C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\02_2024\PT_pagamento_SIAFI_8000\planilha_entrada.xlsx'
        # file=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\03_2024\PT_empenho_siafi_1_a_2000\registra_empenho.xlsx"

        file = r"C:\Users\m1379117\Desktop\registra.xlsx"
        file_out = file
        # file_out=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_2\resposta_registra_empenho.xlsx"

    if oper == "Anula Empenho":
        file = r"C:\Users\m1379117\Desktop\registra.xlsx"
        file_out = file
        # file_out=r"C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\12-2022\dativos 3a listagem\lista_2\resposta_anula_empenho.xlsx"

    # list_wins()
    # list_class(name)
    # abre_siafi_run(local)
    import sys

    import nest_asyncio
    nest_asyncio.apply()

    df = pd.read_excel(file)
    df = df.fillna("x")
    df = df.round(decimals=2)

    # df['Resultado'] = pd.Series(dtype='str')

    print(df)

    # TRATA NOME DAS COLUNAS
    column_names = list(df.columns)

    s = column_names[1]  # sei / cpha
    p = column_names[2]  # processo
    c = column_names[4]  # cpf
    n = column_names[5]  # nit
    b = column_names[6]  # banco
    # aa=column_names[7] # agencia + digito
    a = column_names[7]  # agencia
    dva = column_names[8]  # dv agencia
    ct = column_names[9]  # conta
    dvct = column_names[10]  # dv conta
    v = column_names[11]  # valor arbitrado
    r = column_names[12]  # valor devido INSS
    o = column_names[23]  # observação
    rr = column_names[24]  # resultado

    print(c, b, a, dva, ct, dvct, v, r, o, rr)

    types = {s: str, p: str,  c: str,  n: str, b: str, a: str,
             dva: str, ct: str, dvct: str, v: str, r: str, o: str, rr: str}

    df2 = df.astype(types)
    print(df2.dtypes)

    # A unidade executora para o Dativo Administrativo é 1080012.
    ue = "1080012"

    """ Infome aqui o login e senha do SIAFI"""

    login = "m1379117"
    senha = "lv2024"

    if login == "xxxx" or senha == "xxxx":
        ev = pi.confirm(text='Registe seu login e senha do SAIFI,\nnas linhas 1364 e 1365 do programa,\nantes de prosseguir',
                        title='', buttons=['OK', 'Cancel'])
        if ev == 'Cancel':
            sys.exit()

    local = "itautec"
    ano = '2024'

    # controle_siafi contains abre_siafi e loga_siafi
    resp = asyncio.run(controle_siafi(login, senha, ue, local, ano))

    print("de volta a ===>  __name__ == '__main__':")

    # print(xxxx)

    ### EMPENHO ###
    # Unidadae executora
    # ue="1080012" # A unidade executora para o Dativo Administrativo é 1080012.
    # unidade orcamentaria = uo
    uo = "1081"
    # ordenador de despesa = od
    # od="10606119" # Dr Fabio PT
    od = "1120530"  # Dra Karem PT

    # Prog trabalho
    ptrab = '78030001'  # old 10140001
    # #Nat. Despesa
    nat_desp = '339091'  # old  339036
    # Item
    item = '16'  # old 21
    # Fonte. Proc/IPE
    fonte = '1090'  # old 1010

    if oper == "Registra Empenho":
        preparo_inicial()
    if oper == "Anula Empenho":
        preparo_inicial_anula()
    if oper != "Registra Empenho" and oper != "Anula Empenho":
        ev = pi.confirm(text=f'OPeracao Indefinida Anula/Registra Empenho \n\n O proghrama será encerrado',
                        title='', buttons=['OK', 'Cancel'])
        sys.exit()

    print("INICIO do loop")
    x = len(df)
    print(x)
    k = 0
    i = 0
    for i in range(x):

        # df.iloc[linha,coluna]
        cpf_cnpj = df2.iloc[i, 4]

        cpf_cnpj = trata_cpf_cnpj(cpf_cnpj)

        # PROVISORIO
        #############################################
        valor = df2.iloc[i, 11]  # valor arbitrado
        # valor=df2.iloc[i,18] # valor liquido , usar somente uma vez para cancelar bloqueios previos
        ##############################################

        valor = trata_valor_4(valor)

        nome = df2.iloc[i, 3]  # nome
        nome = str(nome)
        nome = nome.strip()

        nit = df2.iloc[i, 5]  # nit
        nit = str(nit)
        nit = nit.strip()
        pos1 = nit.find(".")
        if pos1 >= 0:
            nit = nit[:pos1]
            nit = nit.strip()

        sei = df2.iloc[i, 1]  # sei/cpha
        sei = trata_sei_cpha(sei)

        processo = df2.iloc[i, 2]  # processo
        processo = trata_processo(processo)

        print(f'seij_cpha ==> {sei} , {type(sei)}')

        print(f'cpj_cnpj ==> {cpf_cnpj} , {type(cpf_cnpj)}')

        print(f'valor ==> {valor}  , {type(valor)}')

        # sys.exit()

        print(f'linha ==> {i}')
        if i == 0:
            tipo = "inicial"
        else:
            tipo = "outros"
        print(f'tipo = {tipo}')

        obs = df2.iloc[i, 23]  # observacao
        obs = obs.strip()
        print(f' obs  ==> {obs}')

        ct2 = ""
        if oper == "Anula Empenho":
            ######################
            # CANCELA EMPENHO
            ######################

            pos23 = obs.find("-")
            pos24 = obs.find("NAO")
            pos25 = obs.find("NOK")
            pos26 = obs.find("SUCESSO")
            if (pos23 >= 0 and pos24 < 0) or (pos23 >= 0 and pos25 < 0 and pos26 > 0) or obs == "nan" or obs == "X" or obs == "x" or obs == "REPETIR":
                emp_num = obs[pos23+1:]
                emp_num = emp_num.strip()
                tipo = "xx"
                emp = anula_empenho(emp_num, uo, tipo, valor, od)
            else:
                emp = obs
                ct2 = "ja executado"
                # emp="NAO FOI POSSIVEL CANCELAR EMPENHO-FALTA NUMERO"

        if oper == "Registra Empenho":
            ######################
            # EXECUTA EMPENHO
            ######################

            # EMPENHO DA DESPESA NAO REALIZADO-NOK
            # EMPENHO DA DESPESA REALIZADO COM SUCESSO-NOK

            pos23 = obs.find("-")
            pos24 = obs.find("NAO")
            pos25 = obs.find("NOK")
            pos26 = obs.find("SUCESSO")
            if (pos23 >= 0 and pos24 > 0) or (pos23 >= 0 and pos25 > 0) or obs == "nan" or obs == "X" or obs == "x" or ("REPETIR" in obs):
                emp = registra_empenho(
                    cpf_cnpj, valor, uo, od, ptrab, nat_desp, item, fonte, tipo, sei, nome, nit)
                if emp == "cpf_cnpj invalido":
                    emp = "EMPENHO NÃO REALIZADO POR {emp}"
                    print(emp)
            else:
                emp = obs
                ct2 = "ja executado"
                # pi.hotkey('f3')

        # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
        # SE OPERACAO NÃO REALIZADA, NÃO ATUALIZA DF2, NEM SALVA EXCEL

        if ct2 != "ja executado":

            print("Change a cell value in dataframe")
            # df2['OBSERVAÇÕES'][i] = emp

            df2[s][i] = f"'{sei}"  # coluna sei / cpha
            df2[c][i] = f"'{cpf_cnpj}"  # coluna cpf
            df2[p][i] = f"'{processo}"  # processo
            # df2[v][i] = valor # coluna valor
            df2[o][i] = emp  # coluna observacoes

            if "SUCESSO" in emp:
                df2[rr][i] = "OPERACAO REALIZADA COM SUCESSO"
            else:
                df2[rr][i] = emp

            # print(df2)
            # print todas linhas e colunas c,v,o only
            print(df2.loc[:, [c, v, o]])

            """
            k+=1
            if k==10:
                k=0
                print(df2)
            """

            df2.to_excel(file_out, index=False)

    pi.alert(text='---> FIM  EMPENHO <---', title='', button='OK')
    # print(xxxxx)

    # if i==2:
    # break
