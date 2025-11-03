# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 17:52:02 2024

@author: m1379117

release 11

"""
import pyautogui as pi
import pandas as pd
import sys
import os
import time
#import subprocess
import asyncio
#from asgiref.sync import sync_to_async
#import pyperclip
#import re
#import datetime, calendar
import pygetwindow as gw
import platform
import datetime

#############
pi.PAUSE=0.4
pi.FAILSAFE=True
#############

tit='pw3270'
def close_wins(tit):
    print("INICIO da funcao close_wins")
    """ Close all windows with title = tit"""
    #https://github.com/asweigart/PyGetWindow
    
    x=gw.getAllTitles()
    print(x)
    
    qte=x.count(tit)
    print(qte)
    if qte >0:
        for i in range(qte):
            janela=gw.getWindowsWithTitle(tit)[0]
            janela.close()
            print(f'janela {tit}-{i+1} fechada com sucesso')
            time.sleep(1)
    print('ok')
    return

""" Close all windows with title = tit"""    
def close_wins_2(tit):  
    print("INICIO da funcao close_wins_2")
    #https://github.com/asweigart/PyGetWindow
    x=gw.getAllTitles()
    relacao=[]      
    for k, v in enumerate(x):
        # check if text is present in this string from list
        if tit in v:
            # get the name of string which contains the text
            relacao.append(v)
    print(relacao)
    for k,v in enumerate (relacao):
        janela=gw.getWindowsWithTitle(v)[0]        
        janela.close()
        print(f'janela {v}-{k+1} fechada com sucesso')
        time.sleep(1)
    print('ok')
    return

def camada(login,senha,dir1, oper, ue, local, ano, uo, od, ptrab, nat_desp, item, fonte):
    
        
    if time_in_interval('2000','0800'):
        resp='Fora do Horario operacional SIAFI'
        print(resp)
        return ["NOK",'Night']
    
    
    contador=0
    #ano='2024'
    
    
# =============================================================================
#     if oper=="Pagamento - Restos a pagar":
#         ano_1=int(ano)-1
#         ano_rp=ano_1
#         #ano_rp=pi.prompt('Informe o ano do restos a Pagar', title='Restos a pagar', default='2023')
#         #if ano_rp=="Cancel":
#             #sys.exit()
#     else:
#         ano_rp=ano
# =============================================================================
    
    
    
    if oper=="Pagamento - Restos a pagar":
        ano_rp=pi.prompt('Informe o ano do restos a Pagar', title='Restos a pagar', default='2023')
        if ano_rp=="Cancel":
            sys.exit()
    else:
        ano_rp=ano
    
    
    
    
    file_out=dir1
    try:
        df=pd.read_excel(dir1)        
    except:
        print('Falha na abertura do excel')
        av1='arquivo de entrada não localizado'
        msgav='ARQUIVO NAO ENCONTRADO'
        msgsol="Reveja o nome e/ou local do seu arquivo excel."
        msg222='Pode ser falta de algum pacote'
        print(f'{av1}\n{msgav}\n{msgsol}\n{msg222}')
        return ["NOK", av1]

        
    df=df.fillna("x")# Minusculo
    df=df.round(decimals=2)
    print(df)
    column_names = list(df.columns)
    
    file_name = os.path.basename(dir1)
    print(file_name)
    dir_name = os.path.dirname(dir1)
    print(dir_name)
       
    close_wins_2(tit)
    
    if oper=='Empenho' or oper=="Anula Empenho":
       print(oper)
       x=1
       import siafi_empenho_dativo_cpha_rev1 as op
       #import siafi_empenho_dativo_cpha_rev1_desktop as op
       print('importacao classe Empenho OK')
       
       #file=fr'{dir_name}\registra_empenho.xlsx'
       #file_out=file
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
       r=column_names[12] # valor devido INSS
       operac=column_names[20] #Retorno -(EMPENHO)
       o=column_names[23] #observação
       rr=column_names[24] # resultado
       print(c, b, a, dva, ct, dvct, v, r, o, rr, operac)
       types={s:str , p:str,  c:str,  n:str, b:str, a:str, dva:str, ct:str, dvct:str, v:str, r:str, o:str, rr:str, operac:str}
 
       
    if oper=='Liquidacao' or oper=='Anula Liquidacao': 
        print(oper)
        x=1
        import siafi_liquidacao_dativo_cpha_rev2 as op
        print('importacao classe Liquidacao OK')
        
        #file=fr'{dir_name}\registra_liquidacao.xlsx'
        #file_out=file
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
    
        
    if oper=='Pagamento' or oper=='Pagamento - Restos a pagar':
        print(oper)
        x=1
        import siafi_pagamento_dativo_cpha_rev6 as op
        print('importacao classe Pagamento OK')
        
        
        #ev=pi.confirm(text='Aqui , pagamento_rev6 foi importado com sucesso', title='', buttons=['OK', 'Cancel'])
        #if ev=='Cancel':
            #sys.exit()
        #file=fr'{dir_name}\registra_pagamento.xlsx'
        #file_out=file
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
        stot=column_names[16] # subtotal
        ir=column_names[17] # valor irrf
        o=column_names[23] #observação
        rr=column_names[24] # resultado
        print(c, b, a, dva, ct, dvct, v, r, stot, ir, o, rr)
        types={s:str , p:str,  c:str,  n:str, b:str, a:str, dva:str, ct:str, dvct:str, v:str, r:str, stot:str, ir:str, o:str, rr:str}
    
        

    df2=df.astype(types)
    print(df2.dtypes)  
    
    
    #controle_siafi contains abre_siafi e loga_siafi
    import nest_asyncio
    nest_asyncio.apply()
    resp=asyncio.run(op.controle_siafi(login, senha, ue, local, ano))  
    
    
    st22=op.preparo_inicial(oper) # PREPARO INICIAL (POLIMORFISMO DE OOP)
    
    if (oper=='Pagamento' or oper=='Pagamento - Restos a pagar') and st22=="TRANSMISSAO AOS BANCOS EM PROCESSAMENTO":
        print(st22)
        print("Aguardando 10 minutos para nova tentativa ...")
        time.sleep(600)# Aguardando 
        return ['NOK','restart']
                
        #pi.confirm(text='TRANSMISSAO AOS BANCOS EM PROCESSAMENTO\nAguarde alguns minutos antes de continuar\nFalta automatizar o reinicio deste script após 5 minutos.\nO programa será encerrado', title='', buttons=['OK', 'Cancel'])
        #time.sleep(600)
        #sys.exit()
     
    print("INICIO do loop")
    x=len(df)
    print(x)
    
    
    globals()['qte_solicitada_ptpt']=str(x)
    globals()['qte_realizada_ptpt']='0'
    
    #contador1="0"
    #k=0
    i=0
    for i in range(x):
         globals()['qte_realizada_ptpt']=str(contador)
        
         #ops=pi.confirm(text=f'Verifique a Operacao = {oper}\nlinha={i+2}', title='', buttons=['OK', 'Cancel'])
         #if ops=='Cancel':
             #sys.exit()
                    
         if oper=="Empenho":
         ######################    
         #EXECUTA EMPENHO
         ######################
         
             cpf_cnpj=df2.iloc[i,4]
             cpf_cnpj=op.trata_cpf_cnpj(cpf_cnpj)
            
             #PROVISORIO
             #############################################
             valor=df2.iloc[i,11] # valor arbitrado
             if "," in valor:
                  valor=valor.replace(",",".")
             if valor=='x' or int(float(valor))==0:
                 df2[o][i] = "EMPENHO NAO REALIZADO POR valor inexistente"  # coluna observacoes
                 df2[rr][i] = "EMPENHO NAO REALIZADO" #coluna  resultado
             else:    
                 valor=op.trata_valor_4(valor)
            
             nome=df2.iloc[i,3] # nome
             nome=str(nome)
             nome=nome.strip()
                  
             nit=df2.iloc[i,5] # nit
             nit=str(nit)
             nit=nit.strip()
             pos1=nit.find(".")
             if pos1>=0:
                nit=nit[:pos1]
                nit=nit.strip()
            
             sei=df2.iloc[i,1]#sei/cpha
             sei=op.trata_sei_cpha(sei)
           
             processo=df2.iloc[i,2] # processo
             processo=op.trata_processo(processo)
            
             print(f'seij_cpha ==> {sei} , {type(sei)}')     
             print(f'cpj_cnpj ==> {cpf_cnpj} , {type(cpf_cnpj)}')
             print(f'valor ==> {valor}  , {type(valor)}')
             print(f'linha ==> {i}')
             if i==0:
                tipo="inicial"
             else:
                tipo="outros"
             print(f'tipo = {tipo}')
            
           
             obs=df2.iloc[i,23]# observacao
             obs=obs.strip()
             print(f' obs  ==> {obs}')
         
             #EMPENHO DA DESPESA NAO REALIZADO-NOK
             #EMPENHO DA DESPESA REALIZADO COM SUCESSO-NOK
             pos23=obs.find("-")
             pos24=obs.find("NAO")
             pos25=obs.find("NOK")
             pos26=obs.find("SUCESSO")
             pos27=obs.find("CREDOR INEXISTENTE")
             
             # (tem "-" AND tem "NAO") or (tem "-" AND tem "NOK" AND tem SUCESSO) or ...
             if (pos23>=0 and pos24>0) or (pos23>=0 and pos25>0 and pos26>0) or obs=="nan" or obs=="X" or obs=="x" or obs=="REPETIR" or pos27>=0:
                 
                 emp=op.registra_empenho(cpf_cnpj, valor, uo, od, ptrab, nat_desp, item, fonte, tipo, sei, nome, nit)
                 
                 contador=contador+1
                 
                 if emp=="cpf_cnpj invalido" or emp=="CPF invalido":
                    emp=f"EMPENHO NAO REALIZADO POR {emp}"
                    contador=contador-1
                    
                 a1='REPETIR'
                 b1='repetir'
                 c1='EMPENHO DA DESPESA NAO REALIZADO'
                 if (a1 in emp) or (b1 in emp) or (c1 in emp) and i>6:
                    obb1=df2.iloc[i-1,23]
                    obb2=df2.iloc[i-2,23]
                    if (a1 in obb1 or b1 in obb1 or c1 in obb1) and (a1 in obb2 or b1 in obb2 or c1 in obb2):
                        print('RESTART')
                        
                        contador=contador-1
                        
                        pi.hotkey('alt', 'f4')
                        
                        contador1=str(contador)
                       
                        
                        df2[s][i] = f"'{sei}"# coluna sei / cpha
                        df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                        df2[p][i] = f"'{processo}" # processo
                       
                        df2[o][i] = "INDEFINIDO"
                        df2[rr][i] = "INDEFINIDO"# coluna resultado
                        emp11="X"
                        try:
                            df2.to_excel(file_out , index=False)
                        except:
                           ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                           if ev=='Cancel':
                               sys.exit()
                           df2.to_excel(file_out , index=False)
                        
                        globals()['qte_realizada_ptpt']=str(contador)
                        return ['restart',contador1] 
                        #return 'restart'
                    
                 print(emp)
                 print("Changing a cell value in dataframe")
                 #df2['OBSERVAÇÕES'][i] = emp
                
                 df2[s][i] = f"'{sei}"# coluna sei / cpha
                 df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                 df2[p][i] = f"'{processo}" # processo
                 #df2[v][i] = valor # coluna valor
                 df2[o][i] = emp # coluna observacoes
             
                 if "SUCESSO" in emp:
                     df2[rr][i]="OPERACAO REALIZADA COM SUCESSO"# coluna resultado
                     #EMPENHO DA DESPESA REALIZADO COM SUCESSO-245
                     pos233=emp.find("-")
                     emp11=emp[pos233+1:]
                     emp11=emp11.strip()
                 else:
                     df2[rr][i]='EMPENHO NAO REALIZADO'# coluna resultado
                     emp11="X"
                 
                 df2[operac][i]=emp11
             
                  #print(df2)
                  #print todas linhas e colunas c,v,o only
                 print(df2.loc[:,[c,v,o]])
             
                 #df2.to_excel(file_out , index=False)
                 
                  # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                 try:
                    df2.to_excel(file_out , index=False)
                 except:
                    ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                    if ev=='Cancel':
                        sys.exit()
                    df2.to_excel(file_out , index=False)
             
             else:
                 emp=obs
                 ct2="ja executado"
                 
                 
         if oper=="Anula Empenho":
             print('Anulando Empenho...')
             
             emp=df2.iloc[i,20]#Numero do empenho
             
             
             valor=df2.iloc[i,11] # valor arbitrado
             if valor=='x' or int(float(valor))==0:
                 df2[o][i] = "CANCELAMENTO DE EMPENHO NAO REALIZADO POR valor inexistente"  # coluna observacoes
                 df2[rr][i] = "OPERACAO NAO REALIZADA" #coluna  resultado
             else:    
                 valor=op.trata_valor_4(valor)
             
             
             obs=df2.iloc[i,23]# Resultado da OPeracao anterior
             obs=obs.strip()
             print(obs)
             
             if i==0:
                tipo="inicial"
             else:
                tipo="outros"
             print(f'tipo = {tipo}')
             
             ob1="SALDO LIQUIDADO INEXISTENTE"
             ob2="ANULACAO DE LIQUIDACAO REALIZADO COM SUCESSO"
             ob3="EMPENHO DA DESPESA REALIZADO COM SUCESSO"
             
             
             ob5="FALHA NO SINCRONISMO"
             
             
             
             if ob1 in obs or ob2 in obs or ob3 in obs:
                 
                 emp=op.anula_empenho(uo, emp, tipo, valor, od)
                 # ANULACAO DE EMPENHO REALIZADO COM SUCESSO-{emp}"
                 # OU
                 # VALOR ANULACAO EMPENHO MAIOR QUE SALDO-{emp} 
                 
                 contador=contador+1
                 
                 if (ob5 in emp) and i>6:
                        obb1=df2.iloc[i-1,23]
                        obb2=df2.iloc[i-2,23]
                        if (ob1 in obb1) and (ob1 in obb2):
                            print('RESTART')
                            pi.hotkey('alt', 'f4')
                            
                            contador=contador-1
                            
                            contador1=str(contador)
                          
                            
                            df2[s][i] = f"'{sei}"# coluna sei / cpha
                            df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                            df2[p][i] = f"'{processo}" # processo
                           
                            df2[o][i] = "INDEFINIDO"
                            df2[rr][i] = "INDEFINIDO"# coluna resultado
                            emp11="X"
                            try:
                                df2.to_excel(file_out , index=False)
                            except:
                               ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                               if ev=='Cancel':
                                   sys.exit()
                               df2.to_excel(file_out , index=False)
                            
                            globals()['qte_realizada_ptpt']=str(contador)
                            return ['restart',contador1] 
                 
                        
                 print(emp)
                 print("anula empenho - Change a cell value in dataframe")
                                      
                 #df2[s][i] = f"'{sei}"# coluna sei / cpha
                 #df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                 #df2[p][i] = f"'{processo}" # processo
                 df2[o][i] = emp # coluna observacoes
            
                 if "SUCESSO" in emp:
                     df2[rr][i]="OPERACAO REALIZADA COM SUCESSO"# coluna resultado
                     # EMPENHO DA DESPESA REALIZADO COM SUCESSO-245
                 else:
                     df2[rr][i]='EMPENHO NAO REALIZADO'# coluna resultado
                     #print todas linhas e colunas c,v,o only
                 print(df2.loc[:,[c,v,o]])
            
                 # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                 try:
                     df2.to_excel(file_out , index=False)
                 except:
                     ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                     if ev=='Cancel':
                         sys.exit()
                     df2.to_excel(file_out , index=False)
             else:
                 emp=obs
                 ct2="ja executado" 
             
             
             
                 
         if oper=="Liquidacao":
                ######################    
                #EXECUTA LIQUIDACAO
                ######################
                
                valor=df2.iloc[i,11]#valor arbitrado
                valor=op.trata_valor_4(valor)
                
                inss=df2.iloc[i,13]#valor devido inss
                inss=float(inss)
                inss = round(inss,2)
                inss=str(inss)
                inss=inss.strip()
                
                cpf_cnpj=df2.iloc[i,4]
                cpf_cnpj=op.trata_cpf_cnpj(cpf_cnpj)
                                
                sei=df2.iloc[i,1]
                sei=op.trata_sei_cpha(sei)
                
                processo=df2.iloc[i,2] # processo
                processo=op.trata_processo(processo)
                      
                print(f'cpj_cnpj ==> {cpf_cnpj} , {type(cpf_cnpj)}')
                print(f'valor ==> {valor}  , {type(valor)}')
                print(f'linha ==> {i}')
                
               
                if i==0:
                    tipo="inicial"
                else:
                    tipo="outros"
                #if excessao=="SS":
                    #tipo="inicial"
                
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
                # (tem "LIQUIDACAO NAO REALIZADA, EMPENHO") or
                # (tem "EMPENHO DA DESPESA REALIZADO" AND tem "-" AND tem "SUCESSO" AND nao_tem "NAO" AND nao_tem "NOK") or
                # (tem "LIQUIDACAO" AND tem "NAO") or
                # (tem "LIQUIDACAO" AND tem "NOK") or ...
                
                
                if (ob1 in obs) or (pos27>=0 and pos23>=0 and pos28>=0 and pos24<0 and pos25<0) or (pos26>=0 and pos24>=0) or (pos26>=0 and pos25>0) or obs=="nan" or obs=="X" or obs=="x" or obs=="REPETIR":
                    num_emp=obs[pos23+1:]
                    num_emp=num_emp.strip()
                    
                    liq=op.liquida_dativo(num_emp,uo,tipo,valor,od, inss)
                    
                    contador=contador+1
                    
                    if (ob1 in liq) and i>6:
                        obb1=df2.iloc[i-1,23]
                        obb2=df2.iloc[i-2,23]
                        if (ob1 in obb1) and (ob1 in obb2):
                            print('RESTART')
                            pi.hotkey('alt', 'f4')
                            
                            contador=contador-1
                            
                            contador1=str(contador)
                          
                            
                            df2[s][i] = f"'{sei}"# coluna sei / cpha
                            df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                            df2[p][i] = f"'{processo}" # processo
                           
                            df2[o][i] = "INDEFINIDO"
                            df2[rr][i] = "INDEFINIDO"# coluna resultado
                            emp11="X"
                            try:
                                df2.to_excel(file_out , index=False)
                            except:
                               ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                               if ev=='Cancel':
                                   sys.exit()
                               df2.to_excel(file_out , index=False)
                            
                            globals()['qte_realizada_ptpt']=str(contador)
                            return ['restart',contador1] 
                    
                    
                    print("=========> Change a cell value in dataframe")
                    df2[s][i] = f"'{sei}"# coluna sei / cpha
                    df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                    df2[p][i] = f"'{processo}" # processo
                    #df2[v][i] = valor # coluna valor
                    df2[o][i] = liq # coluna observacoes
                    
                    if "SUCESSO" in liq:
                        df2[rr][i]="OPERACAO REALIZADA COM SUCESSO"
                    else:
                        if "LIQUIDACAO NAO REALIZADA" in liq:
                            liq="LIQUIDACAO NAO REALIZADA"
                        df2[rr][i]=liq
                    
                    #print(df2)
                    #print todas linhas e colunas c,v,o only
                    print(df2.loc[:,[c,v,o]])
                    
                    # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                    #df2.to_excel(file_out , index=False)
                    
                     # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                    try:
                        df2.to_excel(file_out , index=False)
                    except:
                        ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                        if ev=='Cancel':
                            sys.exit()
                        df2.to_excel(file_out , index=False)
                        time.sleep(2)
                        a,b=pi.size()
                        pi.click(x=a/2, y=b/2)
                        pi.hotkey('home')
                    
                    
                else:
                    liq=obs
                    ct2="ja executado"
                    # SE OPERACAO NÃO REALIZADA, NÃO ATUALIZA DF2, NEM SALVA EXCEL
            
         if oper=="Anula Liquidacao":
            print("Anulando Liquidacao...")
             
             
            valor=df2.iloc[i,11]#valor arbitrado
            valor1=op.trata_valor_4(valor)
            
            
            #ev=pi.confirm(text=f'Verifique valor antes {valor} e depois {valor1}' , title='inicio anula liquidacao', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            #valor=valor1
            
            #inss=df2.iloc[i,13]#valor devido inss
            #inss=float(inss)
            #inss = round(inss,2)
            #inss=str(inss)
            #inss=inss.strip()
            
            cpf_cnpj=df2.iloc[i,4]
            cpf_cnpj=op.trata_cpf_cnpj(cpf_cnpj)
                            
            sei=df2.iloc[i,1]
            sei=op.trata_sei_cpha(sei)
            
            processo=df2.iloc[i,2] # processo
            processo=op.trata_processo(processo)
                  
            print(f'cpj_cnpj ==> {cpf_cnpj} , {type(cpf_cnpj)}')
            print(f'valor ==> {valor}  , {type(valor)}')
            print(f'linha ==> {i}')
            
           
            if i==0:
                tipo="inicial"
            else:
                tipo="outros"
        
           
            obs=df2.iloc[i,23]
            obs=obs.strip()
            print(obs)
            
            #ev=pi.confirm(text=f'Verifique obs={obs}', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            
            
            pos23=obs.find("-")
            num_emp=obs[pos23+1:]
            num_emp=num_emp.strip()
            print(num_emp)
            
            #ev=pi.confirm(text=f'Verifique obs={obs}  e empenho ={num_emp}', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            
            
            #ev=pi.confirm(text=f'Verifique numero do empenho da planilha =  {num_emp}', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
                       
            
            ob3="ANULACAO DE LIQUIDACAO REALIZADO COM SUCESSO"
            if ob3 in obs:
                ct2="ja executado"
                continue
            
            ob1="ANULACAO DE LIQUIDACAO NAO REALIZADA, EMPENHO"
            ob2="LIQUIDACAO REALIZADA COM SUCESSO, EMPENHO"
            ob4="SALDO LIQUIDADO INEXISTENTE"
            
            
            ####################################################
            
    
            l2=[obs, ob1, ob2, ob4]
            l3=[ob1 in obs ,  ob2 in obs, i>6 , ob4 not in obs]
            for k in l2:
                print(k)
            print(l3)
            ##################################################
            
            #ev=pi.confirm(text=f'Verifique parametros do sistema = {l3}', title='', buttons=['OK', 'Cancel'])
            #if ev=='Cancel':
                #sys.exit()
            
            
            
            
            if (ob1 in obs or ob2 in obs) and (ob4 not in obs):
            
                liq=op.anula_liquidacao(num_emp, uo, valor, od, tipo)
                
                contador=contador+1
                
                
                #ev=pi.confirm(text=f'Verifique resultado da Anulacao = \n{liq}', title='', buttons=['OK', 'Cancel'])
                #if ev=='Cancel':
                    #sys.exit()
                
                #ANULACAO DE LIQUIDACAO NAO REALIZADA, EMPENHO-41376, SALDO LIQUIDADO INEXISTENTE
                
                
                #if ob1 in liq and i>6 and ob4 not in obs:
                if ob1 in obs and i>6 and ob4 not in obs:
                    obb1=df2.iloc[i-1,23]
                    obb2=df2.iloc[i-2,23]
                    if (ob1 in obb1) and (ob1 in obb2):
                        print('RESTART')
                        
                        ####################################################
                        #l2=[obs, ob1, ob2, ob4, liq ]
                        #l3=[ob1 in obs ,  i>6 , ob4 not in obs , ob1 in liq]
                        #for k in l2:
                            #print(k)
                        #print(l3)
                        ##################################################
                        
                        #ev=pi.confirm(text=f'Verifique parametros de restart = {l3}', title='', buttons=['OK', 'Cancel'])
                        #if ev=='Cancel':
                            #sys.exit()
                       
                        pi.hotkey('alt', 'f4')
                        
                        contador=contador-1
                        
                        contador1=str(contador)
                        
                        df2[s][i] = f"'{sei}"# coluna sei / cpha
                        df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                        df2[p][i] = f"'{processo}" # processo
                       
                        df2[o][i] = "INDEFINIDO"
                        df2[rr][i] = "INDEFINIDO"# coluna resultado
                        emp11="X"
                        try:
                            df2.to_excel(file_out , index=False)
                        except:
                           ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                           if ev=='Cancel':
                               sys.exit()
                           df2.to_excel(file_out , index=False)
                        
                        globals()['qte_realizada_ptpt']=str(contador)
                        return ['restart',contador1] 
                
                
                print("=========> Change a cell value in dataframe")
                df2[s][i] = f"'{sei}"# coluna sei / cpha
                df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                df2[p][i] = f"'{processo}" # processo
                #df2[v][i] = valor # coluna valor
                df2[o][i] = liq # COLUNA observacoes
                
                
                # COLUNA Resultado
                if "SUCESSO" in liq:
                    df2[rr][i]="OPERACAO REALIZADA COM SUCESSO"
                else:
                    if "ANULACAO DE LIQUIDACAO NAO REALIZADA" in liq:
                        liq="OPERACAO NAO REALIZADA COM SUCESSO"
                    df2[rr][i]=liq
                
                #print(df2)
                #print todas linhas e colunas c,v,o only
                print(df2.loc[:,[c,v,o]])
                
                # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                #df2.to_excel(file_out , index=False)
                
                 # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                try:
                    df2.to_excel(file_out , index=False)
                except:
                    ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e reinicie./O programa será encerrado' , title='', buttons=['OK', 'Cancel'])
                    sys.exit()
                    #df2.to_excel(file_out , index=False)
                    #time.sleep(2)
                    #a,b=pi.size()
                    #pi.click(x=a/2, y=b/2)
                    #pi.hotkey('esc')
                
                
            else:
                liq=obs
                ct2="ja executado"
                print(ct2)
                # SE OPERACAO NÃO REALIZADA, NÃO ATUALIZA DF2, NEM SALVA EXCEL
            
                         
            
         if oper=="Pagamento" or oper=="Pagamento - Restos a pagar":
                ######################    
                #EXECUTA PAGAMENTO
                ######################
                
                print(f'====================================Inicio da linha da planilha ==> {i+2} ================================================')
               
                # df.iloc[linha,coluna]
                cpf_cnpj=df2.iloc[i,4]
                cpf_cnpj=op.trata_cpf_cnpj(cpf_cnpj)
                 
               
                valor=df2.iloc[i,16]# valor devido  (bruto - retencao inss)
                valor=op.trata_valor_4(valor)
                
                valor_bruto=df2.iloc[i,11]# valor arbitrado (bruto - retencao inss)
                valor_bruto=op.trata_valor_4(valor_bruto)
                
                inss=df2.iloc[i,13]#valor devido inss
                if inss==None or inss=="x" or inss=="X":
                    inss=0
                
                inss=float(inss)
                inss = round(inss,2)
                inss=str(inss)
                inss=inss.strip()
                inss=op.trata_valor_4(inss)
                               
                
                sei=df2.iloc[i,1]# sei
                sei=op.trata_sei_cpha(sei)
                
                
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
                if irr==None or irr=="x" or irr=="X":
                    irr=0
                #irr=op.trata_irr(irr)    
                irr=op.trata_valor_4(irr)
               
                obs=df2.iloc[i,23]# observacao
                obs=obs.strip()
               
                #pg=paga_dativo(emp, uo, tipo1, cpf_cnpj, banco, agencia, dv_a,  conta,  dv_c, valor, od, inss, sei):
                processo=df2.iloc[i,2] # processo
                processo=op.trata_processo(processo)
        
                print(f'cpj_cnpj ==> {cpf_cnpj} , {type(cpf_cnpj)}')
                print(f'valor ==> {valor}  , {type(valor)}')
                print(f'linha ==> {i}')
                
                if i==0:
                    tipo="inicial"
                else:
                    tipo="outros"
                print(tipo)
                print(f'obs ==> {obs}')
                
                pos23=obs.find("-")
                pos24=obs.find("NAO")
                pos25=obs.find("NOK")
                pos26=obs.find("PAGAMENTO")
                pos27=obs.find("LIQUIDACAO REALIZADA COM SUCESSO")
                pos28=obs.find("SUCESSO")
                #pos29=obs.find("INEXISTENTE")
                pos29=obs.find("AGENCIA")
                pos32=obs.find("FALTA")
                
                if pos29>0 or pos32>0:
                    pos29=1
                pos30=obs.find("BANCO")
                pos31=obs.find("SALDO NAO DISPONIVEL")
                
                ob1="PAGAMENTO NAO REALIZADO, EMPENHO"
                #PAGAMENTO NAO REALIZADO, EMPENHO-15852, BANCO, AGENCIA OU CONTA INEXISTENTE 341 / 876 / 4 / 000000000507 / 3
                
                aa=bb=cc=dd=ee=ff=gg=hh=ii=jj="False"
                
                # Tem "PAGAMENTO NAO REALIZADO, EMPENHO"
                # Não tem "INEXISTENTE"
                if (ob1 in obs and pos29<0):
                    if pos31>0:
                        print('Resto a pagar nao disponivel')
                    if 'NIT/PIS/PASEP' in obs:
                        print('NIT/PIS/PASEP LOCALIZADO')
                    if (pos31<=0 and 'NIT/PIS/PASEP' not in obs):
                        print("Executado pelo criterio A")
                        aa="True"
                    
                # Não tem "PAGAMENTO", Não tem "SUCESSO", Não tem "INEXISTENTE"
                # Tem "LIQUIDACAO REALIZADA COM SUCESSO"
                if (pos26<0 and pos28<0 and pos29<0 and pos27>=0):
                    if ('NIT/PIS/PASEP' not in obs) and ('CPF' not in obs):
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
                    ob1='Sincronismo'
                    
                    pg=op.paga_dativo(num_emp, uo, tipo, cpf_cnpj, banco, agencia, dv_a,  conta,  dv_c, valor, valor_bruto, od, inss, irr, sei, processo, i, ano_rp)
                    
                    contador=contador+1
                     
                    if (ob1 in pg) and i>6:
                        obb1=df2.iloc[i-1,23]
                        obb2=df2.iloc[i-2,23]
                        if (ob1 in obb1) and (ob1 in obb2):
                            print('RESTART')
                            pi.hotkey('alt', 'f4')
                            
                            contador=contador-1
                            contador1=str(contador)
                            df2[s][i] = f"'{sei}"# coluna sei / cpha
                            df2[c][i] = f"'{cpf_cnpj}" # coluna cpf
                            df2[p][i] = f"'{processo}" # processo
                           
                            df2[o][i] = "INDEFINIDO"
                            df2[rr][i] = "INDEFINIDO"# coluna resultado
                            emp11="X"
                            try:
                                df2.to_excel(file_out , index=False)
                            except:
                               ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                               if ev=='Cancel':
                                   sys.exit()
                               df2.to_excel(file_out , index=False)
                            
                            globals()['qte_realizada_ptpt']=str(contador)
                            return ['restart',contador1] 
                            
                    
                    print("Change a cell value in dataframe")
                    
                    if dv_a=='x':
                        df2[dva][i] = None
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
                            
                     # SE OPERACAO REALIZADA ATUALIZA dF2 E SALVA DF2 NO EXCEL
                    try:
                        df2.to_excel(file_out , index=False)
                    except:
                        ev=pi.confirm(text='Se a planilha estiver aberta, feche a mesma e clique OK, Para continuar.' , title='', buttons=['OK', 'Cancel'])
                        if ev=='Cancel':
                            sys.exit()
                        df2.to_excel(file_out , index=False)
                        
                        
    
                else:
                    print("Nao Executado por nao atender nenhum criterio")
                    pg=obs
                    ct2="ja executado"
                    #SE OPERACAO NÃO REALIZADA, NÃO ATUALIZA DF2, NEM SALVA EXCEL
         
                    
         
         if i+1==len(df):
             print("ajustando colunas automaticamente ...")
             x1 = pd.ExcelFile(file_out)
             pla=list(x1.sheet_names)  # see all sheet names
             print(type(pla))
             sheet_name=(pla[0])
             print(sheet_name)
             
             #fit_cols(file_out,sheet_name)
    
    globals()['qte_realizada_ptpt']=str(contador)
    contador1=str(contador)
    resp=["OK",contador1]


    #ev=pi.confirm(text=f'Verifique o resultado do processo de camada:\n{resp}' , title='camada', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
      #sys.exit()          
    
    return resp    


def fit_cols(file,sheet_name):
    """ sets excel sheet, Column width automatically """
    print("INICIO da funcao fit_col")
    import win32com.client as win32
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    time.sleep(1)
    wb = excel.Workbooks.Open(file)
    ws = wb.Worksheets(sheet_name)
    ws.Columns.AutoFit()
    #wb.Save()
    time.sleep(1)
    wb.SaveAs(file)
    excel.Application.Quit()
    
    return "OK"

def time_in_interval(timeStart,timeEnd):
    """ Boolean -> Check if current time is inside an interval"""
    #timeStart = '0300' Hora Inicial
    #timeEnd = '0600' Hora Final
    
    from datetime import datetime

    def isNowInTimePeriod(startTime, endTime, nowTime):
        if startTime < endTime:
            return nowTime >= startTime and nowTime <= endTime
        else: #Over midnight
            return nowTime >= startTime or nowTime <= endTime
      
    
    timeEnd = datetime.strptime(timeEnd, '%H%M').time()
    timeStart = datetime.strptime(timeStart, '%H%M').time()
    timeNow = datetime.now().time()
    
    resp=isNowInTimePeriod(timeStart, timeEnd, timeNow)
    print(resp)
    return resp
    
    

def today_now():
    print("INICIO da funcao today_now")
    #https://www.geeksforgeeks.org/get-current-date-and-time-using-python/
    
    from datetime import datetime
   
    
    ct=datetime.now()
   
    print(ct)
    print(type(ct))
    
    dd=str(ct.day)
    mo=str(ct.month)
    aa=str(ct.year)
    hh=str(ct.hour)
    mm=str(ct.minute)
    ss=str(ct.second)

    resp=dd+"/"+mo+"/"+aa+" "+hh+":"+mm+":"+ss
    print(resp)
    
    return resp


def reg_api_sedig(usuario, depto, processo, inicio, fim, qte_solicitada, qte_realizada, sucesso, erro):
    print("INICIO da funcao reg_api_sedig")
    
    
    import requests
       
# =============================================================================
#     if qte_realizada==qte_solicitada and qte_solicitada != "0":
#         print("sucesso")
#         sucesso="SIM"
#         obs="X"
#     else:
#         print("falha")
#         sucesso="NAO"
#         obs=erro
# =============================================================================
        
    obs=erro
    
    if qte_realizada=='Night' or qte_solicitada=='Night':
       sucesso="NAO"
       obs=erro

    depto=str(depto)
    processo=str(processo)
    usuario=str(usuario)
    inicio=str(inicio)
    fim=str(fim)
    qte_solicitada=str(qte_solicitada)
    qte_realizada=str(qte_realizada)
    sucesso=str(sucesso)
    obs=str(obs)
    
    print(type(depto))
    print(type(processo))
    print(type(usuario))
    print(type(inicio))
    print(type(fim))
    print(type(qte_solicitada))
    print(type(qte_realizada))
    print(type(sucesso))
    print(type(obs))
    
    
    print(depto)
    print(processo)
    print(usuario)
    print(inicio)
    print(fim)
    print(qte_solicitada)
    print(qte_realizada)
    print(sucesso)
    print(obs)
    
    
    depto=depto.strip()
    processo=processo.strip()
    usuario=usuario.strip()
    inicio=inicio.strip()
    fim=fim.strip()
    qte_solicitada=qte_solicitada.strip()
    qte_realizada=qte_realizada.strip()
    sucesso=sucesso.strip()
    obs=obs.strip()
    
    if len(obs)>400:
        obs=obs[:390]
    
    
    
    data= {
          "nome_usuario": usuario,
          "setor_atendido": depto,
          "processo_requisitado": processo,
          "dia_hora_inicio": inicio,
          "dia_hora_fim": fim,
          "qte_realizada": qte_realizada,
          "qte_solicitada": qte_solicitada,
          "concluido_com_sucesso": sucesso,
          "obs": obs
        }
    
       
    
    """POST"""
    try:
        print("registrando na VM-Prodemge")
        x=requests.post('http://200.198.59.230:8002/robot/', data=data)
    except:
        try:
            print("registrando na VM-AGE")
            x=requests.post('http://192.168.113.60:8002/robot/', data=data)
        except:
            print(f"Erro ao registrar operacao do SPI {processo}\n{x}")
            pass
    
    return



    
def siafi():    
    x=0
    while x<1:
        login=''
        login=pi.prompt(text='Informe seu login do Siafi', title='LOGIN SIAFI-release 14',default=login)
        if login==None:
            sys.exit()
        if login!='login':
            x=1
    x=0
    while x<1:
        senha=''        
        senha=pi.prompt(text='Informe sua senha do Siafi', title='SENHA SIAFI-release 14',default=senha)
        if senha==None:
            sys.exit()
        if senha!='senha':
            x=1
    
    home_dir=os.path.expanduser('~')
    print(home_dir)
    file=fr'{home_dir}\desktop\registra.xlsx'
    #file=r'C:\Users\m1379117\OneDrive\AGE\AGE TELETRABALHO\03_2024\PT pagamento_SIAFI_8471_novo\4001 a 6000\registra_empenho.xlsx'
    
    msgav='Arquivo de Entrada'
    msgsol='Informe diretorio e nome da planilha a ser processada'
    
    globals()['usuario']=login
    x=0
    while x<1:
        dir1=pi.prompt(text=msgsol, title=msgav ,default=file)
        print(dir1)
        if dir1==None:
            sys.exit()
        else:
            x=1
    
    
    
    
    # TRATA DATAS
    xd = datetime.datetime.now()
    print(xd)
    dia=xd.strftime("%x")
    print(f'today ==> {dia}')
    #yy="2022"
    #mm=xd.month
    #dd=x.day
    yy=xd.year
    ano=str(yy)
    print(ano)
    
    #dd1=str(dd)
    
    
    # =============================================================================
    # PARAMETROS #######
    ue="1080012" # A unidade executora para o Dativo Administrativo é 1080012.
    local="itautec"
    #ano='2024'
    uo="1081" #unidade orcamentaria = uo
    od="1120530" # Dra Karem PT  #ordenador de despesa --> Era od="10606119" # Dr Fabio PT
    
    #ptrab='10140001'
    ptrab='78030001'  # NOVO Prog trabalho --> Era ptrab='10140001'
    
    #nat_desp='339036'
    nat_desp='339091' # NOVO Nat. Despesa -->  Era nat_desp='339036'
    
    item='16'  # NOVO Item era 21
    fonte='1090'  # Fonte. Proc/IPE --> 1090
    # =============================================================================
    
        
    x=0
    while x<1:
        oper=pi.prompt(f'Escolha um NÚMERO:\n   --------------------------\n  1 - Empenho\n  2 - Liquidacao\n  3 - Pagamento\n  4 - Pagamento - Restos a pagar\n   ---------------------------\n  5 - Anula Empenho\n  6 - Anula Liquidacao\n\n\nPlanilha ser processada:\n{dir1}', 'OPERACAO SIAFI  - release 14', '')
        if oper==None:
            sys.exit()
            
        escolha = {"1":"Empenho",
                   "2":"Liquidacao",
                   "3":"Pagamento",
                   "4":"Pagamento - Restos a pagar",
                   "5":"Anula Empenho",
                   "6":"Anula Liquidacao"
                   }
        
        for k,v in enumerate(escolha.items()):
            
            print(f' tipo de oper = {type(oper)}')
            print(f' tipo de k =>{type(k)}')
            print(f' tipo de v ==> {type(v[1])}')
            print(k , v)
            try:
                if k==int(oper)-1:
                    oper=v[1]
                    #if oper!="Anula Empenho":
                    x=1
                    break
            except:
                pass
            
        print(f'===> Operacao = {oper}')
       
   
   
   
    depto="PT-FIN"
    processo="PGTO_SIAFI"
    
    hn1=platform.node()
    usuario=f'SPI-{oper} : {login} - {hn1}'
    usuario=usuario[0:48]
    globals()['usuario']=usuario
    
    
    inicio=today_now()
    sucesso="SIM"
    
    
    ctq=10
    for i in range(ctq):
        
        if i==9:
            inicio=today_now()
            resp=camada(login, senha, dir1, oper, ue, local, ano, uo, od, ptrab, nat_desp, item, fonte)
            
            fim=today_now()
            qte_solicitada=globals()['qte_solicitada_ptpt']
            qte_realizada=globals()['qte_realizada_ptpt']
            
            ########## EM TESTE ################
            qte_solicitada=str(int(qte_solicitada)-int(qte_realizada))
            if int(qte_realizada)>=int(qte_solicitada):
                qte_realizada=qte_solicitada
            ###################################
            
            
            sucesso="NAO"
            erro="Programa suspenso após 9 tentativas"
            
            try:
                reg_api_sedig(usuario, depto, processo, inicio, fim, qte_solicitada, qte_realizada, sucesso, erro)
            except:
                pass
            
            pi.alert(text=f'Operacao de {oper} no SIAFI\nNão foi totalmente concluida com sucesso\nVerifique qte executada na planilha.', title='OPERACAO SIAFI', button='OK')
            sys.exit()
        
        if i==0:
            inicio=today_now()
            resp=camada(login, senha, dir1, oper, ue, local, ano, uo, od, ptrab, nat_desp, item, fonte)
            
        
        if resp[0]=='restart':
            
            inicio=today_now()
            
            resp=camada(login, senha, dir1, oper, ue, local, ano, uo, od, ptrab, nat_desp, item, fonte)
            
            fim=today_now()
            qte_solicitada=globals()['qte_solicitada_ptpt']
            qte_realizada=globals()['qte_realizada_ptpt']
                    
            
            ########## EM TESTE ################
            qte_solicitada=str(int(qte_solicitada)-int(qte_realizada))
            if int(qte_realizada)>=int(qte_solicitada):
                qte_realizada=qte_solicitada
            ###################################
            
            sucesso="SIM"
            erro="restarted"
            
            try:
                reg_api_sedig(usuario, depto, processo, inicio, fim, qte_solicitada, qte_realizada, sucesso, erro)
            except:
                pass
            
            
            
        if resp[0]=='OK':
            
            print('FIM')
            pi.hotkey('alt', 'f4')
            
            fim=today_now()
            qte_solicitada=globals()['qte_solicitada_ptpt']
            qte_realizada=globals()['qte_realizada_ptpt']
            sucesso="SIM"
            erro="X"
            #print(usuario, depto, processo, inicio, fim, qte_solicitada, qte_realizada, sucesso, erro)
            
            try:
                reg_api_sedig(usuario, depto, processo, inicio, fim, qte_solicitada, qte_realizada, sucesso, erro)
            except:
                pass
            
            break
        
    print(usuario, depto, processo, inicio, fim, qte_solicitada, qte_realizada, sucesso, erro)
    
    #ev=pi.confirm(text=f'Verifique o resultado do processo de SIAFI:\n{resp}' , title='siafi()', buttons=['OK', 'Cancel'])
    #if ev=='Cancel':
      #sys.exit()     
    
    
    pi.alert(text=f'Operacao de {oper} no SIAFI\nConcluida com Sucesso.', title='OPERACAO SIAFI', button='OK')
    #return "FIM"
    
            

if __name__ == '__main__':
     globals()['qte_solicitada_ptpt']='0'
     globals()['qte_realizada_ptpt']='0'
     globals()['inicio']=today_now()
     globals()['usuario']="xx"
     
     try:
         siafi()
     except Exception as e:
       
        print('Erro na tentativa tentativa 1, indo para EXception')
        
        exc_type, exc_value, tb = sys.exc_info()
        local_vars = {}
        while tb:
            filename = tb.tb_frame.f_code.co_filename
            name = tb.tb_frame.f_code.co_name
            line_no = tb.tb_lineno
            print(f"ERRO:\n{str(e)}\nFile {filename}\nline {line_no}, in {name}")
    
            local_vars = tb.tb_frame.f_locals
            tb = tb.tb_next
        #print(f"Local variables in top frame: {local_vars}")
        
        base=os.path.basename(filename)
        erro=f'Erro no codigo siafi() : {str(e)}, na linha: {line_no} , da funcao: def {name}: , do arquivo: {base}'
        print(f'erro = {erro}')
        
             
        print("===========================")
        
        #fim=today_now()
        #qte_solicitada=globals()['qte_solicitada_ptpt']
        #qte_realizada=globals()['qte_realizada_ptpt']
        #sucesso="S"
        #erro="restarted"
        
        
        
        fim=today_now()
        sucesso="NAO"
        solicitado=globals()['qte_solicitada_ptpt']
        #solicitado=str(int(solicitado)+1)
        realizado=globals()['qte_realizada_ptpt']
        inicio=globals()['inicio']
        usuario=globals()['usuario']
       
        
        hn1=platform.node()
        depto="PT-FIN"
        processo="PGTO_SIAFI"    
        
        #usuario=f'SPI-{usuario} : da {depto} no {hn1}'
        #usuario=usuario[0:47]
        usuario=globals()['usuario']
        
        reg_api_sedig(usuario, depto, processo, inicio, fim, solicitado, realizado, sucesso, erro)
        
        pi.alert(text='Ocorreu um erro durante o processamento da sua solicitacao.', title='inicio-main', button='OK')
        