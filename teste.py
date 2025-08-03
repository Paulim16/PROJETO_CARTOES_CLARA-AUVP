import csv
import datetime

from datetime import datetime, timedelta, date
import pandas as pd
from dateutil.relativedelta import relativedelta
data_de_hoje = date.today()
mes_atual = data_de_hoje.month
mes_seguinte = int(mes_atual) + 1
import openpyxl





#Variáveis na função que mudam por subsidiária
contaDNS = '2.1.2.01.042 CLARA CARTÃO - DNS'
contaTB = '2.1.2.01.043 CLARA CARTÃO - THE BRAIN'
contaSPN = '2.1.2.01.045 CLARA CARTÃO - SUPERNOVA'
contaAUVP = '2.1.2.01.044 CLARA CARTÃO - AUVP'
idDNS = '1828'
idTB ='1829'
idSPN = '1832'
idAUVP = '1831'
locDNS = '2'
locTB = '3'
locSPN = '1'
locAUVP = '6'
vencDNS = f"15/{mes_atual}/2025"
vencTB = f"15/{mes_atual}/2025"
vencSPN = f"12/{mes_atual}/2025"
vencAUVP = f"22/{mes_atual}/2025"
empDNS = '7'
empTB = '10'
emSPN = '6'
empAUVP = '16'

dados = {
    'EMPRESA': [],
    'CONTA': [],
    'MEMORANDO': [],
    'ENTIDADE': [],
    'DEPARTAMENTO': [],
    'CONTA DESPESAS': [],
    'CLASSE': [],
    'DATA': [],
    'PERÍODO CONTÁBIL': [],  #VERIFICAR IDEIA DE SENDO DO DIA 16 AO DIA 31, MÊS ATUAL, E DO DIA 1 AO 16, MÊS SEGUINTE
    'VENCIMENTO': [],
    'VALOR TXT': [],
    'VALOR': [],
    'Nº REF': [], #f'="Cartão" & "_" & "2025." & H{linha} & ".00" & LIN(H{linha})'
    'ID CONTA': [],
    'ID CONTA DESPESA': [], #f"=PROCX(F{linha};Planodecontas!$E:$E;Planodecontas!$A:$A;)"
    'ID FORNECEDOR': [], #f"=PROCX(D{linha};Fornecedores!$B:$B;Fornecedores!$A:$A;)"
    'ID EMPRESA': [],
    'ID CLASSE': [], #f"=PROCX(G{linha};classes!$B:$B;classes!$A:$A;)"
    'ID DEPARTAMENTO': [], #f"=PROCX(E{linha};Departamentos!$B:$B;Departamentos!$A:$A;)"
    'ID LOCALIDADE': []
    }

def importa(empresa, conta, entidade, departamento, contadespesa, classe, vencimento, idconta, idempresa, idloc):
    dados['EMPRESA'].append(empresa)
    dados['CONTA'].append(conta)
    dados['MEMORANDO'].append(f'CARTÃO {linha[2]} - {linha[4]}')
    dados['ENTIDADE'].append(entidade)
    dados['DEPARTAMENTO'].append(departamento)
    dados['CONTA DESPESAS'].append(contadespesa)
    dados['CLASSE'].append(classe)
    dados['DATA'].append(linha[0])        
    dados['PERÍODO CONTÁBIL'].append(f"01/{str(linha[1][3:5])}/2025")
    dados['VENCIMENTO'].append(vencimento)
    dados['VALOR TXT'].append(nvalor) 
    dados['VALOR'].append(nvalor)
    dados['Nº REF'].append(f'="Cartão" & "_" & "2025." & H{i + 1} & ".00" & LIN(H{i + 1})') #f'="Cartão" & "_" & "2025." & H{linha} & ".00" & LIN(H{linha})'
    dados['ID CONTA'].append(idconta)
    dados['ID CONTA DESPESA'].append(f"=PROCX(F{i};Planodecontas!$E:$E;Planodecontas!$A:$A;)")  #f"=PROCX(F{linha};Planodecontas!$E:$E;Planodecontas!$A:$A;)"
    dados['ID FORNECEDOR'].append(f"=PROCX(D{i};Fornecedores!$B:$B;Fornecedores!$A:$A;)")  #f"=PROCX(D{linha};Fornecedores!$B:$B;Fornecedores!$A:$A;)"
    dados['ID EMPRESA'].append(idempresa)
    dados['ID CLASSE'].append(f"=PROCX(G{i};classes!$B:$B;classes!$A:$A;)")  #f"=PROCX(G{linha};classes!$B:$B;classes!$A:$A;)"
    dados['ID DEPARTAMENTO'].append(f"=PROCX(E{i};Departamentos!$B:$B;Departamentos!$A:$A;)")  #f"=PROCX(E{linha};Departamentos!$B:$B;Departamentos!$A:$A;)"
    dados['ID LOCALIDADE'].append(idloc)

with open('TESTE-SELENIUM-DNS-JAN_FEV (2).csv','r') as cartao:
    leitor = csv.reader(cartao, delimiter=",")
   

    for i, linha in enumerate(leitor):
        
        
        if i >= 1:

            nvalor = linha[7].replace('.', ',')

            data = datetime.strptime(linha[1], "%d-%m-%Y")
            mes_seguinte = data + relativedelta(months=1)

            
            if int(linha[1][0:2]) >=16:
                month = mes_seguinte.month
                year = mes_seguinte.year

            else:
                month = data.month
                year = data.year
                
            vencDNS = datetime.strftime(datetime(day=15, month = month, year=year),"%d/%m/%Y")
            vencTB = datetime.strftime(datetime(day=15, month = month, year=year),"%d/%m/%Y")



            # cartao('DNS', contaDNS, 'ent', 'dep', 'cont', 'class', vencDNS, idDNS, empDNS,locDNS)

            ### ------------------------- COMPRAS DNS -------------------------------- ###
            # FACEBOOK
            if 'FACEBK' in linha[4] and '2289' in linha[2]: 
                importa('DNS', contaDNS, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'PUBLICIDADE', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
              
            elif 'FACEBK' in linha[4] and '4836' in linha[2]:
                importa('DNS', contaDNS, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'PUBLICIDADE', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'AUVP ANALÍTICA', vencDNS, idDNS, empDNS,locDNS)  
            # Imagify
            elif 'WP MEDIA - IMAGIFY' in linha[4] and '8389' in linha[2] and int(linha[1][0:2]) == 9:
                importa('DNS', contaDNS, 'IMAGIFY', 'PUBLICIDADE', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # Openai dia 09 e 23 DNS - alyf
            elif 'OPENAI *CHATGPT SUBSCR' in linha[4] and '3559' in linha[2] and (int(linha[1][0:2]) == 9 or int(linha[1][0:2]) == 23) :
                importa('DNS', contaDNS, 'OPENAI,LLC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            #Amazon Prime
            elif 'AmazonPrimeBR' in linha[4] and '8389' in linha[2] and int(linha[1][0:2]) == 10 and nvalor == '19,90':
                importa('DNS', contaDNS, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'CONTROLADORIA E FINANÇAS', '3.5.1.05.023 OUTRAS DESPESAS ADMINISTRATIVAS', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # Fundamentei
            elif 'FUNDAMENTEI.COM' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 10 and nvalor == '49,00':
                importa('DNS', contaDNS, 'FUNDAMENTEI SERVICOS DE INFORMACAO LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # Brapi Dev 
            elif 'BRAPI.DEV' in linha[4] and '2289' in linha[2] and int(linha[1][0:2]) == 10 and nvalor == '49,99':
                importa('DNS', contaDNS, 'BRAPI ASL TECNOLOGIA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # STAPE INC
            elif 'STAPE, INC' in linha[4] and '2289' in linha[2] and 9 <= int(linha[1][0:2]) == 11:
                importa('DNS', contaDNS, 'STAPE, INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # EBN Canva
            elif 'EBN *Canva' in linha[4] and '8409' in linha[2] and int(linha[1][0:2]) == 13 and nvalor == '174,50':
                importa('DNS', contaDNS, 'CANVA PTY LTD.', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # LEARNWORLDS CY LTD dia 13, 22 e 29
            elif 'LEARNWORLDS CY LTD' in linha[4] and '2289' in linha[2] and (int(linha[1][0:2]) == 13 or int(linha[1][0:2]) == 22 or int(linha[1][0:2]) == 29):
                importa('DNS', contaDNS, 'LEARNWORLDS (CY) LTD', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # WINDSOR.AI
            elif 'WINDSOR.AI' in linha[4] and '2289' in linha[2] and int(linha[1][0:2]) == 14:
                importa('DNS', contaDNS, 'WINDSOR.AI', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # AOVS
            elif 'AOVS SISTEMAS DE INFOR' in linha[4] and '2289' in linha[2] and int(linha[1][0:2]) == 16 and nvalor == '530,00':
                importa('DNS', contaDNS, 'AOVS SISTEMAS DE INFORMATICA SA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # ELEVENLABS
            elif 'ELEVENLABS.IO' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 17:
                importa('DNS', contaDNS, 'ELEVENLABS.IO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # MICROSOFT
            elif 'MSFT *' in linha[4] and '2289' in linha[2] and int(linha[1][0:2]) == 18:
                importa('DNS', contaDNS, 'MICROSOFT INFORMATICA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # KIWIFY NOTIFICA
            elif 'NOTIFICACOES INTELIGEN' in linha[4] and '2236' in linha[2] and int(linha[1][0:2]) == 18:
                importa('DNS', contaDNS, 'KIWIFY PAGAMENTOS, TECNOLOGIA E SERVICOS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS) 
            # UAZAPI
            elif 'UAZAPI - API WHATSAPP' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 19 and nvalor == '29,00' :
                importa('DNS', contaDNS, 'UAZAPI', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # STAPE
            elif 'STAPE, INC.' in linha[4] and '2289' in linha[2] and int(linha[1][0:2]) == 19:
                importa('DNS', contaDNS, 'STAPE, INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # TINY
            elif 'TINY ERP' in linha[4] and '8389' in linha[2] and int(linha[1][0:2]) == 20:
                importa('DNS', contaDNS, 'OLIST TINY TECNOLOGIA LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # ADOBE DIA 20 e 06   
            elif 'ADOBE' in linha[4] and '2289' in linha[2] and (int(linha[1][0:2]) == 20 or int(linha[1][0:2]) == 6) and nvalor == '275,00':
                importa('DNS', contaDNS, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # SEMRUSH   
            elif 'EBN *SEMRUSH' in linha[4] and '8389' in linha[2] and int(linha[1][0:2]) == 21:
                importa('DNS', contaDNS, 'SEMRUSH', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # FIGMA
            elif 'FIGMA' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 21:
                importa('DNS', contaDNS, 'FIGMA MONTHLY RENEWAL', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # BITLy  
            elif 'BITLY.COM' in linha[4] and '2289' in linha[2] and int(linha[1][0:2]) == 22:
                importa('DNS', contaDNS, 'BITLY COM', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # MANYCHAT    
            elif 'MANYCHAT.COM' in linha[4] and '5518' in linha[2] and int(linha[1][0:2]) == 25:
                importa('DNS', contaDNS, 'MANYCHAT INC.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # SUPABASE    
            elif 'SUPABASE' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 26:
                importa('DNS', contaDNS, 'SUPABASE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP SEMPRE', vencDNS, idDNS, empDNS,locDNS)
            # NOTAZZ  
            elif 'PG *NOTAZZ GESTAO FISC' in linha[4] and '32289' in linha[2] and int(linha[1][0:2]) == 26:
                importa('DNS', contaDNS, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # CLAUDE  
            elif 'CLAUDE.AI' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 27 and nvalor == '110,00':
                importa('DNS', contaDNS, 'CLAUDE.AI', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # BOLT STACKBLITZ  
            elif 'BOLT (BY STACKBLITZ)' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 25:
                importa('DNS', contaDNS, 'STACKBLITZ, INC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencDNS, idDNS, empDNS,locDNS)
            # CALENDY
            elif 'CALENDLY' in linha[4] and '7886' in linha[2] and int(linha[1][0:2]) == 28:
                importa('DNS', contaDNS, 'CALENDLY LLC', 'INSIDE SALES', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # VERCEL
            elif 'VERCEL INC.' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 29:
                importa('DNS', contaDNS, 'VERCEL INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # ENVATO
            elif 'ENVATO' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 29:
                importa('DNS', contaDNS, 'EVANATO ELEMENTES PTY LTD', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # TOPINVEST   
            elif 'TOPINVEST ED*TOP INVES' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 1 and nvalor == '99,90':
                importa('DNS', contaDNS, 'TOPINVEST EDUCACAO FINANCEIRA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.04.011 LIVROS CURSOS E TREINAMENTOS - CUSTO', 'AUVP PRO', vencDNS, idDNS, empDNS,locDNS)
            # OPENAI 
            elif 'OPENAI' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 2:
                importa('DNS', contaDNS, 'OPENAI,LLC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP SEMPRE', vencDNS, idDNS, empDNS,locDNS)
            # AMAZON AWS
            elif 'Amazon AWS Servicos Br' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 2:
                importa('DNS', contaDNS, 'AMAZON AWS SERVICOS BRASIL LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.002 SERVIÇO DE HOSPEDAGEM E NUVEM', 'OPERAÇÃO & PRODUÇÃO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # BR DID dia 03 e 17 Maurício atendimento
            elif 'PG *BR DID TELEFONIA' in linha[4] and '2889' in linha[2] and ((int(linha[1][0:2]) == 3 and nvalor == '11,90') or (int(linha[1][0:2]) == 17 and nvalor == '29,80')):
                importa('DNS', contaDNS, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'COMERCIAL: DNS', vencDNS, idDNS, empDNS,locDNS)
            # BR DID dia 16 e 19 alyf
            elif 'PG *BR DID TELEFONIA' in linha[4] and '3559' in linha[2] and (int(linha[1][0:2]) == 16 or int(linha[1][0:2]) == 19)and nvalor == '23,90':
                importa('DNS', contaDNS, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # CHATGPT VITÃO
            elif 'OPENAI *CHATGPT SUBSCR' in linha[4] and '2889' in linha[2] and int(linha[1][0:2]) == 7:
                importa('DNS', contaDNS, 'OPENAI,LLC', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'CANAL INVESTIDOR SARDINHA', vencDNS, idDNS, empDNS,locDNS)
            # LOVABLE
            elif 'LOVABLE' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 7:
                importa('DNS', contaDNS, 'TECNOLOGIA E DESENVOLVIMENTO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', vencDNS, idDNS, empDNS,locDNS)
            # USERBACK
            elif 'USERBACK*' in linha[4] and '3559' in linha[2] and int(linha[1][0:2]) == 7:
                importa('DNS', contaDNS, 'USERBACK.IO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'AUVP ANALíTICA', vencDNS, idDNS, empDNS,locDNS)
            
            ### ------------------------- COMPRAS THE BRAIN -------------------------------- ###
            
            # WEBFLOW
            elif 'WEBFLOW.COM' in linha[4] and '7003' in linha[2]:
                importa('THEBRAIN', contaTB, 'WEBFLOW INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # BIGPSY
            elif 'BIGSPY' in linha[4] and '5206' in linha[2] and int(linha[1][0:2]) == 10:
                importa('THEBRAIN', contaTB, 'BIGSPY', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencTB, idTB, empTB,locTB)
            # PADDLE NET CLOUD
            elif 'PADDLE.NET * N8N CLOUD1' in linha[4] and '7003' in linha[2] and int(linha[1][0:2]) == 12 and nvalor == '360,00':
                importa('THEBRAIN', contaTB, 'CLOUD1 SERVICOS DE INFORMATICA LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.002 SERVIÇO DE HOSPEDAGEM E NUVEM', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # CLICKUP DIA 13
            elif 'CLICKUP' in linha[4] and '7003' in linha[2] and int(linha[1][0:2]) == 13:
                importa('THEBRAIN', contaTB, 'ClickUp - Mango Technologies, Inc.', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # TINY
            elif 'TINY ERP' in linha[4] and '5206' in linha[2] and int(linha[1][0:2]) == 20 and nvalor == "149,90":
                importa('THEBRAIN', contaTB, 'OLIST TINY TECNOLOGIA LTDA', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # ADOBE DIA 26 PRODUTO
            elif 'ADOBE' in linha[4] and '8537' in linha[2] and int(linha[1][0:2]) == 26 and nvalor == '139,00':
                importa('THEBRAIN', contaTB, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # NOTAZZ
            elif 'PG *NOTAZZ GESTAO FISC' in linha[4] and '7003' in linha[2] and int(linha[1][0:2]) == 27 and nvalor == "910,90":
                importa('THEBRAIN', contaTB, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # GOOGLE WTF
            elif 'Google GSUITE_wtf.mais' in linha[4] and '7003' in linha[2] and int(linha[1][0:2]) == 1:
                importa('THEBRAIN', contaTB, 'GOOGLE - GSUITE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # MONGODBCLOUD
            elif 'MONGODBCLOUD PAULO' in linha[4] and '7003' in linha[2] and int(linha[1][0:2]) == 2:
                importa('THEBRAIN', contaTB, 'MONGODB SERVICOS DE SOFTWARE NO BRASIL LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # MANYCHAT
            elif 'MANYCHAT.COM' in linha[4] and '5206' in linha[2] and int(linha[1][0:2]) == 3:
                importa('THEBRAIN', contaTB, 'MANYCHAT INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'O SUPERPODER', vencTB, idTB, empTB,locTB)
            # GURU DIA 4 INSIDE SALES
            elif 'GURU-DISCIPULO PLUS 3' in linha[4] and '2599' in linha[2] and int(linha[1][0:2]) == 4:
                importa('THEBRAIN', contaTB, 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA', 'INSIDE SALES', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'O SUPERPODER', vencTB, idTB, empTB,locTB)
            # CANVA
            elif 'EBN *Canva' in linha[4] and '7003' in linha[2] and int(linha[1][0:2]) == 5 and nvalor == '44,90':
                importa('THEBRAIN', contaTB, 'CANVA PTY LTD.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # GOOGLE AMAZOM
            elif 'DL*GOOGLE Amazon' in linha[4] and '5206' in linha[2] and int(linha[1][0:2]) == 6 and nvalor == '19,90':
                importa('THEBRAIN', contaTB, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            # ADOBE DIA 6 AUDIOVISUAL
            elif 'ADOBE' in linha[4] and '5206' in linha[2] and int(linha[1][0:2]) == 6 and nvalor == '275,00':
                importa('THEBRAIN', contaTB, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', vencTB, idTB, empTB,locTB)
            
            
            ### ------------------------- COMPRAS SUPERNOVA -------------------------------- ###

            # GOOGLE GSUITE SUPERNOVA
            elif 'DL *GOOGLE GSUITEasupe' in linha[4] and '5393' in linha[2] and int(linha[1][0:2]) == 6:
                importa('SUPERNOVA', contaSPN, 'GOOGLE - GSUITE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # LOVABLE
            elif 'LOVABLE' in linha[4] and '0799' in linha[2] and int(linha[1][0:2]) == 13:
                importa('SUPERNOVA', contaSPN, 'LOVABLE', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencSPN, idSPN, empSPN,locSPN)
            #  MANUS
            elif 'MANUS AI' in linha[4] and '0799' in linha[2] and int(linha[1][0:2]) == 15:
                importa('SUPERNOVA', contaSPN, 'MANUS AI', 'PRODUTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # OPENAI CONTROLADORIA DIA 17
            elif 'OPENAI *CHATGPT SUBSCR' in linha[4] and '7666' in linha[2] and int(linha[1][0:2]) == 17:
                importa('SUPERNOVA', contaSPN, 'OPENAI,LLC', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # OPENAIO PRODUTO DIA 3 E 19
            elif 'OPENAI *CHATGPT SUBSCR' in linha[4] and '6614' in linha[2] and (int(linha[1][0:2]) == 19 or int(linha[1][0:2]) == 3):
                importa('SUPERNOVA', contaSPN, 'OPENAI,LLC', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # GOOGLE ONE AUDIOVISUAL
            elif 'DL *GOOGLE Google One' in linha[4] and '3341' in linha[2] and int(linha[1][0:2]) == 25:
                importa('SUPERNOVA', contaSPN, 'GOOGLE - GSUITE', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # VERCEL
            elif 'VERCEL INC.' in linha[4] and '5393' in linha[2] and int(linha[1][0:2]) == 28:
                importa('SUPERNOVA', contaSPN, 'VERCEL INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # WORKSPACE FACE
            elif 'FACEBK' in linha[4] and '6614' in linha[2] and int(linha[1][0:2]) == 1:
                importa('SUPERNOVA', contaSPN, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # OPENAI JURIDICO CAMILA EMILY
            elif 'OPENAI *CHATGPT SUBSCR' in linha[4] and '6614' in linha[2] and int(linha[1][0:2]) == 4:
                importa('SUPERNOVA', contaSPN, 'OPENAI,LLC', 'JURÍDICO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', vencSPN, idSPN, empSPN,locSPN)
            # SENDGRID 
            elif 'TWILIO SENDGRID' in linha[4] and '6614' in linha[2] and int(linha[1][0:2]) == 3:
                importa('SUPERNOVA', contaSPN, 'TWILIO EXPANSION LLC', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', vencSPN, idSPN, empSPN,locSPN)

            
            
            ### ------------------------- COMPRAS SUPERNOVA -------------------------------- ###

            # NOTAZZ
            elif 'PG *NOTAZZ GESTAO FISC' in linha[4] and '9342' in linha[2] and int(linha[1][0:2]) == 15 and nvalor == '728,90':
                importa('AUVP', contaAUVP, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'AUVP CONSULTORIA', vencAUVP, idAUVP, empAUVP,locAUVP)
            # INVESTING
            elif 'INVESTING.COM' in linha[4] and '9033' in linha[2] and int(linha[1][0:2]) == 22 and nvalor == '99,00':
                importa('AUVP', contaAUVP, 'INVESTING.COM', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.003 SERVIÇO DE ACESSO A CONTEÚDO', 'AUVP CONSULTORIA', vencAUVP, idAUVP, empAUVP,locAUVP)
            # BR DID
            elif 'PG *BR DID TELEFONIA' in linha[4] and '3428' in linha[2] and int(linha[1][0:2]) == 25 and nvalor == '23,90':
                importa('AUVP', contaAUVP, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', vencAUVP, idAUVP, empAUVP,locAUVP)
            # TURBOSCRIBE
            elif 'TURBOSCRIBE' in linha[4] and '3428' in linha[2] and int(linha[1][0:2]) == 29:
                importa('AUVP', contaAUVP, 'TURBOSCRIBE', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', vencAUVP, idAUVP, empAUVP,locAUVP)
            # OPUSCLIP
            elif 'OPUS CLIP' in linha[4] and '9342' in linha[2] and int(linha[1][0:2]) == 6:
                importa('AUVP', contaAUVP, 'OPUS CLIP', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', vencAUVP, idAUVP, empAUVP,locAUVP)
            # GURU APRENDIZ
            elif 'GURU-APRENDIZ-II' in linha[4] and '9342' in linha[2] and int(linha[1][0:2]) == 9:
                importa('AUVP', contaAUVP, 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP CONSULTORIA', vencAUVP, idAUVP, empAUVP,locAUVP)
            # AMAZON PRIME
            elif 'AmazonPrimeBR' in linha[4] and '3306' in linha[2] and int(linha[1][0:2]) == 11 and nvalor == '19,90':
                importa('AUVP', contaAUVP, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'CONTROLADORIA E FINANÇAS', '3.5.1.05.023 OUTRAS DESPESAS ADMINISTRATIVAS', 'ADMINISTRATIVO: AUVP CONSULTORIA', vencAUVP, idAUVP, empAUVP,locAUVP)
            # RECLAME AQUI
            elif 'RECLAMEAQUI' in linha[4] and '3428' in linha[2] and int(linha[1][0:2]) == 15 and nvalor == '49,00':
                importa('AUVP', contaAUVP, 'OBVIO BRASIL SOFTWARE E SERVICOS S.A.', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', vencAUVP, idAUVP, empAUVP,locAUVP)

            
            

            else:
                importa('DNS', contaDNS, '', '', '', '', vencDNS, idDNS, empDNS,locDNS)
                
df = pd.DataFrame.from_dict(dados)
df.to_excel("teste.xlsx", index = False)







        ###CRIAR A LINHA NO ARQUIVO:
            # EMPRESA("DNS")
            # CONTA ("2.1.2.01.042 CLARA CARTÃO - DNS")
            # MEMORANDO 'CARTÃO' + NÚMERO DO CARTÃO(linha[2]) + '-' + MEMORANDO(linha[4])
            # ENTIDADE ("0" POR ENQUANTO< DEPOIS LAPIDA PARA CASOS E CASOS)
            # DEPARTAMENTO ("0" POR ENQUANTO< DEPOIS LAPIDA PARA CASOS E CASOS)
            # CONTA DESPESAS ("0" POR ENQUANTO< DEPOIS LAPIDA PARA CASOS E CASOS)
            # CLASSE ("0" POR ENQUANTO< DEPOIS LAPIDA PARA CASOS E CASOS)
            # DATA(linha[0])  
            # PERÍODO CONTÁBIL
            # VENCIMENTO ("15/{mês}/2025")
            # VALOR TXT (f"{nvalor}")
            # VALOR (f"{nvalor}")
            # Nº REF (f'="Cartão" & "_" & "2025." & H{linha} & ".00" & LIN(H{linha})')
            # ID CONTA ("1828")
            # ID CONTA DESPESA(f"=PROCX(F{linha};Planodecontas!$E:$E;Planodecontas!$A:$A;)")
            # ID FORNECEDOR (f"=PROCX(D{linha};Fornecedores!$B:$B;Fornecedores!$A:$A;)")
            # ID EMPRESA ("7")
            # ID CLASSE (f"=PROCX(G{linha};classes!$B:$B;classes!$A:$A;)")
            # ID DEPARTAMENTO (f"=PROCX(E{linha};Departamentos!$B:$B;Departamentos!$A:$A;)")    

