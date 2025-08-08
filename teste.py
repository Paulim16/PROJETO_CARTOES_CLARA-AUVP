import csv
import datetime

from datetime import datetime, timedelta, date
import pandas as pd
from dateutil.relativedelta import relativedelta
data_de_hoje = date.today()
mes_atual = data_de_hoje.month
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
vencDNS = f"15/0{mes_atual}/2025"
vencTB = f"15/0{mes_atual}/2025"
vencSPN = f"12/0{mes_atual}/2025"
vencAUVP = f"22/0{mes_atual}/2025"
empDNS = '7'
empTB = '10'
empSPN = '6'
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
    dados['MEMORANDO'].append(f'CARTÃO {linha[6]} - {linha[2]}')
    dados['ENTIDADE'].append(entidade)
    dados['DEPARTAMENTO'].append(departamento)
    dados['CONTA DESPESAS'].append(contadespesa)
    dados['CLASSE'].append(classe)
    dados['DATA'].append(linha[0])        
    dados['PERÍODO CONTÁBIL'].append(f"01/{str(linha[0][5:7])}/2025")
    dados['VENCIMENTO'].append(vencimento)
    dados['VALOR TXT'].append(nvalor) 
    dados['VALOR'].append(nvalor)
    dados['Nº REF'].append(f'="Cartão" & "_" & "2025." & H{i + 1} & ".00" & LIN(H{i + 1})') #f'="Cartão" & "_" & "2025." & H{linha} & ".00" & LIN(H{linha})'
    dados['ID CONTA'].append(idconta)
    dados['ID CONTA DESPESA'].append(f"=PROCX(F{i + 1};Planodecontas!$E:$E;Planodecontas!$A:$A;)")  #f"=PROCX(F{linha};Planodecontas!$E:$E;Planodecontas!$A:$A;)"
    dados['ID FORNECEDOR'].append(f"=PROCX(D{i + 1};Fornecedores!$B:$B;Fornecedores!$A:$A;)")  #f"=PROCX(D{linha};Fornecedores!$B:$B;Fornecedores!$A:$A;)"
    dados['ID EMPRESA'].append(idempresa)
    dados['ID CLASSE'].append(f"=PROCX(G{i + 1};classes!$B:$B;classes!$A:$A;)")  #f"=PROCX(G{linha};classes!$B:$B;classes!$A:$A;)"
    dados['ID DEPARTAMENTO'].append(f"=PROCX(E{i + 1};Departamentos!$B:$B;Departamentos!$A:$A;)")  #f"=PROCX(E{linha};Departamentos!$B:$B;Departamentos!$A:$A;)"
    dados['ID LOCALIDADE'].append(idloc)

with open('csv.csv','r') as cartao:
    leitor = csv.reader(cartao, delimiter=",")
   

    for i, linha in enumerate(leitor):
        
        
        if i >= 1:

            nvalor = linha[5].replace('.', ',')

            data = datetime.strptime(linha[0], "%Y-%m-%d")
            mes_seguinte = data + relativedelta(months=1)

            entidade = ''
            departamento = ''
            contaDespesa = ''

            if linha[6] in ['2289', '5518' ,'3559' , '8389' , '2236' , '6693' , '5800' , '8409' , '5913' , '1582' , '2843' ,'9802' , '1686' , '4836' , '4952' , '5786' , '8549' , '7886' , '5168' , '5570' , '8773' , '3280']:
                empresa = 'DNS'
                vencimento = vencDNS
                conta = contaDNS
                id = idDNS
                loc = locDNS
                emp = empDNS

            elif linha[6] in  ['7003'  , '8537' ,'5206' , '2599' , '6187' , '8534' , '4123' , '3484' , '1011' , '5790' , '8058' , '2508' , '9481' , '1145' , '5874' , '4758']:
                empresa = 'THEBRAIN'
                vencimento = vencTB
                conta = contaTB
                id = idTB
                loc = locTB
                emp = empTB

            elif linha[6] in ['7737'  , '9074' , '0799' , '3341' , '3342' , '0126' , '6614' , '5409' , '6065' , '7666' , '5393' , '9926' , '9316' , '6761']:
                empresa = 'SUPERNOVA'
                vencimento = vencSPN
                conta = contaSPN
                id = idSPN
                loc = locSPN
                emp = empSPN

            elif linha[6] in ['4388'  , '9033' , '1450' , '7256' , '8931' , '3428' , '5550' , '0156' , '9342' , '6055' , '0577' , '3791' , '0694' , '9544' , '8405' , '5129' , '5678', '9637' , '1588' , '9244' , '3288' , '3306' , '6613']:
                empresa = 'AUVP CONSULTORIA'
                vencimento = vencAUVP
                conta = contaAUVP
                id = idAUVP
                loc = locAUVP
                emp = empAUVP
            

            if int(linha[0][8:10]) >= (int(vencimento[0:2]) - 6):
                month = mes_seguinte.month
                year = mes_seguinte.year

            else:
                month = data.month
                year = data.year
             

            venc = datetime.strftime(datetime(day = int(vencimento[0:2]), month = month, year=year),"%d/%m/%Y")




            ###---------------------------------------------------------------------------------###
            ###------------------------ DEPARTAMENTOS E ENTIDADES-------------------------------###
            ###---------------------------------------------------------------------------------###
#       Aqui é onde ocorre a seleção para preencher, na planilha de importação, a aba de Departamentos, de acordo com o usuário do cartão (suscetível à alterações na validação),
#   e a aba de ENTIDADES de acordo com o Memorando que está contido no CSV das transações do cartão.
            ###------------------------------- DEPARTAMENTOS -----------------------------------###

            if 'Alyf' in linha[16]:
                departamento = 'TECNOLOGIA E DESENVOLVIMENTO'
            elif 'Beatriz  Henriques' in linha[16]:
                departamento = 'PRODUTO'
            elif 'Mauricio  Imparato' in linha[16]:
                departamento = 'ATENDIMENTO E CX'
            elif 'Lucas  Cassimiro' in linha[16]:
                departamento = 'PRODUÇÃO AUDIOVISUAL'
            elif 'Bruna  Alencar' in linha[16]:
                departamento = 'CAPITAL HUMANO'
            elif 'Brenner   Nepomuceno' in linha[16]:
                departamento = 'CONSULTORIA E INVESTIMENTOS'

        ###------------------------------- ENTIDADES -----------------------------------###
            
            if 'FACEBK' in linha[2]:
                entidade = 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA'
            elif 'WP MEDIA - IMAGIFY' in linha[2]:
                entidade = 'IMAGIFY'
            elif 'OPENAI *CHATGPT SUBSCR' in linha[2]:
                entidade = 'OPENAI,LLC'
            elif linha[2] in ['AmazonPrimeBR', 'AMAZON BR']:
                entidade = 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.'
            elif 'FUNDAMENTEI.COM' in linha[2]:
                entidade = 'FUNDAMENTEI SERVICOS DE INFORMACAO LTDA'
            elif 'BRAPI.DEV' in linha[2]:
                entidade = 'BRAPI ASL TECNOLOGIA LTDA'
            elif 'STAPE, INC' in linha[2]:
                entidade = 'STAPE, INC.'
            elif 'EBN *Canva' in linha[2]:
                entidade = 'CANVA PTY LTD.'
            elif 'LEARNWORLDS CY LTD' in linha[2]:
                entidade = 'LEARNWORLDS (CY) LTD'
            elif 'WINDSOR.AI' in linha[2]:
                entidade = 'WINDSOR.AI'
            elif 'AOVS SISTEMAS DE INFOR' in linha[2]:
                entidade = 'AOVS SISTEMAS DE INFORMATICA SA'
            elif 'ELEVENLABS.IO' in linha[2]:
                entidade = 'ELEVENLABS.IO'
            elif 'MSFT *' in linha[2]:
                entidade = 'MICROSOFT INFORMATICA LTDA'
            elif 'NOTIFICACOES INTELIGEN' in linha[2]:
                entidade = 'KIWIFY PAGAMENTOS, TECNOLOGIA E SERVICOS LTDA'
            elif 'UAZAPI - API WHATSAPP' in linha[2]:
                entidade = 'UAZAPI'
            elif 'TINY ERP' in linha[2]:
                entidade = 'OLIST TINY TECNOLOGIA LTDA'
            elif 'ADOBE' in linha[2]:
                entidade = 'ADOBE SYSTEMS BRASIL LTDA.'
            elif 'EBN *SEMRUSH' in linha[2]:
                entidade = 'SEMRUSH'
            elif 'FIGMA' in linha[2]:
                entidade = 'FIGMA MONTHLY RENEWAL'
            elif 'BITLY.COM' in linha[2]:
                entidade = 'BITLY COM'
            elif 'MANYCHAT.COM' in linha[2]:
                entidade = 'MANYCHAT INC.'
            elif 'SUPABASE' in linha[2]:
                entidade = 'SUPABASE'
            elif 'PG *NOTAZZ GESTAO FISC' in linha[2]:
                entidade = 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA'
            elif 'CLAUDE.AI' in linha[2] and '3559':
                entidade = 'CLAUDE.AI'
            elif 'STACKBLITZ' in linha[2]:
                entidade = 'STACKBLITZ, INC'
            elif 'CALENDLY' in linha[2]:
                entidade = 'CALENDLY LLC'    
            elif 'VERCEL INC.' in linha[2]:
                entidade = 'VERCEL INC.'
            elif 'ENVATO' in linha[2]:
                entidade = 'EVANATO ELEMENTES PTY LTD'
            elif 'TOPINVEST ED*TOP INVES' in linha[2]:
                entidade = 'TOPINVEST EDUCACAO FINANCEIRA LTDA'
            elif 'OPENAI' in linha[2]:
                entidade = 'OPENAI,LLC'
            elif 'Amazon AWS Servicos Br' in linha[2]:
                entidade = 'AMAZON AWS SERVICOS BRASIL LTDA'
            elif 'PG *BR DID TELEFONIA' in linha[2]:
                entidade = 'BR TECH TECNOLOGIA EM SISTEMAS LTDA'
            elif 'LOVABLE' in linha[2]:
                entidade = 'LOVABLE'
            elif 'USERBACK*' in linha[2]:
                entidade = 'USERBACK.IO'
            elif 'WEBFLOW.COM' in linha[2]:
                entidade = 'WEBFLOW INC.'
            elif 'BIGSPY' in linha[2]:
                entidade = 'BIGSPY'
            elif 'PADDLE.NET * N8N CLOUD1' in linha[2]:
                entidade = 'CLOUD1 SERVICOS DE INFORMATICA LTDA.'
            elif 'CLICKUP' in linha[2]:
                entidade = 'ClickUp - Mango Technologies, Inc.'
            elif linha [2] in ['Google GSUITE_wtf.mais', 'DL *GOOGLE GSUITEasupe' , 'DL *GOOGLE Google One']:
                entidade = 'GOOGLE - GSUITE'
            elif 'MONGODBCLOUD PAULO' in linha[2]:
                entidade = 'MONGODB SERVICOS DE SOFTWARE NO BRASIL LTDA.'
            elif 'GURU-DISCIPULO PLUS 3' in linha[2]:
                entidade = 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA'
            elif 'EBN *Canva' in linha[2]:
                entidade = 'CANVA PTY LTD.'
            elif 'DL*GOOGLE Amazon' in linha[2]:
                entidade = 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.'
            elif 'MANUS AI' in linha[2]:
                entidade = 'MANUS AI'
            elif 'TWILIO SENDGRID' in linha[2]:
                entidade = 'TWILIO EXPANSION LLC'
                entidade = 'INVESTING.COM'
            elif 'TURBOSCRIBE' in linha[2]:
                entidade = 'TURBOSCRIBE'
            elif 'OPUS CLIP' in linha[2]:
                entidade = 'OPUS CLIP'
            elif 'RECLAMEAQUI' in linha[2]:
                entidade = 'OBVIO BRASIL SOFTWARE E SERVICOS S.A.'
            elif 'URUAQUE' in linha[2]:
                entidade = 'URUAQUE GOIANIA LTDA'
            elif 'Uber UBER *TRIP HELP.U' in linha[2]:
                entidade = 'UBER DO BRASIL TECNOLOGIA LTDA.'
            elif 'LATAM AIR' in linha[2]:
                entidade = 'LATAM AIRLINES BRASIL'
            elif 'MERCADOLIVRE*' in linha[2]:
                entidade = 'MERCADOLIVRE.COM'
            elif 'OPUS CLIP' in linha[2]:
                entidade = 'OPUS CLIP'
            elif 'IOF - COMPRA INTERNACIONAL' in linha[2]:
                contaDespesa = '3.6.1.01.005 IOF'
           

            ###---------------------------------------------------------------------------------###
            ###--------------------------- COMPRAS RECORRENTES ---------------------------------###
            ###---------------------------------------------------------------------------------###

# Aqui é onde ocorre a seleção para preencher, na planilha de importação, algumas abas de acordo com as assinaturas/compras/pagamentos que são feitos recorrentemente nos cartões, 
# que, assim sendo, já contêm as abas de Entidade, Departamento, Centro de Custo e Classe já bem definidos.

            ### ------------------------- COMPRAS DNS -------------------------------- ###
            # FACEBOOK
            if 'FACEBK' in linha[2] and '2289' in linha[6]: 
                importa(empresa, conta, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'PUBLICIDADE', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'AUVP ESCOLA', venc, id, emp,loc)
              
            elif 'FACEBK' in linha[2] and '4836' in linha[6]:
                importa(empresa, conta, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'PUBLICIDADE', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'AUVP ANALÍTICA', venc, id, emp,loc)  
            # Imagify
            elif 'WP MEDIA - IMAGIFY' in linha[2] and '8389' in linha[6] and int(linha[0][8:10]) == 9:
                importa(empresa, conta, 'IMAGIFY', 'PUBLICIDADE', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # Openai dia 09 e 23 DNS - alyf
            elif 'OPENAI *CHATGPT SUBSCR' in linha[2] and '3559' in linha[6] and (int(linha[0][8:10]) == 9 or int(linha[0][8:10]) == 23) :
                importa(empresa, conta, 'OPENAI,LLC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            #Amazon Prime
            elif 'AmazonPrimeBR' in linha[2] and '8389' in linha[6] and int(linha[0][8:10]) == 10 and nvalor == '19,9':
                importa(empresa, conta, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'CONTROLADORIA E FINANÇAS', '3.5.1.05.023 OUTRAS DESPESAS ADMINISTRATIVAS', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # Fundamentei
            elif 'FUNDAMENTEI.COM' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 10 and nvalor == '49,0':
                importa(empresa, conta, 'FUNDAMENTEI SERVICOS DE INFORMACAO LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # Brapi Dev 
            elif 'BRAPI.DEV' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 10 and nvalor == '49,99':
                importa(empresa, conta, 'BRAPI ASL TECNOLOGIA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # STAPE INC
            elif 'STAPE, INC' in linha[2] and '2289' in linha[6] and 9 <= int(linha[0][8:10]) == 11:
                importa(empresa, conta, 'STAPE, INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # EBN Canva
            elif 'EBN *Canva' in linha[2] and '8409' in linha[6] and int(linha[0][8:10]) == 13 and nvalor == '174,5':
                importa(empresa, conta, 'CANVA PTY LTD.', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # LEARNWORLDS CY LTD dia 13, 22 e 29
            elif 'LEARNWORLDS CY LTD' in linha[2] and '2289' in linha[6] and (int(linha[0][8:10]) == 13 or int(linha[0][8:10]) == 22 or int(linha[0][8:10]) == 29):
                importa(empresa, conta, 'LEARNWORLDS (CY) LTD', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # WINDSOR.AI
            elif 'WINDSOR.AI' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 14:
                importa(empresa, conta, 'WINDSOR.AI', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # AOVS
            elif 'AOVS SISTEMAS DE INFOR' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 16 and nvalor == '530,0':
                importa(empresa, conta, 'AOVS SISTEMAS DE INFORMATICA SA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # ELEVENLABS
            elif 'ELEVENLABS.IO' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 17:
                importa(empresa, conta, 'ELEVENLABS.IO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # MICROSOFT
            elif 'MSFT *' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 18:
                importa(empresa, conta, 'MICROSOFT INFORMATICA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # KIWIFY NOTIFICA
            elif 'NOTIFICACOES INTELIGEN' in linha[2] and '2236' in linha[6] and int(linha[0][8:10]) == 18:
                importa(empresa, conta, 'KIWIFY PAGAMENTOS, TECNOLOGIA E SERVICOS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc) 
            # UAZAPI
            elif 'UAZAPI - API WHATSAPP' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 19 and nvalor == '29,0' :
                importa(empresa, conta, 'UAZAPI', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # STAPE
            elif 'STAPE, INC.' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 19:
                importa(empresa, conta, 'STAPE, INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # TINY
            elif 'TINY ERP' in linha[2] and '8389' in linha[6] and int(linha[0][8:10]) == 20:
                importa(empresa, conta, 'OLIST TINY TECNOLOGIA LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # ADOBE DIA 20 e 06   
            elif 'ADOBE' in linha[2] and '2289' in linha[6] and (int(linha[0][8:10]) == 20 or int(linha[0][8:10]) == 6) and nvalor == '275,0':
                importa(empresa, conta, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # SEMRUSH   
            elif 'EBN *SEMRUSH' in linha[2] and '8389' in linha[6] and int(linha[0][8:10]) == 21:
                importa(empresa, conta, 'SEMRUSH', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # FIGMA
            elif 'FIGMA' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 21:
                importa(empresa, conta, 'FIGMA MONTHLY RENEWAL', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # BITLy  
            elif 'BITLY.COM' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 22:
                importa(empresa, conta, 'BITLY COM', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # MANYCHAT    
            elif 'MANYCHAT.COM' in linha[2] and '5518' in linha[6] and int(linha[0][8:10]) == 25:
                importa(empresa, conta, 'MANYCHAT INC.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # SUPABASE    
            elif 'SUPABASE' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 26:
                importa(empresa, conta, 'SUPABASE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP SEMPRE', venc, id, emp,loc)
            # NOTAZZ  
            elif 'PG *NOTAZZ GESTAO FISC' in linha[2] and '2289' in linha[6] and int(linha[0][8:10]) == 26:
                importa(empresa, conta, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # CLAUDE  
            elif 'CLAUDE.AI' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 27 and nvalor == '110,0':
                importa(empresa, conta, 'CLAUDE.AI', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # BOLT STACKBLITZ  
            elif 'BOLT (BY STACKBLITZ)' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 27:
                importa(empresa, conta, 'STACKBLITZ, INC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # CALENDY
            elif 'CALENDLY' in linha[2] and '7886' in linha[6] and int(linha[0][8:10]) == 28:
                importa(empresa, conta, 'CALENDLY LLC', 'INSIDE SALES', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # VERCEL
            elif 'VERCEL INC.' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 29:
                importa(empresa, conta, 'VERCEL INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # ENVATO
            elif 'ENVATO' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 29:
                importa(empresa, conta, 'EVANATO ELEMENTES PTY LTD', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # TOPINVEST   
            elif 'TOPINVEST ED*TOP INVES' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 1 and nvalor == '99,9':
                importa(empresa, conta, 'TOPINVEST EDUCACAO FINANCEIRA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.04.011 LIVROS CURSOS E TREINAMENTOS - CUSTO', 'AUVP PRO', venc, id, emp,loc)
            # OPENAI 
            elif 'OPENAI' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 2:
                importa(empresa, conta, 'OPENAI,LLC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP SEMPRE', venc, id, emp,loc)
            # AMAZON AWS
            elif 'Amazon AWS Servicos Br' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 2:
                importa(empresa, conta, 'AMAZON AWS SERVICOS BRASIL LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.002 SERVIÇO DE HOSPEDAGEM E NUVEM', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # BR DID dia 03 e 17 Maurício atendimento
            elif 'PG *BR DID TELEFONIA' in linha[2] and '2889' in linha[6] and ((int(linha[0][8:10]) == 3 and nvalor == '11,9') or (int(linha[0][8:10]) == 17 and nvalor == '29,8')):
                importa(empresa, conta, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'COMERCIAL: DNS', venc, id, emp,loc)
            # BR DID dia 16 e 19 alyf
            elif 'PG *BR DID TELEFONIA' in linha[2] and '3559' in linha[6] and (int(linha[0][8:10]) == 16 or int(linha[0][8:10]) == 19)and nvalor == '23,9':
                importa(empresa, conta, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # CHATGPT VITÃO
            elif 'OPENAI *CHATGPT SUBSCR' in linha[2] and '2889' in linha[6] and int(linha[0][8:10]) == 7:
                importa(empresa, conta, 'OPENAI,LLC', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'CANAL INVESTIDOR SARDINHA', venc, id, emp,loc)
            # LOVABLE
            elif 'LOVABLE' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 7:
                importa(empresa, conta, 'LOVABLE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # USERBACK
            elif 'USERBACK*' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 7:
                importa(empresa, conta, 'USERBACK.IO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'AUVP ANALíTICA', venc, id, emp,loc)
            #RAILWAY
            elif 'RAILWAY' in linha[2] and '3559' in linha[6] and int(linha[0][8:10]) == 4:
                importa(empresa, conta, 'RAILWAY CORP.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            
            ### ------------------------- COMPRAS THE BRAIN -------------------------------- ###
            
            # WEBFLOW
            elif 'WEBFLOW.COM' in linha[2] and '7003' in linha[6]:
                importa(empresa, conta, 'WEBFLOW INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # BIGPSY
            elif 'BIGSPY' in linha[2] and '5206' in linha[6] and int(linha[0][8:10]) == 10:
                importa(empresa, conta, 'BIGSPY', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # PADDLE NET CLOUD
            elif 'PADDLE.NET * N8N CLOUD1' in linha[2] and '7003' in linha[6] and int(linha[0][8:10]) == 12 and nvalor == '360,0':
                importa(empresa, conta, 'CLOUD1 SERVICOS DE INFORMATICA LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.002 SERVIÇO DE HOSPEDAGEM E NUVEM', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # CLICKUP DIA 13
            elif 'CLICKUP' in linha[2] and '7003' in linha[6] and int(linha[0][8:10]) == 13:
                importa(empresa, conta, 'ClickUp - Mango Technologies, Inc.', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', venc, id, emp,loc)
            # ZOOM
            elif 'ZOOM.COM 888-799-9666' in linha[2] and '6187' in linha[6] and int(linha[0][8:10]) == 17:
                importa(empresa, conta, 'ClickUp - Mango Technologies, Inc.', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', venc, id, emp,loc)
            # TINY
            elif 'TINY ERP' in linha[2] and '5206' in linha[6] and int(linha[0][8:10]) == 20 and nvalor == "149,9":
                importa(empresa, conta, 'OLIST TINY TECNOLOGIA LTDA', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # ADOBE DIA 26 PRODUTO
            elif 'ADOBE' in linha[2] and '8537' in linha[6] and int(linha[0][8:10]) == 26 and nvalor == '139,0':
                importa(empresa, conta, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # NOTAZZ
            elif 'PG *NOTAZZ GESTAO FISC' in linha[2] and '7003' in linha[6] and int(linha[0][8:10]) == 27 and nvalor == "910,9":
                importa(empresa, conta, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', venc, id, emp,loc)
            # GOOGLE WTF
            elif 'Google GSUITE_wtf.mais' in linha[2] and '7003' in linha[6] and int(linha[0][8:10]) == 1:
                importa(empresa, conta, 'GOOGLE - GSUITE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # MONGODBCLOUD
            elif 'MONGODBCLOUD PAULO' in linha[2] and '7003' in linha[6] and int(linha[0][8:10]) == 2:
                importa(empresa, conta, 'MONGODB SERVICOS DE SOFTWARE NO BRASIL LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # MANYCHAT
            elif 'MANYCHAT.COM' in linha[2] and '5206' in linha[6] and int(linha[0][8:10]) == 3:
                importa(empresa, conta, 'MANYCHAT INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'O SUPERPODER', venc, id, emp,loc)
            # GURU DIA 4 INSIDE SALES
            elif 'GURU-DISCIPULO PLUS 3' in linha[2] and '2599' in linha[6] and int(linha[0][8:10]) == 4:
                importa(empresa, conta, 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA', 'INSIDE SALES', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'O SUPERPODER', venc, id, emp,loc)
            # CANVA
            elif 'EBN *Canva' in linha[2] and '7003' in linha[6] and int(linha[0][8:10]) == 5 and nvalor == '44,9':
                importa(empresa, conta, 'CANVA PTY LTD.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # GOOGLE AMAZON
            elif 'DL*GOOGLE Amazon' in linha[2] and '5206' in linha[6] and int(linha[0][8:10]) == 6 and nvalor == '19,9':
                importa(empresa, conta, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # ADOBE DIA 6 AUDIOVISUAL
            elif 'ADOBE' in linha[2] and '5206' in linha[6] and int(linha[0][8:10]) == 6 and nvalor == '275,0':
                importa(empresa, conta, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            
            
            ### ------------------------- COMPRAS SUPERNOVA -------------------------------- ###

            # GOOGLE GSUITE SUPERNOVA
            elif 'DL *GOOGLE GSUITEasupe' in linha[2] and '5393' in linha[6] and int(linha[0][8:10]) == 6:
                importa(empresa, conta, 'GOOGLE - GSUITE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # LOVABLE
            elif 'LOVABLE' in linha[2] and '0799' in linha[6] and (int(linha[0][8:10]) == 13 or int(linha[0][8:10]) == 1):
                importa(empresa, conta, 'LOVABLE', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            #  MANUS
            elif 'MANUS AI' in linha[2] and '0799' in linha[6] and int(linha[0][8:10]) == 15:
                importa(empresa, conta, 'MANUS AI', 'PRODUTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # OPENAI CONTROLADORIA DIA 17
            elif 'OPENAI *CHATGPT SUBSCR' in linha[2] and '7666' in linha[6] and int(linha[0][8:10]) == 17:
                importa(empresa, conta, 'OPENAI,LLC', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # OPENAIO PRODUTO DIA 3 E 19
            elif 'OPENAI *CHATGPT SUBSCR' in linha[2] and '6614' in linha[6] and (int(linha[0][8:10]) == 19 or int(linha[0][8:10]) == 3):
                importa(empresa, conta, 'OPENAI,LLC', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # GOOGLE ONE AUDIOVISUAL
            elif 'DL *GOOGLE Google One' in linha[2] and '3341' in linha[6] and int(linha[0][8:10]) == 25:
                importa(empresa, conta, 'GOOGLE - GSUITE', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # VERCEL
            elif 'VERCEL INC.' in linha[2] and '5393' in linha[6] and int(linha[0][8:10]) == 28:
                importa(empresa, conta, 'VERCEL INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # WORKSPACE FACE
            elif 'FACEBK' in linha[2] and '6614' in linha[6] and int(linha[0][8:10]) == 1:
                importa(empresa, conta, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # OPENAI JURIDICO CAMILA EMILY
            elif 'OPENAI *CHATGPT SUBSCR' in linha[2] and '6614' in linha[6] and int(linha[0][8:10]) == 4:
                importa(empresa, conta, 'OPENAI,LLC', 'JURÍDICO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # SENDGRID 
            elif 'TWILIO SENDGRID' in linha[2] and '6614' in linha[6] and int(linha[0][8:10]) == 3:
                importa(empresa, conta, 'TWILIO EXPANSION LLC', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)   
            
            ### ------------------------- COMPRAS AUVP -------------------------------- ###

            # NOTAZZ
            elif 'PG *NOTAZZ GESTAO FISC' in linha[2] and '9342' in linha[6] and int(linha[0][8:10]) == 15 and nvalor == '728,9':
                importa(empresa, conta, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'AUVP CONSULTORIA', venc, id, emp,loc)
            # INVESTING
            elif 'INVESTING.COM' in linha[2] and '9033' in linha[6] and int(linha[0][8:10]) == 22 and nvalor == '99,0':
                importa(empresa, conta, 'INVESTING.COM', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.003 SERVIÇO DE ACESSO A CONTEÚDO', 'AUVP CONSULTORIA', venc, id, emp,loc)
            # BR DID
            elif 'PG *BR DID TELEFONIA' in linha[2] and '3428' in linha[6] and int(linha[0][8:10]) == 25 and nvalor == '23,9':
                importa(empresa, conta, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)
            # TURBOSCRIBE
            elif 'TURBOSCRIBE' in linha[2] and '3428' in linha[6] and int(linha[0][8:10]) == 29:
                importa(empresa, conta, 'TURBOSCRIBE', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)
            # OPUSCLIP
            elif 'OPUS CLIP' in linha[2] and '9342' in linha[6] and int(linha[0][8:10]) == 6:
                importa(empresa, conta, 'OPUS CLIP', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)
            # GURU APRENDIZ
            elif 'GURU-APRENDIZ-II' in linha[2] and '9342' in linha[6] and int(linha[0][8:10]) == 9:
                importa(empresa, conta, 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP CONSULTORIA', venc, id, emp,loc)
            # AMAZON PRIME
            elif 'AmazonPrimeBR' in linha[2] and '3306' in linha[6] and int(linha[0][8:10]) == 11 and nvalor == '19,9':
                importa(empresa, conta, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'CONTROLADORIA E FINANÇAS', '3.5.1.05.023 OUTRAS DESPESAS ADMINISTRATIVAS', 'ADMINISTRATIVO: AUVP CONSULTORIA', venc, id, emp,loc)
            # RECLAME AQUI
            elif 'RECLAMEAQUI' in linha[2] and '3428' in linha[6] and int(linha[0][8:10]) == 15 and nvalor == '49,9':
                importa(empresa, conta, 'OBVIO BRASIL SOFTWARE E SERVICOS S.A.', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)  
            
            else:

                importa(empresa, conta, entidade, departamento, contaDespesa, '', venc, id, emp,loc)
                
df = pd.DataFrame.from_dict(dados)
df.to_excel("teste.xlsx", index = False)