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

dados = {
    'EMPRESA': [],
    'CONTA': [],
    'MEMORANDO': [],
    'ENTIDADE': [],
    'DEPARTAMENTO': [],
    'CONTA DESPESAS': [],
    'CLASSE': [],
    'DATA': [],
    'PERÍODO CONTÁBIL': [],  
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

# Plano de contas por empresa
COMPANY_INFO = {
    "DNS": {
        "nome": "DNS",
        "conta": "2.1.2.01.042 CLARA CARTÃO - DNS",
        "idconta": "1828",
        "loc": "2",
        "emp": "7",
        "vencimento": "22/01/2025"
    },
    "THEBRAIN": {
        "nome": "THEBRAIN",
        "conta": "2.1.2.01.043 CLARA CARTÃO - THE BRAIN",
        "idconta": "1829",
        "loc": "3",
        "emp": "10",
        "vencimento": "22/01/2025"
    },
    "SUPERNOVA": {
        "nome": "SUPERNOVA",
        "conta": "2.1.2.01.045 CLARA CARTÃO - SUPERNOVA",
        "idconta": "1832",
        "loc": "1",
        "emp": "6",
        "vencimento": "22/01/2025"
    },
    "AUVP": {
        "nome": "AUVP CONSULTORIA",
        "conta": "2.1.2.01.044 CLARA CARTÃO - AUVP",
        "idconta": "1831",
        "loc": "6",
        "emp": "16",
        "vencimento": "22/01/2025"
    },
}

# Mapeamento de cartões por empresa (números como strings exatamente como aparecem no CSV)
CARD_GROUPS = {
    "DNS": {'2289','5518','3559','8389','2236','6693','5800','8409','5913','1582','2843','9802','1686','4836','4952','5786','8549','7886','5168','5570','8773','3280'},
    "THEBRAIN": {'7003','8537','5206','2599','6187','8534','4123','3484','1011','5790','8058','2508','9481','1145','5873','4758'},
    "SUPERNOVA": {'7737','9074','0799','3341','3342','0126','6614','5409','6065','7666','5393','9926','9316','6761'},
    "AUVP": {'4388','9033','1450','7256','8931','3428','5550','0156','9342','6055','0577','3791','0694','9544','8405','5129','5678','9637','1588','9244','3288','3306','6613'},
}


            ###---------------------------------------------------------------------------------###
            ###------------------------ DEPARTAMENTOS E ENTIDADES-------------------------------###
            ###---------------------------------------------------------------------------------###
#   Aqui é onde ocorre a seleção para preencher, na planilha de importação, a aba de Departamentos, de acordo com o usuário do cartão (suscetível à alterações na validação),
#   e a aba de ENTIDADES de acordo com o Memorando que está contido no CSV das transações do cartão.

# Departamentos por titular (coluna linha[16], match por substring)

DEPT_BY_TITULAR = {
    'Alyf': 'TECNOLOGIA E DESENVOLVIMENTO',
    'Beatriz  Henriques': 'PRODUTO',
    'Mauricio  Imparato': 'ATENDIMENTO E CX',
    'Lucas  Cassimiro': 'PRODUÇÃO AUDIOVISUAL',
    'Bruna  Alencar': 'CAPITAL HUMANO',
    'Brenner   Nepomuceno': 'CONSULTORIA E INVESTIMENTOS',
}

ENTITY_BY_MEMO = [
    # (condição_substring, entidade)
    ('FACEBK', 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA'),
    ('WP MEDIA - IMAGIFY', 'IMAGIFY'),
    ('OPENAI *CHATGPT SUBSCR', 'OPENAI,LLC'),
    (['AmazonPrimeBR','AMAZON BR'], 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.'),
    ('FUNDAMENTEI.COM', 'FUNDAMENTEI SERVICOS DE INFORMACAO LTDA'),
    ('BRAPI.DEV', 'BRAPI ASL TECNOLOGIA LTDA'),
    ('STAPE, INC', 'STAPE, INC.'),
    ('EBN *Canva', 'CANVA PTY LTD.'),
    ('LEARNWORLDS CY LTD', 'LEARNWORLDS (CY) LTD'),
    ('WINDSOR.AI', 'WINDSOR.AI'),
    ('AOVS SISTEMAS DE INFOR', 'AOVS SISTEMAS DE INFORMATICA SA'),
    ('ELEVENLABS.IO', 'ELEVENLABS.IO'),
    ('MSFT *', 'MICROSOFT INFORMATICA LTDA'),
    ('NOTIFICACOES INTELIGEN', 'KIWIFY PAGAMENTOS, TECNOLOGIA E SERVICOS LTDA'),
    ('UAZAPI - API WHATSAPP', 'UAZAPI'),
    ('TINY ERP', 'OLIST TINY TECNOLOGIA LTDA'),
    ('ADOBE', 'ADOBE SYSTEMS BRASIL LTDA.'),
    ('EBN *SEMRUSH', 'SEMRUSH'),
    ('FIGMA', 'FIGMA MONTHLY RENEWAL'),
    ('BITLY.COM', 'BITLY COM'),
    ('MANYCHAT.COM', 'MANYCHAT INC.'),
    ('SUPABASE', 'SUPABASE'),
    ('PG *NOTAZZ GESTAO FISC', 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA'),
    ('STACKBLITZ', 'STACKBLITZ, INC'),
    ('CALENDLY', 'CALENDLY LLC'),
    ('VERCEL INC.', 'VERCEL INC.'),
    ('ENVATO', 'EVANATO ELEMENTES PTY LTD'),
    ('TOPINVEST ED*TOP INVES', 'TOPINVEST EDUCACAO FINANCEIRA LTDA'),
    ('OPENAI', 'OPENAI,LLC'),
    ('Amazon AWS Servicos Br', 'AMAZON AWS SERVICOS BRASIL LTDA'),
    ('PG *BR DID TELEFONIA', 'BR TECH TECNOLOGIA EM SISTEMAS LTDA'),
    ('LOVABLE', 'LOVABLE'),
    ('USERBACK*', 'USERBACK.IO'),
    ('WEBFLOW.COM', 'WEBFLOW INC.'),
    ('BIGSPY', 'BIGSPY'),
    ('PADDLE.NET * N8N CLOUD1', 'CLOUD1 SERVICOS DE INFORMATICA LTDA.'),
    ('CLICKUP', 'ClickUp - Mango Technologies, Inc.'),
    (['Google GSUITE_wtf.mais','DL *GOOGLE GSUITEasupe','DL *GOOGLE Google One'],'GOOGLE - GSUITE'),
    ('MONGODBCLOUD PAULO', 'MONGODB SERVICOS DE SOFTWARE NO BRASIL LTDA.'),
    ('GURU-DISCIPULO PLUS 3', 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA'),
    ('DL*GOOGLE Amazon','AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.'),
    ('MANUS AI','MANUS AI'),
    ('TWILIO SENDGRID','TWILIO EXPANSION LLC'),
    ('TURBOSCRIBE','TURBOSCRIBE'),
    ('OPUS CLIP','OPUS CLIP'),
    ('RECLAMEAQUI','OBVIO BRASIL SOFTWARE E SERVICOS S.A.'),
    ('URUAQUE','URUAQUE GOIANIA LTDA'),
    ('Uber UBER *TRIP HELP.U','UBER DO BRASIL TECNOLOGIA LTDA.'),
    ('LATAM AIR','LATAM AIRLINES BRASIL'),
    ('MERCADOLIVRE*','MERCADOLIVRE.COM'),
]

# ----------------------------- HELPERS -----------------------------

def company_by_card(card_number: str):
    """Retorna (key, info) da empresa a partir do número do cartão, ou (None, None)."""
    for key, cards in CARD_GROUPS.items():
        if card_number in cards:
            return key, COMPANY_INFO[key]
    return None, None

def guess_department(holder: str) -> str:
    if holder is None:
        return ''
    for k, v in DEPT_BY_HOLDER.items():
        if k in holder:
            return v
    return ''

def guess_entity(memo: str, card: str) -> str:
    if memo is None:
        return ''
    for pattern, entity in ENTITY_BY_MEMO:
        if isinstance(pattern, list):
            if memo in pattern or any(pat in memo for pat in pattern):
                return entity
        else:
            if pattern in memo:
                return entity
    return ''

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

with open('auvp.csv','r') as cartao:
    leitor = csv.reader(cartao, delimiter=",")
   

    for i, linha in enumerate(leitor):
        
        
        if i >= 1:

            nvalor = linha[5].strip().replace('"', '').replace('.', ',')

            data = datetime.strptime(linha[0], "%Y-%m-%d")
            mes_seguinte = data + relativedelta(months=1)

            card = linha[6]
            memo = linha [2]
            titular = linha[16]
            dia = linha[0][8:10]
            

            entidade = ''
            departamento = ''
            contaDespesa = ''
            

            if int(dia) >= (int(vencimento[0:2]) - 6):
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

            if 'Alyf' in titular:
                departamento = 'TECNOLOGIA E DESENVOLVIMENTO'
            elif 'Beatriz  Henriques' in titular:
                departamento = 'PRODUTO'
            elif 'Mauricio  Imparato' in titular:
                departamento = 'ATENDIMENTO E CX'
            elif 'Lucas  Cassimiro' in titular:
                departamento = 'PRODUÇÃO AUDIOVISUAL'
            elif 'Bruna  Alencar' in titular:
                departamento = 'CAPITAL HUMANO'
            elif 'Brenner   Nepomuceno' in titular:
                departamento = 'CONSULTORIA E INVESTIMENTOS'

        ###------------------------------- ENTIDADES -----------------------------------###
            
            if 'FACEBK' in memo:
                entidade = 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA'
            elif 'WP MEDIA - IMAGIFY' in memo:
                entidade = 'IMAGIFY'
            elif 'OPENAI *CHATGPT SUBSCR' in memo:
                entidade = 'OPENAI,LLC'
            elif memo in ['AmazonPrimeBR', 'AMAZON BR']:
                entidade = 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.'
            elif 'FUNDAMENTEI.COM' in memo:
                entidade = 'FUNDAMENTEI SERVICOS DE INFORMACAO LTDA'
            elif 'BRAPI.DEV' in memo:
                entidade = 'BRAPI ASL TECNOLOGIA LTDA'
            elif 'STAPE, INC' in memo:
                entidade = 'STAPE, INC.'
            elif 'EBN*Canva' in memo.replace(" ",""):
                entidade = 'CANVA PTY LTD.'
            elif 'LEARNWORLDS CY LTD' in memo:
                entidade = 'LEARNWORLDS (CY) LTD'
            elif 'WINDSOR.AI' in memo:
                entidade = 'WINDSOR.AI'
            elif 'AOVS SISTEMAS DE INFOR' in memo:
                entidade = 'AOVS SISTEMAS DE INFORMATICA SA'
            elif 'ELEVENLABS.IO' in memo:
                entidade = 'ELEVENLABS.IO'
            elif 'MSFT*' in memo.replace(" ",""):
                entidade = 'MICROSOFT INFORMATICA LTDA'
            elif 'NOTIFICACOES INTELIGEN' in memo:
                entidade = 'KIWIFY PAGAMENTOS, TECNOLOGIA E SERVICOS LTDA'
            elif 'UAZAPI - API WHATSAPP' in memo:
                entidade = 'UAZAPI'
            elif 'TINY ERP' in memo:
                entidade = 'OLIST TINY TECNOLOGIA LTDA'
            elif 'ADOBE' in memo:
                entidade = 'ADOBE SYSTEMS BRASIL LTDA.'
            elif 'EBN*SEMRUSH' in memo.replace(" ",""):
                entidade = 'SEMRUSH'
            elif 'FIGMA' in memo:
                entidade = 'FIGMA MONTHLY RENEWAL'
            elif 'BITLY.COM' in memo:
                entidade = 'BITLY COM'
            elif 'MANYCHAT.COM' in memo:
                entidade = 'MANYCHAT INC.'
            elif 'SUPABASE' in memo:
                entidade = 'SUPABASE'
            elif 'PG*NOTAZZGESTAOFISC' in memo.replace(" ",""):
                entidade = 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA'
            elif 'CLAUDE.AI' in memo:
                entidade = 'CLAUDE.AI'
            elif 'STACKBLITZ' in memo:
                entidade = 'STACKBLITZ, INC'
            elif 'CALENDLY' in memo:
                entidade = 'CALENDLY LLC'    
            elif 'VERCEL INC.' in memo:
                entidade = 'VERCEL INC.'
            elif 'ENVATO' in memo:
                entidade = 'EVANATO ELEMENTES PTY LTD'
            elif 'TOPINVEST ED*TOP INVES' in memo:
                entidade = 'TOPINVEST EDUCACAO FINANCEIRA LTDA'
            elif 'OPENAI' in memo:
                entidade = 'OPENAI,LLC'
            elif 'Amazon AWS Servicos Br' in memo:
                entidade = 'AMAZON AWS SERVICOS BRASIL LTDA'
            elif 'PG*BRDIDTELEFONIA' in memo.replace(" ",""):
                entidade = 'BR TECH TECNOLOGIA EM SISTEMAS LTDA'
            elif 'LOVABLE' in memo:
                entidade = 'LOVABLE'
            elif 'USERBACK*' in memo:
                entidade = 'USERBACK.IO'
            elif 'WEBFLOW.COM' in memo:
                entidade = 'WEBFLOW INC.'
            elif 'BIGSPY' in memo:
                entidade = 'BIGSPY'
            elif 'PADDLE.NET * N8N CLOUD1' in memo:
                entidade = 'CLOUD1 SERVICOS DE INFORMATICA LTDA.'
            elif 'CLICKUP' in memo:
                entidade = 'ClickUp - Mango Technologies, Inc.'
            elif linha [2] in ['Google GSUITE_wtf.mais', 'DL *GOOGLE GSUITEasupe' , 'DL *GOOGLE Google One']:
                entidade = 'GOOGLE - GSUITE'
            elif 'MONGODBCLOUD PAULO' in memo:
                entidade = 'MONGODB SERVICOS DE SOFTWARE NO BRASIL LTDA.'
            elif 'GURU-DISCIPULO PLUS 3' in memo:
                entidade = 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA'
            elif 'DL*GOOGLEAmazon' in memo.replace(" ",""):
                entidade = 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.'
            elif 'MANUS AI' in memo:
                entidade = 'MANUS AI'
            elif 'TWILIO SENDGRID' in memo:
                entidade = 'TWILIO EXPANSION LLC'
            elif 'TURBOSCRIBE' in memo:
                entidade = 'TURBOSCRIBE'
            elif 'OPUS CLIP' in memo:
                entidade = 'OPUS CLIP'
            elif 'RECLAMEAQUI' in memo:
                entidade = 'OBVIO BRASIL SOFTWARE E SERVICOS S.A.'
            elif 'URUAQUE' in memo:
                entidade = 'URUAQUE GOIANIA LTDA'
            elif 'Uber UBER *TRIP HELP.U' in memo:
                entidade = 'UBER DO BRASIL TECNOLOGIA LTDA.'
            elif 'LATAM AIR' in memo:
                entidade = 'LATAM AIRLINES BRASIL'
            elif 'MERCADOLIVRE*' in memo:
                entidade = 'MERCADOLIVRE.COM'
            elif 'OPUS CLIP' in memo:
                entidade = 'OPUS CLIP'
            elif 'DECOLAR' in memo:
                entidade = 'DECOLAR. COM LTDA'
            elif 'IOF - COMPRA INTERNACIONAL' in memo:
                contaDespesa = '3.6.1.01.005 IOF'
           

            ###---------------------------------------------------------------------------------###
            ###--------------------------- COMPRAS RECORRENTES ---------------------------------###
            ###---------------------------------------------------------------------------------###

# Aqui é onde ocorre a seleção para preencher, na planilha de importação, algumas abas de acordo com as assinaturas/compras/pagamentos que são feitos recorrentemente nos cartões, 
# que, assim sendo, já contêm as abas de Entidade, Departamento, Centro de Custo e Classe já bem definidos.

            ### ------------------------- COMPRAS DNS -------------------------------- ###
            # FACEBOOK
            if 'FACEBK' in memo and '2289' in card: 
                importa(empresa, conta, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'PUBLICIDADE', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'AUVP ESCOLA', venc, id, emp,loc)
              
            elif 'FACEBK' in memo and '4836' in card:
                importa(empresa, conta, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'PUBLICIDADE', '3.4.1.03.001 PROPAGANDA E PUBLICIDADE - TRAFEGO', 'AUVP ANALÍTICA', venc, id, emp,loc)  
            # Imagify
            elif 'WP MEDIA - IMAGIFY' in memo and '8389' in card and int(dia) == 9 and linha[3] == '5.99':
                importa(empresa, conta, 'IMAGIFY', 'PUBLICIDADE', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # Openai dia 09 e 23 DNS - alyf
            elif 'OPENAI *CHATGPT SUBSCR' in memo and '3559' in card and (int(dia) == 9 or int(dia) == 23) and linha[3] == '20.0':
                importa(empresa, conta, 'OPENAI,LLC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            #Amazon Prime
            elif 'AmazonPrimeBR' in memo and '8389' in card and int(dia) == 10 and nvalor == '19,9':
                importa(empresa, conta, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'CONTROLADORIA E FINANÇAS', '3.5.1.05.023 OUTRAS DESPESAS ADMINISTRATIVAS', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # Fundamentei
            elif 'FUNDAMENTEI.COM' in memo and '3559' in card and int(dia) == 10 and nvalor == '49,0':
                importa(empresa, conta, 'FUNDAMENTEI SERVICOS DE INFORMACAO LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # Brapi Dev 
            elif 'BRAPI.DEV' in memo and '2289' in card and int(dia) == 10 and nvalor == '49,99':
                importa(empresa, conta, 'BRAPI ASL TECNOLOGIA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # STAPE INC
            elif 'STAPE, INC' in memo and '2289' in card and (int(dia) == 11 or int(dia) == 19) and linha[3] == '20.0':
                importa(empresa, conta, 'STAPE, INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # EBN Canva
            elif 'EBN*Canva' in memo.replace(" ","") and '8409' in card and int(dia) == 13 and nvalor == '174,5':
                importa(empresa, conta, 'CANVA PTY LTD.', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # LEARNWORLDS CY LTD dia 13, 22 e 29
            elif 'LEARNWORLDS CY LTD' in memo and '2289' in card and (int(dia) == 13 or int(dia) == 22 or int(dia) == 29):
                importa(empresa, conta, 'LEARNWORLDS (CY) LTD', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # WINDSOR.AI
            elif 'WINDSOR.AI' in memo and '2289' in card and int(dia) == 14  and linha[3] == '299.0':
                importa(empresa, conta, 'WINDSOR.AI', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # AOVS
            elif 'AOVS SISTEMAS DE INFOR' in memo and '2289' in card and int(dia) == 16 and nvalor == '530,0':
                importa(empresa, conta, 'AOVS SISTEMAS DE INFORMATICA SA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # ELEVENLABS
            elif 'ELEVENLABS.IO' in memo and '3559' in card and int(dia) == 17 and linha[3] == '5.0':
                importa(empresa, conta, 'ELEVENLABS.IO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # MICROSOFT
            elif 'MSFT*' in memo.replace(" ","") and '2289' in card and int(dia) == 18:
                importa(empresa, conta, 'MICROSOFT INFORMATICA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # KIWIFY NOTIFICA
            elif 'NOTIFICACOES INTELIGEN' in memo and '2236' in card and int(dia) == 18:
                importa(empresa, conta, 'KIWIFY PAGAMENTOS, TECNOLOGIA E SERVICOS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc) 
            # UAZAPI
            elif 'UAZAPI - API WHATSAPP' in memo and '3559' in card and int(dia) == 19 and nvalor == '29,0' :
                importa(empresa, conta, 'UAZAPI', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # TINY
            elif 'TINY ERP' in memo and '8389' in card and int(dia) == 20 and nvalor == '135,89':
                importa(empresa, conta, 'OLIST TINY TECNOLOGIA LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # ADOBE DIA 20 e 06   
            elif 'ADOBE' in memo and '2289' in card and (int(dia) == 20 or int(dia) == 6) and nvalor == '275,0':
                importa(empresa, conta, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # SEMRUSH   
            elif 'EBN*SEMRUSH' in memo.replace(" ","") and '8389' in card and int(dia) == 21:
                importa(empresa, conta, 'SEMRUSH', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # FIGMA
            elif 'FIGMA' in memo and '3559' in card and int(dia) == 21:
                importa(empresa, conta, 'FIGMA MONTHLY RENEWAL', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # BITLY
            elif 'BITLY.COM' in memo and '2289' in card and int(dia) == 22 and linha[3] == '35.0':
                importa(empresa, conta, 'BITLY COM', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # MANYCHAT    
            elif 'MANYCHAT.COM' in memo and '5518' in card and int(dia) == 25:
                importa(empresa, conta, 'MANYCHAT INC.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # SUPABASE    
            elif 'SUPABASE' in memo and '3559' in card and int(dia) == 26:
                importa(empresa, conta, 'SUPABASE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP SEMPRE', venc, id, emp,loc)
            # NOTAZZ  
            elif 'PG*NOTAZZGESTAOFISC' in memo.replace(" ","") and '2289' in card and int(dia) == 26 and nvalor == '367,90':
                importa(empresa, conta, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # CLAUDE  
            elif 'CLAUDE.AI' in memo and '3559' in card and int(dia) == 27 and nvalor == '110,0':
                importa(empresa, conta, 'CLAUDE.AI', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # BOLT STACKBLITZ  
            elif 'BOLT (BY STACKBLITZ)' in memo and '3559' in card and int(dia) == 27 and linha[3] == '20.0':
                importa(empresa, conta, 'STACKBLITZ, INC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # CALENDY
            elif 'CALENDLY' in memo and '7886' in card and int(dia) == 28 and linha[3] == '20.0':
                importa(empresa, conta, 'CALENDLY LLC', 'INSIDE SALES', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # VERCEL
            elif 'VERCEL INC.' in memo and '3559' in card and int(dia) == 29 and linha[3] == '20.0':
                importa(empresa, conta, 'VERCEL INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # ENVATO
            elif 'ENVATO' in memo and '3559' in card and int(dia) == 29 and linha[3] == '33.0':
                importa(empresa, conta, 'EVANATO ELEMENTES PTY LTD', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # TOPINVEST   
            elif 'TOPINVEST ED*TOP INVES' in memo and '3559' in card and int(dia) == 1 and nvalor == '99,9':
                importa(empresa, conta, 'TOPINVEST EDUCACAO FINANCEIRA LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.04.011 LIVROS CURSOS E TREINAMENTOS - CUSTO', 'AUVP PRO', venc, id, emp,loc)
            # OPENAI 
            elif 'OPENAI' in memo and '3559' in card and int(dia) == 2:
                importa(empresa, conta, 'OPENAI,LLC', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP SEMPRE', venc, id, emp,loc)
            # AMAZON AWS
            elif 'Amazon AWS Servicos Br' in memo and '3559' in card and int(dia) == 2:
                importa(empresa, conta, 'AMAZON AWS SERVICOS BRASIL LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.002 SERVIÇO DE HOSPEDAGEM E NUVEM', 'OPERAÇÃO & PRODUÇÃO: DNS', venc, id, emp,loc)
            # BR DID dia 03 e 17 Maurício atendimento
            elif 'PG*BRDIDTELEFONIA' in memo.replace(" ","") and '2889' in card and ((int(dia) == 3 and nvalor == '11,9') or (int(dia) == 17 and nvalor == '29,8')):
                importa(empresa, conta, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'COMERCIAL: DNS', venc, id, emp,loc)
            # BR DID dia 16 e 19 alyf
            elif 'PG*BRDIDTELEFONIA' in memo.replace(" ","") and '3559' in card and (int(dia) == 16 or int(dia) == 19)and nvalor == '23,9':
                importa(empresa, conta, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # CHATGPT VITÃO
            elif 'OPENAI*CHATGPT SUBSCR' in memo.replace(" ","") and '2889' in card and int(dia) == 7 and linha[3] == '20.0':
                importa(empresa, conta, 'OPENAI,LLC', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'CANAL INVESTIDOR SARDINHA', venc, id, emp,loc)
            # LOVABLE
            elif 'LOVABLE' in memo and '3559' in card and int(dia) == 7 and linha[3] == '25.0':
                importa(empresa, conta, 'LOVABLE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            # USERBACK
            elif 'USERBACK*' in memo and '3559' in card and int(dia) == 7 and linha[3] == '68.0':
                importa(empresa, conta, 'USERBACK.IO', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'AUVP ANALíTICA', venc, id, emp,loc)
            #RAILWAY
            elif 'RAILWAY' in memo and '3559' in card and int(dia) == 4 and linha[3] == '20.0':
                importa(empresa, conta, 'RAILWAY CORP.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: DNS', venc, id, emp,loc)
            
            ### ------------------------- COMPRAS THE BRAIN -------------------------------- ###
            
            # WEBFLOW
            elif 'WEBFLOW.COM' in memo and '7003' in card:
                importa(empresa, conta, 'WEBFLOW INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # BIGPSY
            elif 'BIGSPY' in memo and '5206' in card and int(dia) == 10 and linha[3] == '99.0':
                importa(empresa, conta, 'BIGSPY', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            # PADDLE NET CLOUD
            elif 'PADDLE.NET*N8NCLOUD1' in memo.replace(" ","") and '7003' in card and int(dia) == 12 and nvalor == '360,0':
                importa(empresa, conta, 'CLOUD1 SERVICOS DE INFORMATICA LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.002 SERVIÇO DE HOSPEDAGEM E NUVEM', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # CLICKUP DIA 13
            elif 'CLICKUP' in memo and '7003' in card and int(dia) == 13:
                importa(empresa, conta, 'ClickUp - Mango Technologies, Inc.', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', venc, id, emp,loc)
            # ZOOM
            elif 'ZOOM.COM 888-799-9666' in memo and '6187' in card and int(dia) == 17:
                importa(empresa, conta, 'ClickUp - Mango Technologies, Inc.', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', venc, id, emp,loc)
            # TINY
            elif 'TINY ERP' in memo and '5206' in card and int(dia) == 20 and nvalor == "149,9":
                importa(empresa, conta, 'OLIST TINY TECNOLOGIA LTDA', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # ADOBE DIA 26 PRODUTO
            elif 'ADOBE' in memo and '8537' in card and int(dia) == 26 and nvalor == '139,0':
                importa(empresa, conta, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # NOTAZZ
            elif 'PG*NOTAZZGESTAOFISC' in memo.replace(" ","") and '7003' in card and int(dia) == 27 and nvalor == "910,9":
                importa(empresa, conta, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: THE BRAIN', venc, id, emp,loc)
            # GOOGLE WTF
            elif 'Google GSUITE_wtf.mais' in memo and '7003' in card and int(dia) == 1:
                importa(empresa, conta, 'GOOGLE - GSUITE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # MONGODBCLOUD
            elif 'MONGODBCLOUD PAULO' in memo and '7003' in card and int(dia) == 2:
                importa(empresa, conta, 'MONGODB SERVICOS DE SOFTWARE NO BRASIL LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # MANYCHAT
            elif 'MANYCHAT.COM' in memo and '5206' in card and int(dia) == 3 and linha[3] == '65.0':
                importa(empresa, conta, 'MANYCHAT INC.', 'PUBLICIDADE', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'O SUPERPODER', venc, id, emp,loc)
            # GURU DIA 4 INSIDE SALES
            elif 'GURU-DISCIPULO PLUS 3' in memo and '2599' in card and int(dia) == 4 and linha[3] == '187.35':
                importa(empresa, conta, 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA', 'INSIDE SALES', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'O SUPERPODER', venc, id, emp,loc)
            # CANVA
            elif 'EBN*Canva' in memo.replace(" ","") and '7003' in card and int(dia) == 5 and nvalor == '44,99':
                importa(empresa, conta, 'CANVA PTY LTD.', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # GOOGLE AMAZON
            elif 'DL*GOOGLEAmazon' in memo.replace(" ","") and '5206' in card and int(dia) == 6 and nvalor == '19,9':
                importa(empresa, conta, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            # ADOBE DIA 6 AUDIOVISUAL
            elif 'ADOBE' in memo and '5206' in card and int(dia) == 6 and nvalor == '275,0':
                importa(empresa, conta, 'ADOBE SYSTEMS BRASIL LTDA.', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: THE BRAIN', venc, id, emp,loc)
            
            
            ### ------------------------- COMPRAS SUPERNOVA -------------------------------- ###

            # GOOGLE GSUITE SUPERNOVA
            elif 'DL*GOOGLE GSUITEasupe' in memo.replace(" ","") and '5393' in card and int(dia) == 6:
                importa(empresa, conta, 'GOOGLE - GSUITE', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # LOVABLE
            elif 'LOVABLE' in memo and '0799' in card and (int(dia) == 13 or int(dia) == 1):
                importa(empresa, conta, 'LOVABLE', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)
            #  MANUS
            elif 'MANUS AI' in memo and '0799' in card and int(dia) == 15 and linha[3] == '199.0':
                importa(empresa, conta, 'MANUS AI', 'PRODUTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # OPENAI CONTROLADORIA DIA 17
            elif 'OPENAI*CHATGPTSUBSCR' in memo.replace(" ","") and '7666' in card and int(dia) == 17 and linha[3] == '20.0':
                importa(empresa, conta, 'OPENAI,LLC', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # OPENAIO PRODUTO DIA 3 E 19
            elif 'OPENAI*CHATGPTSUBSCR' in memo.replace(" ","") and '6614' in card and (int(dia) == 19 or int(dia) == 3) and linha[3] == '20.0':
                importa(empresa, conta, 'OPENAI,LLC', 'PRODUTO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # GOOGLE ONE AUDIOVISUAL
            elif 'DL*GOOGLEGoogleOne' in memo.replace(" ","") and '3341' in card and int(dia) == 25 and nvalor == '609,0':
                importa(empresa, conta, 'GOOGLE - GSUITE', 'PRODUÇÃO AUDIOVISUAL', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # VERCEL
            elif 'VERCEL INC.' in memo and '5393' in card and int(dia) == 28 and linha[3] == '20.0':
                importa(empresa, conta, 'VERCEL INC.', 'TECNOLOGIA E DESENVOLVIMENTO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # WORKSPACE FACE
            elif 'FACEBK' in memo and '6614' in card and int(dia) == 1:
                importa(empresa, conta, 'FACEBOOK SERVICOS ONLINE DO BRASIL LTDA', 'CAPITAL HUMANO', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'ADMINISTRATIVO: SUPERNOVA', venc, id, emp,loc)
            # OPENAI JURIDICO CAMILA EMILY
            elif 'OPENAI*CHATGPTSUBSCR' in memo.replace(" ","") and '6614' in card and int(dia) == 4 and linha[3] == '20.0':
                importa(empresa, conta, 'OPENAI,LLC', 'JURÍDICO', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'OPERAÇÃO & PRODUÇÃO: SUPERNOVA', venc, id, emp,loc)
            # SENDGRID 
            elif 'TWILIO SENDGRID' in memo and '6614' in card and int(dia) == 3 and linha[3] == '89.95':
                importa(empresa, conta, 'TWILIO EXPANSION LLC', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP ESCOLA', venc, id, emp,loc)   
            
            ### ------------------------- COMPRAS AUVP -------------------------------- ###

             # OPENAI KAIQUE
            elif 'OPENAI*CHATGPTSUBSCR' in memo.replace(" ","") and '5550' in card and int(dia) == 14 and linha[3] == '20.0':
                importa(empresa, conta, 'OPENAI,LLC', 'DIRETORIA', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'ADMINISTRATIVO: AUVP CONSULTORIA', venc, id, emp,loc)            
            # NOTAZZ
            elif 'PG*NOTAZZGESTAOFISC' in memo.replace(" ","") and '9342' in card and int(dia) == 15 and nvalor == '728,9':
                importa(empresa, conta, 'NOTAZZ GESTAO FISCAL E LOGISTICA LTDA', 'CONTROLADORIA E FINANÇAS', '3.5.1.04.002 LICENÇAS E USO DE SOFTWARES', 'AUVP CONSULTORIA', venc, id, emp,loc)
            # INVESTING
            elif 'INVESTING.COM' in memo and '9033' in card and int(dia) == 22 and nvalor == '99,0':
                importa(empresa, conta, 'INVESTING.COM', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.003 SERVIÇO DE ACESSO A CONTEÚDO', 'AUVP CONSULTORIA', venc, id, emp,loc)
            # BR DID
            elif 'PG*BRDIDTELEFONIA' in memo.replace(" ","") and '3428' in card and int(dia) == 25 and nvalor == '23,9':
                importa(empresa, conta, 'BR TECH TECNOLOGIA EM SISTEMAS LTDA', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)
            # TURBOSCRIBE
            elif 'TURBOSCRIBE' in memo and '3428' in card and int(dia) == 29 and linha[3] == '20.0':
                importa(empresa, conta, 'TURBOSCRIBE', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)
            # OPUSCLIP
            elif 'OPUS CLIP' in memo and '9342' in card and int(dia) == 6 and linha[3] == '19.0':
                importa(empresa, conta, 'OPUS CLIP', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)
            # GURU APRENDIZ
            elif 'GURU-APRENDIZ-II' in memo and '9342' in card and int(dia) == 9 and linha[3] == '64.87':
                importa(empresa, conta, 'DIGITAL MANAGER GURU - MARGEM INQUESTIONÁVEL SA', 'CONSULTORIA E INVESTIMENTOS', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP CONSULTORIA', venc, id, emp,loc)
            # AMAZON PRIME
            elif 'AmazonPrimeBR' in memo and '3306' in card and int(dia) == 11 and nvalor == '19,9':
                importa(empresa, conta, 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA.', 'CONTROLADORIA E FINANÇAS', '3.5.1.05.023 OUTRAS DESPESAS ADMINISTRATIVAS', 'ADMINISTRATIVO: AUVP CONSULTORIA', venc, id, emp,loc)
            # RECLAME AQUI
            elif 'RECLAMEAQUI' in memo and '3428' in card and int(dia) == 15 and nvalor == '49,9':
                importa(empresa, conta, 'OBVIO BRASIL SOFTWARE E SERVICOS S.A.', 'ATENDIMENTO E CX', '3.4.1.06.001 CUSTO COM MANUTENÇÃO, LICENÇA E USO DE SOFTWARE', 'AUVP BANCO', venc, id, emp,loc)  
            
            else:

                importa(empresa, conta, entidade, departamento, contaDespesa, '', venc, id, emp,loc)
                
df = pd.DataFrame.from_dict(dados)
df.to_excel("teste.xlsx", index = False)