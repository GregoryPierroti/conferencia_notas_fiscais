#CONFERENCIA NECESSARIA PARA VERIFICAR SE AS NOTAS GERADAS NO MES NO sistema BATE COM AS 
#NOTAS GERADAS NO SITE DA UNIDADE EM QUESTÃO  - sistema(EXCEL) SITE(XML)
import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import os
from datetime import datetime
#pd.set_option('max_columns', None)
#pd.set_option("max_rows", None)

mes = 1
ano = 2022

#lendo a base de dados de notas faturadas no sistema para conferencia, filtros utilizados para filtrar a empresa correta,
# ano atual, o mes atual, e resetando o index da tabela para manipular a tabela mais a frente de forma mais facilitada.
sistema =pd.read_excel('C:/Users/Gregory Toledo/conferencia notas fiscais/RELATORIO SISTEMA/planilha do sistema.xlsx')
sistema = sistema[sistema['UNIDADE FATURAMENTO'] ==('CONSULTORIA DE LOGISTICAS LTDA' )]
sistema = sistema[sistema['ANO'] == (ano)]
sistema = sistema[sistema['MES'] == (mes)]
sistema = sistema.reset_index(drop=True)

#Acessando o arquivo XML do mes passado retirado do site de nova lima contendo todas as notas fiscais
caminho = 'C:/Users/Gregory Toledo/conferencia notas fiscais/XML'
filename = os.listdir(caminho)[0]
fullname = os.path.join(caminho, filename)
tree = ET.parse(fullname)

#Criando as listas para receber as informações  do XML
nfe = []
nf = []
data_emissao = []
serv = []
valor_bruto = []
valor_liquido = []
PIS = []
COFINS = []
CSLL = []
IR = []
ISS = []
nome_cliente = []
razao = []

#Coletando as informações e adicionando as listas
for elm in tree.findall(".//Numero"):
    nfe.append(elm.text)
for elm in tree.findall(".//Competencia"):
    data_emissao.append(elm.text)    
for elm in tree.findall(".//ItemListaServico"):
    serv.append(elm.text)   
for elm in tree.findall(".//ValorServicos"):
    valor_bruto.append(elm.text)   
for elm in tree.findall(".//ValorLiquidoNfse"):
    valor_liquido.append(elm.text)   
for elm in tree.findall(".//ValorPis"):
    PIS.append(elm.text) 
for elm in tree.findall(".//ValorCofins"):
    COFINS.append(elm.text) 
for elm in tree.findall(".//ValorCsll"):
    CSLL.append(elm.text) 
for elm in tree.findall(".//ValorIr"):
    IR.append(elm.text) 
for elm in tree.findall(".//ValorIssRetido"):
    ISS.append(elm.text) 
for elm in tree.findall(".//RazaoSocial"):
    nome_cliente.append(elm.text)

# a tag<numero> é utilizada para numero da nota e numero de endereços,
#para que seja colhido somente o numero das notas fiscais é preciso que usemos o while abaixo
#que ignora os numeros de endereços contidos na nota, pulando de 4 em 4 indices e agregando a uma
#segunda lista filtrada com somente numeros de NF's
cont = 0
while (cont < len(nfe)):
    nf.append(nfe[cont])
    cont = cont + 4

#a tag<RazaoSocial> é utilizada para os clientes e para o emissor da nf(no caso a inter), utilizamos
# o while abaixo ignora a razão social da inter e colhe a do cliente para a lista filtrada com somente
# clientes

cont = 1
while (cont < len(nome_cliente)):
    razao.append(nome_cliente[cont])
    cont = cont + 2


#Dataframes utilizados para organizarmos todas as informações
df = pd.DataFrame(columns=['NF','DATA EMISSÃO','COMPETENCIA','ITEM SERVIÇO','VALOR BRUTO SISTEMA', 'ISS SISTEMA','PIS SISTEMA','COFINS SISTEMA', 'CSLL SISTEMA','IR SISTEMA','VALOR LIQUIDO SISTEMA','VALOR BRUTO SITE', 'ISS SITE','PIS SITE', 'COFINS SITE','CSLL SITE','IR SITE','VALOR LIQUIDO SITE','DIFERENÇA','CLIENTE SISTEMA','CLIENTE SITE','COMPARAÇÃO NOME'])
df_site = pd.DataFrame(columns=['NF', 'COMPETENCIA', 'ITEM SERVIÇO','VALOR BRUTO SITE','ISS SITE','PIS SITE','COFINS SITE','CSLL SITE','IR SITE','VALOR LIQUIDO SITE','CLIENTE SITE'])

#Etapa muito importante para que juntemos todas as notas geradas no site e no sistema, e removermos as duplicatas
#podem ter notas geradas somente no sistema ou no site, e a conferencia serve para identificar as mesmas e tratar esses erros

#lista com as notas geradas e validas do sistema
nf_sistema = sistema['NR.NOTA DE SERVIÇO'].tolist()
#convertendo as notas do site de string para inteiro
nf = list(map(int,nf))
#juntando todas as notas em uma lista só
nf_def = nf_sistema + nf
#removendo as duplicatas presentes na nova lista e reformando em lista
nf_def = list(set(nf_def))
#inclundo todas as notas no dataframe principal
df['NF'] = nf_def


#df_ste é um dataframe secundario criado para auxiliar o df a colher informações referentes ao site de maneira mais 
#adequada e organizada, agregando todas as listas com a informação do site no dataframe.

df_site['NF'] = nf
df_site['COMPETENCIA'] = data_emissao
df_site['ITEM SERVIÇO'] = serv
df_site['VALOR BRUTO SITE'] = list(map(float,valor_bruto))
df_site['VALOR LIQUIDO SITE'] = list(map(float,valor_liquido))
df_site['ISS SITE'] = list(map(float,ISS))
df_site['PIS SITE'] = list(map(float,PIS))
df_site['COFINS SITE'] = list(map(float,COFINS))
df_site['CSLL SITE'] = list(map(float,CSLL))
df_site['IR SITE'] = list(map(float,IR))
df_site['CLIENTE SITE'] = razao

#inserindo as informações do sistema(excel) no df de acordo com as notas fiscais
for cont4 in range(len(df)):
    for cont5 in range(len(sistema)):
        if df['NF'].loc[cont4] == sistema['NR.NOTA DE SERVIÇO'].loc[cont5] :
            df['DATA EMISSÃO'].loc[cont4] = sistema['DT. EMISSAO'].loc[cont5]
            df['VALOR BRUTO SISTEMA'].loc[cont4] = sistema['VLR.RECEITA'].loc[cont5]
            df['ISS SISTEMA'].loc[cont4] = sistema['VLR.ISS'].loc[cont5]
            df['PIS SISTEMA'].loc[cont4] = sistema['VLR.PIS'].loc[cont5]
            df['COFINS SISTEMA'].loc[cont4] = sistema['VLR.COFINS'].loc[cont5]
            df['CSLL SISTEMA'].loc[cont4] = sistema['VLR.CSLL'].loc[cont5]
            df['IR SISTEMA'].loc[cont4] = sistema['VLR.IRRF'].loc[cont5]
            df['VALOR LIQUIDO SISTEMA'].loc[cont4] = sistema['VLR LIQUIDO'].loc[cont5]
            df['CLIENTE SISTEMA'].loc[cont4] = sistema['CLIENTE'].loc[cont5]


#inserindo as informações do df_site no df de acordo com as notas fiscais
cont4 = 0
cont5 = 0
for cont4 in range(len(df)):
    for cont5 in range(len(df_site)):
        if df['NF'].loc[cont4] == df_site['NF'].loc[cont5] :
            df['COMPETENCIA'].loc[cont4] = df_site['COMPETENCIA'].loc[cont5]
            df['ITEM SERVIÇO'].loc[cont4] = df_site['ITEM SERVIÇO'].loc[cont5]
            df['VALOR BRUTO SITE'].loc[cont4] = df_site['VALOR BRUTO SITE'].loc[cont5]
            df['ISS SITE'].loc[cont4] = df_site['ISS SITE'].loc[cont5]
            df['PIS SITE'].loc[cont4] = df_site['PIS SITE'].loc[cont5]
            df['COFINS SITE'].loc[cont4] = df_site['COFINS SITE'].loc[cont5]
            df['CSLL SITE'].loc[cont4] = df_site['CSLL SITE'].loc[cont5]
            df['IR SITE'].loc[cont4] = df_site['IR SITE'].loc[cont5]
            df['VALOR LIQUIDO SITE'].loc[cont4] = df_site['VALOR LIQUIDO SITE'].loc[cont5]
            df['CLIENTE SITE'].loc[cont4] = df_site['CLIENTE SITE'].loc[cont5]


#substituindo os NaN da coluna em valor numerico '0.00'  
df['ISS SISTEMA'] = df['ISS SISTEMA'].replace(np.nan, 0.00, regex=True)
df['PIS SISTEMA'] = df['PIS SISTEMA'].replace(np.nan, 0.00, regex=True)
df['COFINS SISTEMA'] = df['COFINS SISTEMA'].replace(np.nan, 0.00, regex=True)
df['CSLL SISTEMA'] = df['CSLL SISTEMA'].replace(np.nan, 0.00, regex=True)
df['IR SISTEMA'] = df['IR SISTEMA'].replace(np.nan, 0.00, regex=True)
#Calculando a diferença entre as notas no valor liquido:
#for cont8 in range(df):
    #df['DIFERENÇA'].loc[cont8] = float(df['VALOR LIQUIDO SISTEMA'].loc[cont8]) - float(df['VALOR LIQUIDO SITE'].loc[cont8])
 #   if df['CLIENTE SISTEMA'].loc[cont8] == df['CLIENTE SITE'].loc[cont8]:
  #      df['COMPARAÇÃO NOME'] = 'ok'
   # else:
    #    df['COMPARAÇÃO NOME'] = 'nome divergente'
    

#ordenando por numero da NF
df = df.sort_values(by='NF')

df['VALOR LIQUIDO SISTEMA'] = df['VALOR BRUTO SISTEMA'] - df['ISS SISTEMA'] - df['PIS SISTEMA'] - df['COFINS SISTEMA'] - df['CSLL SISTEMA'] - df['IR SISTEMA']
df['DIFERENÇA'] = df['VALOR LIQUIDO SISTEMA'] - df['VALOR LIQUIDO SITE']

for cont9 in range(len(df)):
    if df['CLIENTE SISTEMA'].loc[cont9] == df['CLIENTE SITE'].loc[cont9]:
        df['COMPARAÇÃO NOME'].loc[cont9] =  'Ok'
    else:
        df['COMPARAÇÃO NOME'].loc[cont9] = 'nome divergente'


df.to_excel('C:/Users/Gregory Toledo/conferencia notas fiscais/planilha final.xlsx',engine = 'openpyxl',encoding='utf-8',index=False)

df
