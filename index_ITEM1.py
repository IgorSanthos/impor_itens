import pandas as pd
from pathlib import Path
from tkinter import filedialog, messagebox
from datetime import date

try:
    #filePathIn = Path(r'C:\Users\Igor\Desktop\ACP_entradas_itens_jettax.xlsx')
    filePathIn = filedialog.askopenfilename(filetypes=[("arquivos XLSX", "*.xlsx")])
    dfItens = pd.read_excel(filePathIn, sheet_name="Relatório Detalhado por Produto", usecols = [i for i in range (1, 49)])

# Criando uma nova coluna que é a concatenação das colunas 19 e 20
#Tratando coluna de CST

    dfItens = dfItens.astype(str)
    dfItens['CONCATENACAO'] = dfItens.iloc[:, 6] + dfItens.iloc[:, 7]
    dfItens['CONCATENACAO'] = dfItens['CONCATENACAO'].str.replace('.', '')
    dfItens['CONCATENACAO'] = dfItens['CONCATENACAO'].str.slice(0, 3)

    if (dfItens['CONCATENACAO'] == '0na').any():
        dfItens.loc[dfItens['CONCATENACAO'] == '0na', 'CONCATENACAO'] = ''

#Tratando coluna CSOSN
    dfItens['CSON'] = dfItens.iloc[:, 6] + dfItens.iloc[:, 8]
    dfItens['CSON'] = dfItens['CSON'].str.replace('.', '')
    dfItens['CSON'] = dfItens['CSON'].str.slice(0, 4)
    
    if dfItens['CSON'].str.contains('na', case=False, na=False).any():
        dfItens.loc[dfItens['CSON'].str.contains('na', case=False, na=False), 'CSON'] = ''

# TRATANDO CST IPI
    if dfItens['CST IPI'].str.contains('na', case=False, na=False).any():
        dfItens.loc[dfItens['CST IPI'].str.contains('na', case=False, na=False), 'CST IPI'] = ''     

# TRATANDO CST cest
    if dfItens['Cest'].str.contains('na', case=False, na=False).any():
        dfItens.loc[dfItens['Cest'].str.contains('na', case=False, na=False), 'Cest'] = ''

    dfItens['Descrição'] = dfItens['Descrição'].str.replace(';', '')
    


    if (dfItens['EAN'] == 'nan').any():
        dfItens.loc[dfItens['EAN'] == 'nan', 'EAN'] = ''


#OUTRO COMENTARIO        

    dfItens.rename(columns={'Cód. prod.': 'COD_ITEM',	'Descrição': 'DESCR_ITEM',	
                            'Categoria': 'TIPO_ITEM',	'NCM': 'COD_NCM',	'EAN': 'COD_BARRA',	
                            'Cest': 'COD_CEST',	'CONCATENACAO': 'CST_ICMS',	
                            'CSON': 'CSOSN',	'Unidade': 'UNID_INV',	'Alíquota de ICMS': 'ALIQ_ICMS',
                            'CST IPI': 'CST_IPI_ENTRADA',	'Alíquota de IPI': 'ALIQ_IPI',	
                            'Alíquota de PIS': 'ALIQ_PIS',	'CST COFINS': 'CST_PIS_COFINS_ENTRADA',	
                            'Alíquota de COFINS': 'ALIQ_COFINS'}, inplace=True)
    
    infoColuna = ['COD_ANT_ITEM','EX_IPI','COD_LST','COD_SERV_BLOCO_P',
                   'COD_SEFAZ','CST_IPI_SAIDA','CST_PIS_COFINS_SAIDA',
                   'NAT_REC_PIS_COFINS','APURACAO_PIS_COFINS','COD_CENTRO_DE_CUSTOS',
                   'DATA_INC_ALTERACAO_CUSTOS','NOME_CENTRO_CUSTOS','COD_PLANO_CONTAS_REF',
                   'OBSERVACAO', 'NOME_CONTA']
    for itensLista in infoColuna:
        dfItens[itensLista] = ''

    dfItens = dfItens.assign(**{'COD_GRUPO':'1'})
    dfItens = dfItens.assign(**{'DESC_GRUPO':'GERAL'})
    dfItens = dfItens.assign(**{'PER_RED_BC_ICMS':'0.00'})
    dfItens = dfItens.assign(**{'BC_ICMS_ST':''})
    dfItens = dfItens.assign(**{'CC':'1.1.2.01.00002'})
    dfItens = dfItens.assign(**{'DATA_INC_ALTERACAO':date.today()})
    dfItens = dfItens.assign(**{'COD_NAT':'1'})
    dfItens = dfItens.assign(**{'IND_TIPO_CONTA':'A'})
    dfItens = dfItens.assign(**{'NIVEL':'5'})
    dfItens = dfItens.assign(**{'CNPJ_ESTABELECIMENTO':'  .   .   /    -'})
    dfItens = dfItens.assign(**{'REV_STPISCOFINS':'NÃO'})


      #ORDENAR AS COLUNAS
    ordemColunas = ["COD_ITEM", "DESCR_ITEM", "COD_BARRA", "COD_ANT_ITEM","UNID_INV", "TIPO_ITEM",
                    "COD_NCM", "EX_IPI", "COD_LST", "COD_SERV_BLOCO_P", "ALIQ_ICMS", "COD_GRUPO",
                    "DESC_GRUPO", "COD_SEFAZ", "CSOSN", "CST_ICMS", "PER_RED_BC_ICMS", "BC_ICMS_ST",
                    "CST_IPI_ENTRADA", "CST_IPI_SAIDA", "ALIQ_IPI", "CST_PIS_COFINS_SAIDA",
                    "CST_PIS_COFINS_ENTRADA", "NAT_REC_PIS_COFINS", "APURACAO_PIS_COFINS", "ALIQ_PIS",
                    "ALIQ_COFINS", "CC", "DATA_INC_ALTERACAO", "COD_NAT", "IND_TIPO_CONTA", "NIVEL",
                    "NOME_CONTA", "COD_CENTRO_DE_CUSTOS", "DATA_INC_ALTERACAO_CUSTOS", "NOME_CENTRO_CUSTOS",
                    "COD_PLANO_CONTAS_REF", "CNPJ_ESTABELECIMENTO", "OBSERVACAO", "COD_CEST", "REV_STPISCOFINS"
                    ]    
    dfItens = dfItens[ordemColunas]



    #Tratando coluna de CST
    if (dfItens['TIPO_ITEM'] == 'Uso e Consumo').any():
        dfItens.loc[dfItens['TIPO_ITEM'] == 'Uso e Consumo', 'TIPO_ITEM'] = '07'

    if (dfItens['TIPO_ITEM'] == 'Remessa').any():
        dfItens.loc[dfItens['TIPO_ITEM'] == 'Remessa', 'TIPO_ITEM'] = '09'

    if (dfItens['TIPO_ITEM'] == 'nan').any():
        dfItens.loc[dfItens['TIPO_ITEM'] == 'nan', 'TIPO_ITEM'] = '00'

    if (dfItens['TIPO_ITEM'] == 'Industrialização').any():
        dfItens.loc[dfItens['TIPO_ITEM'] == 'Industrialização', 'TIPO_ITEM'] = '01'

    if  dfItens['TIPO_ITEM'].isnull().values.any():
        dfItens['TIPO_ITEM'].fillna('00', inplace=True)
        

    if messagebox.askyesno("SAlvar Arquivo", "Deseja salvar o arquivo"):
          file_path_out = filedialog.asksaveasfilename(defaultextension = ".csv", filetypes=[("Arquivos CSV", "*.csv")])
          dfItens.to_csv(file_path_out, sep=';', index=False)
          messagebox.showinfo("Sucessp", "Arquivo salvo com sucesso !")
    else:
          messagebox.showinfo("Cancelar", "Operacao cancelada pelo usuario !")
    #print (dfItens['CSOSN'].head(5))   
except Exception as error:
    print("MENSAGEM DE ERRO:", error)