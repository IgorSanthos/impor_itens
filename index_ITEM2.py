import pandas as pd
from pathlib import Path
from tkinter import filedialog, messagebox

try:
        filePathIn = filedialog.askopenfilename(filetypes=[("arquivos XLSX", "*.xlsx")])
        #filePathIn = Path(r'C:\Users\Igor\Desktop\ACP_entradas_itens_jettax.xlsx.xlsx')
        dfItens = pd.read_excel(filePathIn, sheet_name="Relatório Detalhado por Produto", usecols = [i for i in range (1, 49)])


# colunas repetidas    
        dfItens["CodigoItemFornecedor"] = dfItens['Cód. prod.']
        dfItens['CFOPdestino1'] = ''

# Renomeando colunas       

        dfItens.rename(columns={"CNPJ/CPF . Emitente": "CNPJ", "Cód. prod.": "CodigoItemEmpresa", "Categoria": "TipoItem", "CFOP": "CFOPorigem1",
                                "CST PIS": "PISCofinsCst", "Alíquota de PIS": "PISAliq", "Alíquota de COFINS": "COFINSAliq"}, inplace=True)
    
#   colocando informacoes na linhas de colunas especificas 
        dfItens = dfItens.assign(**{'Variacao1':'0'})
        dfItens = dfItens.assign(**{'PISCofinsCst':'70'})
        dfItens = dfItens.assign(**{'PISAliq':'0.00'})
        dfItens = dfItens.assign(**{'COFINSAliq':'0.00'})
        dfItens = dfItens.assign(**{'COFINSAliqReais':'0.00'})
        dfItens = dfItens.assign(**{"PISAliqReais": '0.00'})




#Trantando DE/PARA de USO E CONSUMO TRIBUTADO
        if (    (((dfItens['CSON'] != 500)
        | ~  (dfItens['CST'].isin([60, 10, 70]))) 
        |    (dfItens['TipoItem'] == 'Uso e Consumo')
        |    (dfItens['CFOPorigem1'] != '5949'))
        ).any():
                
                dfItens.loc[((dfItens['CSON'] != 500)
                        | ~  (dfItens['CST'].isin([60, 10, 70])))
                        |    (dfItens['TipoItem'] == 'Uso e Consumo'),
                        'CFOPdestino1'] = '1556'
        ##FORA DO ESTADO
        if (    ((dfItens['CSON'] != 500)
        | ~  (dfItens['CST'].isin([60, 10, 70]))) 
        |    (dfItens['TipoItem'] == 'Uso e Consumo')
        |    ((dfItens['CFOPorigem1']//1000) == 6)
        ).any():
                
                dfItens.loc[((dfItens['CSON'] != 500)
                        | ~  (dfItens['CST'].isin([60, 10, 70])))
                        |    (dfItens['TipoItem'] == 'Uso e Consumo')
                        |    ((dfItens['CFOPorigem1']//1000) == 6),
                        'CFOPdestino1'] = '2556'
                        
#Trantando DE/PARA de USO E CONSUMO ST
        if (   (dfItens['TipoItem'] == 'Uso e Consumo')
                | ((dfItens['CSON'] == 500)
                | (dfItens['CST'].isin([60, 10, 70]))) 
                | ((dfItens['CFOPorigem1']//1000) == 5)
                ).any():
                
                dfItens.loc[(   (dfItens['TipoItem'] == 'Uso e Consumo')
                                & ((dfItens['CSON'] == 500)
                                | (dfItens['CST'].isin([60, 10, 70]))) 
                                & ((dfItens['CFOPorigem1']//1000) == 5)
                                ),
                        'CFOPdestino1'] = '1407'
            
        ## FORA DO ESTADO
        if (    (dfItens['CSON'] == 500)
                | ((dfItens['CFOPorigem1']//1000) == 6)
                | (dfItens['TipoItem'] == 'Uso e Consumo')
                ).any():
                dfItens.loc[(    (dfItens['CSON'] == 500)
                                | ((dfItens['CFOPorigem1']//1000) == 6)
                                |  (dfItens['TipoItem'] == 'Uso e Consumo')
                                ),
                        'CFOPdestino1'] = '2407'

        #TRATAMENTO DE Matéria Prima
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6102', '2101', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6101', '2101', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6403', '2401', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6404', '2401', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6401', '2401', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6102', '2101', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('6102', '2101', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('5101', '1101', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('5403', '1401', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('5405', '1401', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('5401', '1401', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('5102', '1101', regex=True)
        if (dfItens['TipoItem'] == 'Matéria Prima').any():            
                dfItens.loc[(dfItens['TipoItem'] == 'Matéria Prima'), 'CFOPdestino1'] = dfItens['CFOPdestino1'].astype(str).str.replace('5103', '1101', regex=True)        

#------------------------------------------------------------- DE PARA DE CFOPS GERAIS -------------------------------------------------------------
          
        

# Mapeamento de CFOP de origem para CFOP de destino
        if (dfItens['CFOPorigem1'] == 5405).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5405, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5100).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5100, 'CFOPdestino1'] = 1100
        if (dfItens['CFOPorigem1'] == 5101).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5101, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5102, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5102, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5103).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5103, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5104).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5104, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5106).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5106, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5111).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5111, 'CFOPdestino1'] = 1111
        if (dfItens['CFOPorigem1'] == 5117).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5117, 'CFOPdestino1'] = 1117
        if (dfItens['CFOPorigem1'] == 5120).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5120, 'CFOPdestino1'] = 1120
        if (dfItens['CFOPorigem1'] == 5122).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5122, 'CFOPdestino1'] = 1122
        if (dfItens['CFOPorigem1'] == 5124).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5124, 'CFOPdestino1'] = 1124
        if (dfItens['CFOPorigem1'] == 5125).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5125, 'CFOPdestino1'] = 1125
        if (dfItens['CFOPorigem1'] == 5150).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5150, 'CFOPdestino1'] = 1150
        if (dfItens['CFOPorigem1'] == 5200).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5200, 'CFOPdestino1'] = 1200
        if (dfItens['CFOPorigem1'] == 5202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5202, 'CFOPdestino1'] = 1202
        if (dfItens['CFOPorigem1'] == 5205).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5205, 'CFOPdestino1'] = 1205
        if (dfItens['CFOPorigem1'] == 5208).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5208, 'CFOPdestino1'] = 1208
        if (dfItens['CFOPorigem1'] == 5251).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5251, 'CFOPdestino1'] = 1251
        if (dfItens['CFOPorigem1'] == 5254).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5254, 'CFOPdestino1'] = 1254
        if (dfItens['CFOPorigem1'] == 5300).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5300, 'CFOPdestino1'] = 1300
        if (dfItens['CFOPorigem1'] == 5301).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5301, 'CFOPdestino1'] = 1301
        if (dfItens['CFOPorigem1'] == 5304).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5304, 'CFOPdestino1'] = 1304
        if (dfItens['CFOPorigem1'] == 5306).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5306, 'CFOPdestino1'] = 2306
        if (dfItens['CFOPorigem1'] == 5350).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5350, 'CFOPdestino1'] = 1350
        if (dfItens['CFOPorigem1'] == 5351).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5351, 'CFOPdestino1'] = 1351
        if (dfItens['CFOPorigem1'] == 5354).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5354, 'CFOPdestino1'] = 1354
        if (dfItens['CFOPorigem1'] == 5360).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5360, 'CFOPdestino1'] = 1360
        if (dfItens['CFOPorigem1'] == 5400).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5400, 'CFOPdestino1'] = 1400
        if (dfItens['CFOPorigem1'] == 5403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5403, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5402).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5402, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5403, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5408).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5408, 'CFOPdestino1'] = 1408
        if (dfItens['CFOPorigem1'] == 5409).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5409, 'CFOPdestino1'] = 1409
        if (dfItens['CFOPorigem1'] == 5411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5411, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 5411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5411, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 5412).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5412, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 5413).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5413, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 5414).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5414, 'CFOPdestino1'] = 1414
        if (dfItens['CFOPorigem1'] == 5415).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5415, 'CFOPdestino1'] = 1415
        if (dfItens['CFOPorigem1'] == 5450).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5450, 'CFOPdestino1'] = 1450
        if (dfItens['CFOPorigem1'] == 5451).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5451, 'CFOPdestino1'] = 1451
        if (dfItens['CFOPorigem1'] == 5500).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5500, 'CFOPdestino1'] = 1500
        if (dfItens['CFOPorigem1'] == 5504).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5504, 'CFOPdestino1'] = 1504
        if (dfItens['CFOPorigem1'] == 5505).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5505, 'CFOPdestino1'] = 1505
        if (dfItens['CFOPorigem1'] == 5550).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5550, 'CFOPdestino1'] = 1550
        if (dfItens['CFOPorigem1'] == 5551).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5551, 'CFOPdestino1'] = 1551
        if (dfItens['CFOPorigem1'] == 5554).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5554, 'CFOPdestino1'] = 1554
        if (dfItens['CFOPorigem1'] == 5556).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5556, 'CFOPdestino1'] = 1556
        if (dfItens['CFOPorigem1'] == 5557).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5557, 'CFOPdestino1'] = 1407
        if (dfItens['CFOPorigem1'] == 5600).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5600, 'CFOPdestino1'] = 1600
        if (dfItens['CFOPorigem1'] == 5602).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5602, 'CFOPdestino1'] = 1602
        if (dfItens['CFOPorigem1'] == 5603).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5603, 'CFOPdestino1'] = 1603
        if (dfItens['CFOPorigem1'] == 5605).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5605, 'CFOPdestino1'] = 1605
        if (dfItens['CFOPorigem1'] == 5651).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5651, 'CFOPdestino1'] = 2651
        if (dfItens['CFOPorigem1'] == 5652).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5652, 'CFOPdestino1'] = 1652
        if (dfItens['CFOPorigem1'] == 5655).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5655, 'CFOPdestino1'] = 1652
        if (dfItens['CFOPorigem1'] == 5656).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5656, 'CFOPdestino1'] = 1653
        if (dfItens['CFOPorigem1'] == 5660).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5660, 'CFOPdestino1'] = 1660
        if (dfItens['CFOPorigem1'] == 5662).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5662, 'CFOPdestino1'] = 1662
        if (dfItens['CFOPorigem1'] == 5664).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5664, 'CFOPdestino1'] = 1664
        if (dfItens['CFOPorigem1'] == 5900).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5900, 'CFOPdestino1'] = 1900
        if (dfItens['CFOPorigem1'] == 5901).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5901, 'CFOPdestino1'] = 1901
        if (dfItens['CFOPorigem1'] == 5902).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5902, 'CFOPdestino1'] = 1902
        if (dfItens['CFOPorigem1'] == 5903).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5903, 'CFOPdestino1'] = 1903
        if (dfItens['CFOPorigem1'] == 5907).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5907, 'CFOPdestino1'] = 1907
        if (dfItens['CFOPorigem1'] == 5909).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5909, 'CFOPdestino1'] = 1909
        if (dfItens['CFOPorigem1'] == 5910).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5910, 'CFOPdestino1'] = 1910
        if (dfItens['CFOPorigem1'] == 5911).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5911, 'CFOPdestino1'] = 1911
        if (dfItens['CFOPorigem1'] == 5912).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5912, 'CFOPdestino1'] = 1912
        if (dfItens['CFOPorigem1'] == 5915).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5915, 'CFOPdestino1'] = 1915
        if (dfItens['CFOPorigem1'] == 5916).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5916, 'CFOPdestino1'] = 1916
        if (dfItens['CFOPorigem1'] == 5920).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5920, 'CFOPdestino1'] = 1920
        if (dfItens['CFOPorigem1'] == 5922).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5922, 'CFOPdestino1'] = 1922
        if (dfItens['CFOPorigem1'] == 5924).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5924, 'CFOPdestino1'] = 1924
        if (dfItens['CFOPorigem1'] == 5925).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5925, 'CFOPdestino1'] = 1925
        if (dfItens['CFOPorigem1'] == 5926).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5926, 'CFOPdestino1'] = 1926
        if (dfItens['CFOPorigem1'] == 5929).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5929, 'CFOPdestino1'] = 1556
        if (dfItens['CFOPorigem1'] == 5931).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5931, 'CFOPdestino1'] = 1931
        if (dfItens['CFOPorigem1'] == 5933).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5933, 'CFOPdestino1'] = 1933
        if (dfItens['CFOPorigem1'] == 5949).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5949, 'CFOPdestino1'] = 1949
        if (dfItens['CFOPorigem1'] == 6100).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6100, 'CFOPdestino1'] = 1100
        if (dfItens['CFOPorigem1'] == 6102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6102, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6102, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6106).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6106, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6108).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6108, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6111).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6111, 'CFOPdestino1'] = 1111
        if (dfItens['CFOPorigem1'] == 6117).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6117, 'CFOPdestino1'] = 2117
        if (dfItens['CFOPorigem1'] == 6120).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6120, 'CFOPdestino1'] = 1120
        if (dfItens['CFOPorigem1'] == 6122).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6122, 'CFOPdestino1'] = 1122
        if (dfItens['CFOPorigem1'] == 6124).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6124, 'CFOPdestino1'] = 2124
        if (dfItens['CFOPorigem1'] == 6125).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6125, 'CFOPdestino1'] = 1125
        if (dfItens['CFOPorigem1'] == 6150).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6150, 'CFOPdestino1'] = 1150
        if (dfItens['CFOPorigem1'] == 6200).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6200, 'CFOPdestino1'] = 1200
        if (dfItens['CFOPorigem1'] == 6202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6202, 'CFOPdestino1'] = 2202
        if (dfItens['CFOPorigem1'] == 6202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6202, 'CFOPdestino1'] = 2202
        if (dfItens['CFOPorigem1'] == 6205).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6205, 'CFOPdestino1'] = 1205
        if (dfItens['CFOPorigem1'] == 6208).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6208, 'CFOPdestino1'] = 1208
        if (dfItens['CFOPorigem1'] == 6250).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6250, 'CFOPdestino1'] = 1250
        if (dfItens['CFOPorigem1'] == 6251).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6251, 'CFOPdestino1'] = 1251
        if (dfItens['CFOPorigem1'] == 6254).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6254, 'CFOPdestino1'] = 1254
        if (dfItens['CFOPorigem1'] == 6300).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6300, 'CFOPdestino1'] = 1300
        if (dfItens['CFOPorigem1'] == 6301).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6301, 'CFOPdestino1'] = 1301
        if (dfItens['CFOPorigem1'] == 6304).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6304, 'CFOPdestino1'] = 1304
        if (dfItens['CFOPorigem1'] == 6306).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6306, 'CFOPdestino1'] = 2306
        if (dfItens['CFOPorigem1'] == 6350).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6350, 'CFOPdestino1'] = 1350
        if (dfItens['CFOPorigem1'] == 6351).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6351, 'CFOPdestino1'] = 1351
        if (dfItens['CFOPorigem1'] == 6354).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6354, 'CFOPdestino1'] = 1354
        if (dfItens['CFOPorigem1'] == 6400).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6400, 'CFOPdestino1'] = 1400
        if (dfItens['CFOPorigem1'] == 6403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6403, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6402).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6402, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6403, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6404).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6404, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6405).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6405, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6408).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6408, 'CFOPdestino1'] = 1408
        if (dfItens['CFOPorigem1'] == 6409).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6409, 'CFOPdestino1'] = 1409
        if (dfItens['CFOPorigem1'] == 6411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6411, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 6411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6411, 'CFOPdestino1'] = 2411
        if (dfItens['CFOPorigem1'] == 6414).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6414, 'CFOPdestino1'] = 1414
        if (dfItens['CFOPorigem1'] == 6415).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6415, 'CFOPdestino1'] = 1415
        if (dfItens['CFOPorigem1'] == 6500).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6500, 'CFOPdestino1'] = 1500
        if (dfItens['CFOPorigem1'] == 6504).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6504, 'CFOPdestino1'] = 1504
        if (dfItens['CFOPorigem1'] == 6505).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6505, 'CFOPdestino1'] = 1505
        if (dfItens['CFOPorigem1'] == 6550).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6550, 'CFOPdestino1'] = 1550
        if (dfItens['CFOPorigem1'] == 6551).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6551, 'CFOPdestino1'] = 2551
        if (dfItens['CFOPorigem1'] == 6553).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6553, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 6554).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6554, 'CFOPdestino1'] = 1554
        if (dfItens['CFOPorigem1'] == 6556).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6556, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6557).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6557, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6603).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6603, 'CFOPdestino1'] = 1603
        if (dfItens['CFOPorigem1'] == 6651).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6651, 'CFOPdestino1'] = 2651
        if (dfItens['CFOPorigem1'] == 6652).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6652, 'CFOPdestino1'] = 1652
        if (dfItens['CFOPorigem1'] == 6655).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6655, 'CFOPdestino1'] = 2652
        if (dfItens['CFOPorigem1'] == 6656).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6656, 'CFOPdestino1'] = 2653
        if (dfItens['CFOPorigem1'] == 6660).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6660, 'CFOPdestino1'] = 1660
        if (dfItens['CFOPorigem1'] == 6662).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6662, 'CFOPdestino1'] = 1662
        if (dfItens['CFOPorigem1'] == 6664).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6664, 'CFOPdestino1'] = 1664
        if (dfItens['CFOPorigem1'] == 6900).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6900, 'CFOPdestino1'] = 1900
        if (dfItens['CFOPorigem1'] == 6901).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6901, 'CFOPdestino1'] = 2901
        if (dfItens['CFOPorigem1'] == 6902).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6902, 'CFOPdestino1'] = 2902
        if (dfItens['CFOPorigem1'] == 6903).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6903, 'CFOPdestino1'] = 1903
        if (dfItens['CFOPorigem1'] == 6907).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6907, 'CFOPdestino1'] = 1907
        if (dfItens['CFOPorigem1'] == 6909).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6909, 'CFOPdestino1'] = 1909
        if (dfItens['CFOPorigem1'] == 6910).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6910, 'CFOPdestino1'] = 2910
        if (dfItens['CFOPorigem1'] == 6911).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6911, 'CFOPdestino1'] = 2911
        if (dfItens['CFOPorigem1'] == 6912).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6912, 'CFOPdestino1'] = 2912
        if (dfItens['CFOPorigem1'] == 6915).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6915, 'CFOPdestino1'] = 2915
        if (dfItens['CFOPorigem1'] == 6916).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6916, 'CFOPdestino1'] = 2916
        if (dfItens['CFOPorigem1'] == 6919).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6919, 'CFOPdestino1'] = 1919
        if (dfItens['CFOPorigem1'] == 6920).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6920, 'CFOPdestino1'] = 2920
        if (dfItens['CFOPorigem1'] == 6922).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6922, 'CFOPdestino1'] = 1922
        if (dfItens['CFOPorigem1'] == 6924).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6924, 'CFOPdestino1'] = 1924
        if (dfItens['CFOPorigem1'] == 6925).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6925, 'CFOPdestino1'] = 1925
        if (dfItens['CFOPorigem1'] == 6929).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6929, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6931).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6931, 'CFOPdestino1'] = 1931
        if (dfItens['CFOPorigem1'] == 6933).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6933, 'CFOPdestino1'] = 1933
        if (dfItens['CFOPorigem1'] == 6949).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6949, 'CFOPdestino1'] = 2949
        if (dfItens['CFOPorigem1'] == 7100).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7100, 'CFOPdestino1'] = 1100
        if (dfItens['CFOPorigem1'] == 7102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7102, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 7200).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7200, 'CFOPdestino1'] = 1200
        if (dfItens['CFOPorigem1'] == 7205).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7205, 'CFOPdestino1'] = 1205
        if (dfItens['CFOPorigem1'] == 7211).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7211, 'CFOPdestino1'] = 3211
        if (dfItens['CFOPorigem1'] == 7250).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7250, 'CFOPdestino1'] = 1250
        if (dfItens['CFOPorigem1'] == 7251).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7251, 'CFOPdestino1'] = 1251
        if (dfItens['CFOPorigem1'] == 7300).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7300, 'CFOPdestino1'] = 1300
        if (dfItens['CFOPorigem1'] == 7301).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7301, 'CFOPdestino1'] = 1301
        if (dfItens['CFOPorigem1'] == 7350).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7350, 'CFOPdestino1'] = 1350
        if (dfItens['CFOPorigem1'] == 7500).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7500, 'CFOPdestino1'] = 1500
        if (dfItens['CFOPorigem1'] == 7553).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7553, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 7651).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7651, 'CFOPdestino1'] = 2651
        if (dfItens['CFOPorigem1'] == 7900).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7900, 'CFOPdestino1'] = 1900
        if (dfItens['CFOPorigem1'] == 7930).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7930, 'CFOPdestino1'] = 3930
        if (dfItens['CFOPorigem1'] == 5105).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5105, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 6105).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6105, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 5405).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5405, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5100).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5100, 'CFOPdestino1'] = 1100
        if (dfItens['CFOPorigem1'] == 5102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5102, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5102, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5103).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5103, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5104).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5104, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5106).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5106, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 5111).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5111, 'CFOPdestino1'] = 1111
        if (dfItens['CFOPorigem1'] == 5117).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5117, 'CFOPdestino1'] = 1117
        if (dfItens['CFOPorigem1'] == 5120).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5120, 'CFOPdestino1'] = 1120
        if (dfItens['CFOPorigem1'] == 5122).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5122, 'CFOPdestino1'] = 1122
        if (dfItens['CFOPorigem1'] == 5124).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5124, 'CFOPdestino1'] = 1124
        if (dfItens['CFOPorigem1'] == 5125).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5125, 'CFOPdestino1'] = 1125
        if (dfItens['CFOPorigem1'] == 5150).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5150, 'CFOPdestino1'] = 1150
        if (dfItens['CFOPorigem1'] == 5200).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5200, 'CFOPdestino1'] = 1200
        if (dfItens['CFOPorigem1'] == 5202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5202, 'CFOPdestino1'] = 1202
        if (dfItens['CFOPorigem1'] == 5205).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5205, 'CFOPdestino1'] = 1205
        if (dfItens['CFOPorigem1'] == 5208).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5208, 'CFOPdestino1'] = 1208
        if (dfItens['CFOPorigem1'] == 5251).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5251, 'CFOPdestino1'] = 1251
        if (dfItens['CFOPorigem1'] == 5254).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5254, 'CFOPdestino1'] = 1254
        if (dfItens['CFOPorigem1'] == 5300).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5300, 'CFOPdestino1'] = 1300
        if (dfItens['CFOPorigem1'] == 5301).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5301, 'CFOPdestino1'] = 1301
        if (dfItens['CFOPorigem1'] == 5304).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5304, 'CFOPdestino1'] = 1304
        if (dfItens['CFOPorigem1'] == 5306).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5306, 'CFOPdestino1'] = 2306
        if (dfItens['CFOPorigem1'] == 5350).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5350, 'CFOPdestino1'] = 1350
        if (dfItens['CFOPorigem1'] == 5351).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5351, 'CFOPdestino1'] = 1351
        if (dfItens['CFOPorigem1'] == 5354).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5354, 'CFOPdestino1'] = 1354
        if (dfItens['CFOPorigem1'] == 5360).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5360, 'CFOPdestino1'] = 1360
        if (dfItens['CFOPorigem1'] == 5400).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5400, 'CFOPdestino1'] = 1400
        if (dfItens['CFOPorigem1'] == 5403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5403, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5402).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5402, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5403, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5408).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5408, 'CFOPdestino1'] = 1408
        if (dfItens['CFOPorigem1'] == 5409).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5409, 'CFOPdestino1'] = 1409
        if (dfItens['CFOPorigem1'] == 5411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5411, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 5411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5411, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 5412).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5412, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 5413).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5413, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 5414).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5414, 'CFOPdestino1'] = 1414
        if (dfItens['CFOPorigem1'] == 5415).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5415, 'CFOPdestino1'] = 1415
        if (dfItens['CFOPorigem1'] == 5450).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5450, 'CFOPdestino1'] = 1450
        if (dfItens['CFOPorigem1'] == 5451).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5451, 'CFOPdestino1'] = 1451
        if (dfItens['CFOPorigem1'] == 5500).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5500, 'CFOPdestino1'] = 1500
        if (dfItens['CFOPorigem1'] == 5504).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5504, 'CFOPdestino1'] = 1504
        if (dfItens['CFOPorigem1'] == 5505).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5505, 'CFOPdestino1'] = 1505
        if (dfItens['CFOPorigem1'] == 5550).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5550, 'CFOPdestino1'] = 1550
        if (dfItens['CFOPorigem1'] == 5551).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5551, 'CFOPdestino1'] = 1551
        if (dfItens['CFOPorigem1'] == 5554).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5554, 'CFOPdestino1'] = 1554
        if (dfItens['CFOPorigem1'] == 5556).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5556, 'CFOPdestino1'] = 1556
        if (dfItens['CFOPorigem1'] == 5557).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5557, 'CFOPdestino1'] = 1407
        if (dfItens['CFOPorigem1'] == 5600).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5600, 'CFOPdestino1'] = 1600
        if (dfItens['CFOPorigem1'] == 5602).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5602, 'CFOPdestino1'] = 1602
        if (dfItens['CFOPorigem1'] == 5603).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5603, 'CFOPdestino1'] = 1603
        if (dfItens['CFOPorigem1'] == 5605).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5605, 'CFOPdestino1'] = 1605
        if (dfItens['CFOPorigem1'] == 5651).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5651, 'CFOPdestino1'] = 2651
        if (dfItens['CFOPorigem1'] == 5652).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5652, 'CFOPdestino1'] = 1652
        if (dfItens['CFOPorigem1'] == 5655).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5655, 'CFOPdestino1'] = 1652
        if (dfItens['CFOPorigem1'] == 5656).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5656, 'CFOPdestino1'] = 1653
        if (dfItens['CFOPorigem1'] == 5660).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5660, 'CFOPdestino1'] = 1660
        if (dfItens['CFOPorigem1'] == 5662).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5662, 'CFOPdestino1'] = 1662
        if (dfItens['CFOPorigem1'] == 5664).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5664, 'CFOPdestino1'] = 1664
        if (dfItens['CFOPorigem1'] == 5900).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5900, 'CFOPdestino1'] = 1900
        if (dfItens['CFOPorigem1'] == 5901).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5901, 'CFOPdestino1'] = 1901
        if (dfItens['CFOPorigem1'] == 5902).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5902, 'CFOPdestino1'] = 1902
        if (dfItens['CFOPorigem1'] == 5903).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5903, 'CFOPdestino1'] = 1903
        if (dfItens['CFOPorigem1'] == 5907).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5907, 'CFOPdestino1'] = 1907
        if (dfItens['CFOPorigem1'] == 5909).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5909, 'CFOPdestino1'] = 1909
        if (dfItens['CFOPorigem1'] == 5910).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5910, 'CFOPdestino1'] = 1910
        if (dfItens['CFOPorigem1'] == 5911).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5911, 'CFOPdestino1'] = 1911
        if (dfItens['CFOPorigem1'] == 5912).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5912, 'CFOPdestino1'] = 1912
        if (dfItens['CFOPorigem1'] == 5915).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5915, 'CFOPdestino1'] = 1915
        if (dfItens['CFOPorigem1'] == 5916).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5916, 'CFOPdestino1'] = 1916
        if (dfItens['CFOPorigem1'] == 5920).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5920, 'CFOPdestino1'] = 1920
        if (dfItens['CFOPorigem1'] == 5922).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5922, 'CFOPdestino1'] = 1922
        if (dfItens['CFOPorigem1'] == 5924).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5924, 'CFOPdestino1'] = 1924
        if (dfItens['CFOPorigem1'] == 5925).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5925, 'CFOPdestino1'] = 1925
        if (dfItens['CFOPorigem1'] == 5926).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5926, 'CFOPdestino1'] = 1926
        if (dfItens['CFOPorigem1'] == 5929).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5929, 'CFOPdestino1'] = 1556
        if (dfItens['CFOPorigem1'] == 5931).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5931, 'CFOPdestino1'] = 1931
        if (dfItens['CFOPorigem1'] == 5933).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5933, 'CFOPdestino1'] = 1933
        if (dfItens['CFOPorigem1'] == 5949).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5949, 'CFOPdestino1'] = 1949
        if (dfItens['CFOPorigem1'] == 6100).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6100, 'CFOPdestino1'] = 1100
        if (dfItens['CFOPorigem1'] == 6102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6102, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6102, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6106).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6106, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6108).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6108, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6111).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6111, 'CFOPdestino1'] = 1111
        if (dfItens['CFOPorigem1'] == 6117).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6117, 'CFOPdestino1'] = 2117
        if (dfItens['CFOPorigem1'] == 6120).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6120, 'CFOPdestino1'] = 1120
        if (dfItens['CFOPorigem1'] == 6122).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6122, 'CFOPdestino1'] = 1122
        if (dfItens['CFOPorigem1'] == 6124).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6124, 'CFOPdestino1'] = 2124
        if (dfItens['CFOPorigem1'] == 6125).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6125, 'CFOPdestino1'] = 1125
        if (dfItens['CFOPorigem1'] == 6150).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6150, 'CFOPdestino1'] = 1150
        if (dfItens['CFOPorigem1'] == 6200).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6200, 'CFOPdestino1'] = 1200
        if (dfItens['CFOPorigem1'] == 6202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6202, 'CFOPdestino1'] = 2202
        if (dfItens['CFOPorigem1'] == 6202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6202, 'CFOPdestino1'] = 2202
        if (dfItens['CFOPorigem1'] == 6205).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6205, 'CFOPdestino1'] = 1205
        if (dfItens['CFOPorigem1'] == 6208).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6208, 'CFOPdestino1'] = 1208
        if (dfItens['CFOPorigem1'] == 6250).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6250, 'CFOPdestino1'] = 1250
        if (dfItens['CFOPorigem1'] == 6251).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6251, 'CFOPdestino1'] = 1251
        if (dfItens['CFOPorigem1'] == 6254).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6254, 'CFOPdestino1'] = 1254
        if (dfItens['CFOPorigem1'] == 6300).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6300, 'CFOPdestino1'] = 1300
        if (dfItens['CFOPorigem1'] == 6301).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6301, 'CFOPdestino1'] = 1301
        if (dfItens['CFOPorigem1'] == 6304).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6304, 'CFOPdestino1'] = 1304
        if (dfItens['CFOPorigem1'] == 6306).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6306, 'CFOPdestino1'] = 2306
        if (dfItens['CFOPorigem1'] == 6350).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6350, 'CFOPdestino1'] = 1350
        if (dfItens['CFOPorigem1'] == 6351).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6351, 'CFOPdestino1'] = 1351
        if (dfItens['CFOPorigem1'] == 6354).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6354, 'CFOPdestino1'] = 1354
        if (dfItens['CFOPorigem1'] == 6400).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6400, 'CFOPdestino1'] = 1400
        if (dfItens['CFOPorigem1'] == 6403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6403, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6402).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6402, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6403).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6403, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6404).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6404, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6405).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6405, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6408).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6408, 'CFOPdestino1'] = 1408
        if (dfItens['CFOPorigem1'] == 6409).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6409, 'CFOPdestino1'] = 1409
        if (dfItens['CFOPorigem1'] == 6411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6411, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 6411).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6411, 'CFOPdestino1'] = 2411
        if (dfItens['CFOPorigem1'] == 6414).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6414, 'CFOPdestino1'] = 1414
        if (dfItens['CFOPorigem1'] == 6415).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6415, 'CFOPdestino1'] = 1415
        if (dfItens['CFOPorigem1'] == 6500).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6500, 'CFOPdestino1'] = 1500
        if (dfItens['CFOPorigem1'] == 6504).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6504, 'CFOPdestino1'] = 1504
        if (dfItens['CFOPorigem1'] == 6505).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6505, 'CFOPdestino1'] = 1505
        if (dfItens['CFOPorigem1'] == 6550).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6550, 'CFOPdestino1'] = 1550
        if (dfItens['CFOPorigem1'] == 6551).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6551, 'CFOPdestino1'] = 2551
        if (dfItens['CFOPorigem1'] == 6553).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6553, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 6554).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6554, 'CFOPdestino1'] = 1554
        if (dfItens['CFOPorigem1'] == 6556).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6556, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6557).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6557, 'CFOPdestino1'] = 2557
        if (dfItens['CFOPorigem1'] == 6603).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6603, 'CFOPdestino1'] = 1603
        if (dfItens['CFOPorigem1'] == 6651).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6651, 'CFOPdestino1'] = 2651
        if (dfItens['CFOPorigem1'] == 6652).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6652, 'CFOPdestino1'] = 1652
        if (dfItens['CFOPorigem1'] == 6655).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6655, 'CFOPdestino1'] = 2652
        if (dfItens['CFOPorigem1'] == 6656).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6656, 'CFOPdestino1'] = 2653
        if (dfItens['CFOPorigem1'] == 6660).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6660, 'CFOPdestino1'] = 1660
        if (dfItens['CFOPorigem1'] == 6662).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6662, 'CFOPdestino1'] = 1662
        if (dfItens['CFOPorigem1'] == 6664).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6664, 'CFOPdestino1'] = 1664
        if (dfItens['CFOPorigem1'] == 6900).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6900, 'CFOPdestino1'] = 1900
        if (dfItens['CFOPorigem1'] == 6901).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6901, 'CFOPdestino1'] = 2901
        if (dfItens['CFOPorigem1'] == 6902).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6902, 'CFOPdestino1'] = 2902
        if (dfItens['CFOPorigem1'] == 6903).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6903, 'CFOPdestino1'] = 1903
        if (dfItens['CFOPorigem1'] == 6907).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6907, 'CFOPdestino1'] = 1907
        if (dfItens['CFOPorigem1'] == 6909).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6909, 'CFOPdestino1'] = 1909
        if (dfItens['CFOPorigem1'] == 6910).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6910, 'CFOPdestino1'] = 2910
        if (dfItens['CFOPorigem1'] == 6911).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6911, 'CFOPdestino1'] = 2911
        if (dfItens['CFOPorigem1'] == 6912).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6912, 'CFOPdestino1'] = 2912
        if (dfItens['CFOPorigem1'] == 6915).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6915, 'CFOPdestino1'] = 2915
        if (dfItens['CFOPorigem1'] == 6916).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6916, 'CFOPdestino1'] = 2916
        if (dfItens['CFOPorigem1'] == 6919).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6919, 'CFOPdestino1'] = 1919
        if (dfItens['CFOPorigem1'] == 6920).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6920, 'CFOPdestino1'] = 2920
        if (dfItens['CFOPorigem1'] == 6922).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6922, 'CFOPdestino1'] = 1922
        if (dfItens['CFOPorigem1'] == 6924).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6924, 'CFOPdestino1'] = 1924
        if (dfItens['CFOPorigem1'] == 6925).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6925, 'CFOPdestino1'] = 1925
        if (dfItens['CFOPorigem1'] == 6929).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6929, 'CFOPdestino1'] = 2556
        if (dfItens['CFOPorigem1'] == 6931).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6931, 'CFOPdestino1'] = 1931
        if (dfItens['CFOPorigem1'] == 6933).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6933, 'CFOPdestino1'] = 1933
        if (dfItens['CFOPorigem1'] == 6949).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6949, 'CFOPdestino1'] = 2949
        if (dfItens['CFOPorigem1'] == 7100).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7100, 'CFOPdestino1'] = 1100
        if (dfItens['CFOPorigem1'] == 7102).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7102, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 7200).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7200, 'CFOPdestino1'] = 1200
        if (dfItens['CFOPorigem1'] == 7205).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7205, 'CFOPdestino1'] = 1205
        if (dfItens['CFOPorigem1'] == 7211).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7211, 'CFOPdestino1'] = 3211
        if (dfItens['CFOPorigem1'] == 7250).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7250, 'CFOPdestino1'] = 1250
        if (dfItens['CFOPorigem1'] == 7251).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7251, 'CFOPdestino1'] = 1251
        if (dfItens['CFOPorigem1'] == 7300).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7300, 'CFOPdestino1'] = 1300
        if (dfItens['CFOPorigem1'] == 7301).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7301, 'CFOPdestino1'] = 1301
        if (dfItens['CFOPorigem1'] == 7350).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7350, 'CFOPdestino1'] = 1350
        if (dfItens['CFOPorigem1'] == 7500).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7500, 'CFOPdestino1'] = 1500
        if (dfItens['CFOPorigem1'] == 7553).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7553, 'CFOPdestino1'] = 1553
        if (dfItens['CFOPorigem1'] == 7651).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7651, 'CFOPdestino1'] = 2651
        if (dfItens['CFOPorigem1'] == 7900).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7900, 'CFOPdestino1'] = 1900
        if (dfItens['CFOPorigem1'] == 7930).any(): dfItens.loc[dfItens['CFOPorigem1'] == 7930, 'CFOPdestino1'] = 3930
        if (dfItens['CFOPorigem1'] == 5105).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5105, 'CFOPdestino1'] = 1102
        if (dfItens['CFOPorigem1'] == 6105).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6105, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 5401).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5401, 'CFOPdestino1'] = 1403
        if (dfItens['CFOPorigem1'] == 5410).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5410, 'CFOPdestino1'] = 1411
        if (dfItens['CFOPorigem1'] == 6401).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6401, 'CFOPdestino1'] = 2403
        if (dfItens['CFOPorigem1'] == 6101).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6101, 'CFOPdestino1'] = 2102
        if (dfItens['CFOPorigem1'] == 6410).any(): dfItens.loc[dfItens['CFOPorigem1'] == 6410, 'CFOPdestino1'] = 2411
        if (dfItens['CFOPorigem1'] == 2202).any(): dfItens.loc[dfItens['CFOPorigem1'] == 2202, 'CFOPdestino1'] = 2202
        if (dfItens['CFOPorigem1'] == 5101).any(): dfItens.loc[dfItens['CFOPorigem1'] == 5101, 'CFOPdestino1'] = 1101

                
            

    #                           Tratando coluna de Tipo Item

        if (dfItens['TipoItem'] == 'Uso e Consumo').any(): dfItens.loc[dfItens['TipoItem'] == 'Uso e Consumo', 'TipoItem'] = '07'
        if (dfItens['TipoItem'] == 'Remessa').any(): dfItens.loc[dfItens['TipoItem'] == 'Remessa', 'TipoItem'] = '99'        
        if (dfItens['TipoItem'] == 'Matéria Prima').any(): dfItens.loc[dfItens['TipoItem'] == 'Matéria Prima', 'TipoItem'] = '01'        
        if (dfItens['TipoItem'] == 'Revenda').any(): dfItens.loc[dfItens['TipoItem'] == 'Revenda', 'TipoItem'] = '00'        
        if (dfItens['TipoItem'] == 'Industrialização').any(): dfItens.loc[dfItens['TipoItem'] == 'Industrialização', 'TipoItem'] = '01'
        if  dfItens['TipoItem'].isnull().values.any(): dfItens.fillna({'TipoItem': '00'}, inplace=True)

    
    
          #ORDENAR AS COLUNAS
        ordemColunas = ["CNPJ",	"CodigoItemFornecedor",	"CodigoItemEmpresa", 
                        "PISCofinsCst",	"PISAliq",	"COFINSAliq","PISAliqReais", "COFINSAliqReais",
                        "TipoItem",	"CFOPorigem1",	"CFOPdestino1",	"Variacao1",]    
        dfItens = dfItens[ordemColunas]

        # if messagebox.askyesno("SAlvar Arquivo", "Deseja salvar o arquivo"):
        #      file_path_out = filedialog.asksaveasfilename(defaultextension = ".csv", filetypes=[("Arquivos CSV", "*.csv")])
        #      dfItens.to_csv(file_path_out, sep=';', index=False)
        #      messagebox.showinfo("Sucessp", "Arquivo salvo com sucesso !")
        # else:
        #      messagebox.showinfo("Cancelar", "Operacao cancelada pelo usuario !")
        
        print (dfItens.head(5))   
except Exception as error:
        print("MENSAGEM DE ERRO:", error)