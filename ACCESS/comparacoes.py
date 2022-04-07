import pandas as pd
tabela_1 = pd.DataFrame({
'Nome':['João',  'Pedro' , 'Caio','Jorge','João'], 
'Sobrenome':['silva',  'Melo' , 'Rocha','Ribeiro','Rosa'], 
'Telefone': ['12121',  '565656', '787878','63001','343434',], 
'Carros': ['azul', 'verde' , 'amarelo','roxo', 'preto'],
'altura': ['1.74',  '1.85', '1.5','1.7','1.84',]})


tabela_2 = pd.DataFrame({
'Nome':['João', 'João', 'Pedro' , 'Caio','Augusto'], 
'Sobrenome':['silva', 'Sauro', 'Melo' , 'Rocha','Costa'], 
'Telefone': ['12121', '343435', '565656', '787878','4574'], 
'Carros': ['azul', 'preto', 'branco' , 'amarelo','roxo'], 
'altura': ['1.74', '1.57', '1.85', '1.5','1.58']})

#
#tabela_3 = tabela_1[~tabela_1['Carros'].isin(tabela_2['Carros'])]
tabela_3 = tabela_2[0:0]
tabela_4 = tabela_2[0:0]
tabela_5 = tabela_2[0:0]


#diferenças com base no carro



#linhas da tabela 1 que tem na tabela 2 com base no nome
#tabela_3 = pd.concat([tabela_3,tabela_1[tabela_1['Nome'].isin(tabela_2['Nome'])]],ignore_index = False)

#linhas da tabela 3 que tem na tabela 2 com base no telefone
#tabela_4 = pd.concat([tabela_4,tabela_3[~tabela_3['Telefone'].isin(tabela_2['Telefone'])]],ignore_index = False)

#linhas que só tem na tabela 1


tabela_5 = tabela_1[tabela_1.set_index(['Nome','Sobrenome']).index.isin(tabela_2.set_index(['Nome','Sobrenome']).index)]
print(tabela_5)
tabela_6 = tabela_5
for i in range(2, len(tabela_1.columns)):
    col = tabela_1.columns[i]
    print(col)
    tabela_6 = tabela_6[tabela_6.set_index([col]).index.isin(tabela_2.set_index([col]).index)]
    print("---------------------------------------------")
    print(tabela_6)
    
tabela_5 = tabela_5[~tabela_5.set_index(['Nome','Sobrenome']).index.isin(tabela_6.set_index(['Nome','Sobrenome']).index)]
print("---------------------------------------------")
print(tabela_5)
#tabela_3 = pd.concat([tabela_3,tabela_3[tabela_3['Nome'].isin(tabela_2['Nome'])]],ignore_index = False)
#tabela_4 = pd.concat([tabela_4,tabela_2[~tabela_2['Nome'].isin(tabela_1['Nome'])]],ignore_index = False)


#print(tabela_1[~tabela_1['Carros'].isin(tabela_2['Carros'])])



# =============================================================================
# x = ['RTUNO','PNTNO']
# print(x)
# x.append('test')
# print (x)
# =============================================================================

print(tabela_1.index.values)