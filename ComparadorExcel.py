# Python program to compare two excel files
#Autor: Rafael Henrique da Rosa
 

import openpyxl,math,os,sys
from openpyxl.styles import PatternFill,Alignment,Font
from openpyxl.utils import get_column_letter
from openpyxl.styles.borders import Border, Side
import tkinter as tk
from tkinter import ttk,Label,messagebox,filedialog as fd


######################################################################################################################################################
#PARTE TK

root = tk.Tk()
path1 = ""
path2 = ""
root.title('Comparador Excel (.xlsx) V0.1')
root.resizable(False, False)
root.geometry('365x200')

def clicked():
    if(path1 != "" and path2 != ""):
        root.destroy()

def myinfo():
    messagebox.showinfo("Info","Autor: Rafael Henrique da Rosa\nEstagiário Itaipu Binacional - SMIN.DT - Março de 2022\n\
O algoritmo compara duas tabelas e mostra em uma planilha de resultado as linhas excluidas,novas e discrepantes")

def select_file():
    filetypes = (('Excel Files', '*.xlsx'),('All files', '*.*'))
    filename = fd.askopenfilename(title='Selecionar Banco antigo',initialdir='‪C:\\Users\\rafa6899\\Downloads\\',filetypes=filetypes)
    global path1
    path1 = filename
    path1_ = path1
    while(path1_.find("/") != -1):
        path1_ = path1_[1:]
    lbl1.configure(text=path1_)
    
def select_file2():
    filetypes = (('Excel Files', '*.xlsx'), ('All files', '*.*') )

    filename = fd.askopenfilename(title='Selecionar Banco novo',initialdir='‪C:\\Users\\rafa6899\\Downloads\\',filetypes=filetypes)
    global path2
    path2 = filename
    path2_ = path2
    while(path2_.find("/") != -1):
       path2_ = path2_[1:]
    lbl2.configure(text=path2_)
 
lbl4 = Label(root, text="")
lbl4.configure(text="*A comparação pode demorar de acordo com o tamanho do banco")
lbl4.grid(column=0, row=90)  
  
# open button
open_button = ttk.Button(root,text='Selecionar Banco antigo',command=select_file)
open_button.grid(column=0, row=0)
lbl1 = Label(root, text="")
lbl1.configure(text="")
lbl1.grid(column=0, row=1)

open_button2 = ttk.Button(root,text='Selecionar Banco novo',command=select_file2)
open_button2.grid(column=0, row=10)
lbl2 = Label(root, text="")
lbl2.configure(text="")
lbl2.grid(column=0, row=30)

open_button3 = ttk.Button(root,text='Iniciar',command=clicked)
#open_button3.pack(expand=False)
open_button3.grid(column=0, row=40)
lbl3 = Label(root, text="")
lbl3.configure(text="")
lbl3.grid(column=0, row=60)  

# open button
open_button4 = tk.Button(root,text='?',command=myinfo)
open_button4.config(height = 1,width = 1)
#open_button3.pack(expand=False)
open_button4.grid(column=0, row=100)

# run the application
root.mainloop()

######################################################################################################################################################
#PARTE DE COMPARAÇÃO
    
#Se não foram selecionados arquivos não inicia a parte da comparação
if(path1 == "" or path2 == ""):
    exit() #Esse exit funciona porem gera um erro

#define borda e cores para preenchimento de celulas
thin_border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
vermelho = PatternFill(start_color='ff0000', end_color='ff0000', fill_type='solid')
cinza = PatternFill(start_color='787878', end_color='787878', fill_type='solid')
vermelhoClaro = PatternFill(start_color='fa9696', end_color='fa9696', fill_type='solid')
verde = PatternFill(start_color='82e89d', end_color='82e89d', fill_type='solid')


#Abre os arquivos selecionados e seleciona o sheet ativo como o que será copiado

try:
    banco1_obj = openpyxl.load_workbook(path1)
except:
    try:
        path1 = path1.replace("/", "\\")
        banco1_obj = openpyxl.load_workbook(path1)
    except:
        print("ai realmente não ta abrindo o arquivo antigo")

try:
    banco2_obj = openpyxl.load_workbook(path2)
except:
    try:
        path2 = path1.replace("/", "\\")
        banco2_obj = openpyxl.load_workbook(path2)
    except:
        print("ai realmente não ta abrindo o arquivo novo")

banco1Stheet_obj = banco1_obj.active
banco2Stheet_obj = banco2_obj.active

#Tira o caminho e deixa só o nome dos arquivos selecionados
while(path1.find("/") != -1):
   path1 = path1[1:]  
while(path2.find("/") != -1):
    path2 = path2[1:]

#Cria um Workbook de output, cria dois sheets e  muda seus nomes para o nome dos arquivos selecionados
wb = openpyxl.Workbook()
sheetb1 = wb.active
sheetb2 = wb.create_sheet()
sheetb1.title = path1
sheetb2.title = path2

#Copia célula por celula dos arquivos selecionados para o arqvuivo de output
#O Openyxl não possui metodo para copiar o sheet inteiro de uma vez
for row in banco1Stheet_obj:
    for cell in row:
        sheetb1[cell.coordinate].value = cell.value
for row in banco2Stheet_obj:
    for cell in row:
        sheetb2[cell.coordinate].value = cell.value

#cria um terceiro sheet no arquivo de output de relatório das ocorrencias        
sheetResul = wb.create_sheet()
sheetResul.title = "RELATÓRIO"

#Deixa a primeira linha dos 3 sheets de output congelada para melhorar a visualização
sheetb2.freeze_panes = sheetb2['A2']
sheetb1.freeze_panes = sheetb2['A2']       
sheetResul.freeze_panes = sheetb2['A2']       


#Pinta a primeira linha do sheet 'banco1' do output de cinza, alinha e deixa em negrito
for rows in sheetb1.iter_rows(min_row=1, max_row=1, min_col=1):
    for cell in rows:
      cell.fill = cinza
      cell.alignment = Alignment(horizontal='center')
      cell.font = Font(bold= True)
      
#Pinta a primeira linha do sheet 'banco2' do output de cinza, alinha e deixa em negrito
for rows in sheetb2.iter_rows(min_row=1, max_row=1, min_col=1):
    for i in range(sheetb2.max_column):
      sheetb2.cell(1,i+1).fill = cinza
      sheetb2.cell(1,i+1).alignment = Alignment(horizontal='center')
      if(sheetb1.cell(1,i+2).value == sheetb2.cell(1,i+2).value):
          sheetResul.cell(1,i+2).value = sheetb1.cell(1,i+2).value
      else:
          #Se o nome das colunas não sao iguais printa um erro e para o programa
          print("NOME DAS COLUNAS NÃO SÃO IGUAIS")
          sys.exit(0)
      sheetb2.cell(1,i+1).font = Font(bold= True)
      
#Pinta a primeira linha do sheet 'Relatório' do output de cinza, alinha e deixa em negrito
#i+1 desloca as colunas 1 coluna para a direita para anexar uma coluna index no inicio do relatório
#a coluna index mostra o numero da linha que foi copiada para o relatório      
for rows in sheetResul.iter_rows(min_row=1, max_row=1, min_col=1):
    sheetResul.cell(1,1).value = "INDEX"
    for i in range(sheetResul.max_column):
      sheetResul.cell(1,i+1).fill = cinza
      sheetResul.cell(1,i+1).alignment = Alignment(horizontal='center')
      sheetResul.cell(1,i+1).font = Font(bold= True)
 

#Ajusta o comprimento das colunas do sheet 'banco1'    
dims = {}
for row in sheetb1.rows:
    for cell in row:
        if cell.value:
            dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))    
   
for i, column_width in dims.items():  # ,1 to start at 1
    sheetb1.column_dimensions[get_column_letter(i)].width = int(math.ceil(column_width*1.42))

#Ajusta o comprimento das colunas do sheet 'banco2' e deixa o resultado igual o banco2        
for row in sheetb2.rows:
    for cell in row:
        if cell.value:
            dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))    

for i, column_width in dims.items():
    sheetb2.column_dimensions[get_column_letter(i)].width = int(math.ceil(column_width*1.42))
    sheetResul.column_dimensions[get_column_letter(i+1)].width = int(math.ceil(column_width*1.42))
    
    
######################################################################################################################################################
#Parte do resultado

#Array para armazenar as linhas iguais encontradas e as discr   epancias    
encontrado = {}
discrep = {}

#Zera os arrays anteriores (Poderia utilizar outra lógica para evitar)
for i in range(sheetb1.max_row):
    encontrado[i] = 0
    discrep[i] = 0


    
#k representa a linha que ta sendo escrita no sheet de relatorio
k=3
sheetResul.cell(k,1).value = "LINHAS EXCLUIDAS (presentes no banco 1 e não no banco 2):"
k+=1

#Linhas excluidas (presentes no banco 1 e não no banco 2):

#Percorre cada linha nos sheets banco1 e banco2       
for i in range(sheetb1.max_row):
    for j in range(sheetb2.max_row):
        #Se os valores das colunas 3 e 4 são igais
        if(sheetb1.cell(i+1,3).value == sheetb2.cell(j+1,3).value and sheetb1.cell(i+1,4).value == sheetb2.cell(j+1,4).value):
            #marca a linha como encontrada nos dois bancos
            encontrado[i] = j+1
            #checa se existem discrepancias percorrendo cada coluna da linha
            for l in range(sheetb1.max_column):
                if(sheetb1.cell(i+1,l+1).value != sheetb2.cell(j+1,l+1).value):
                    discrep[i] = j
                    #muda a cor das discrepancias
                    sheetb1.cell(i+1,l).fill = vermelho
                    sheetb1.cell(i+1,3).fill = vermelho
                    sheetb1.cell(i+1,4).fill = vermelho
                    sheetb2.cell(j+1,l).fill = vermelho
                    sheetb2.cell(j+1,3).fill = vermelho
                    sheetb2.cell(j+1,4).fill = vermelho
                    
            break
    #Se a linha não for encontrada    
    if(encontrado[i] == 0):
       #Marca a linha como uma linha excluida no sheet banco1 e copia para o sheet de relatório
       for l in range(sheetb1.max_column-1):
           sheetb1.cell(i+1,l+1).fill = vermelhoClaro
           sheetb1.cell(i+1,l+1).fill = vermelhoClaro
           sheetResul.cell(k,l+1+1).value = sheetb1.cell(i+1,l+1).value
           sheetResul.cell(k,l+1).fill = vermelhoClaro
           sheetResul.cell(k,l+1).border = thin_border
           sheetResul.cell(k, 1).value = i+1
       sheetResul.cell(k, 1).font = Font(bold= True)
       sheetResul.cell(k, 1).fill = cinza 
       k+=1
       
#Linhas discrepantes:
    
k+=2
sheetResul.cell(k,1).value = "LINHAS DISCREPANTES:"
k+=1
for i in range(sheetb1.max_row):
    if(discrep[i] !=0):
        sheetResul.cell(k,1).value = path1
        sheetResul.cell(k+2,1).value = path2
        k+=1
        sheetResul.cell(k, 1).value = i+1
        sheetResul.cell(k, 1).font = Font(bold= True)
        sheetResul.cell(k, 1).fill = cinza
        sheetResul.cell(k+2, 1).value = discrep[i]+1
        sheetResul.cell(k+2, 1).font = Font(bold= True)
        sheetResul.cell(k+2, 1).fill = cinza
        sheetResul.cell(k,l+1+1).fill = vermelho
        for l in range(sheetb1.max_column-1):
           sheetResul.cell(k,l+1+1).value = sheetb1.cell(i+1,l+1).value
           sheetResul.cell(k,l+1).border = thin_border
           sheetResul.cell(k+2,l+1+1).value = sheetb2.cell(discrep[i]+1,l+1).value
           sheetResul.cell(k+2,l+1).border = thin_border
           if(sheetResul.cell(k,l+1+1).value != sheetResul.cell(k+2,l+1+1).value):
               sheetResul.cell(k,l+1+1).fill = vermelho
               sheetResul.cell(k+2,l+1+1).fill = vermelho
        k +=4
k +=3


#Linhas Novas (presentes no banco 2 e não no banco 1):

sheetResul.cell(k,1).value = "LINHAS NOVAS (Presentes somente no Banco2):" 
k +=1
for i in range(sheetb2.max_row):
    flag = 0
    for j in range(sheetb1.max_row):
        if(i+1 == encontrado[j]):
            flag = 1
            break
    if(flag==0):
        for l in range(sheetb2.max_column-1):
         sheetResul.cell(k,l+1+1).value = sheetb2.cell(i+1,l+1).value
         sheetResul.cell(k,l+1).border = thin_border
         sheetResul.cell(k,l+1).fill = verde
         sheetb2.cell(i+1,l+1).fill = verde
         sheetb2.cell(i+1,l+1).fill = verde
        sheetResul.cell(k, 1).value = i+1
        sheetResul.cell(k, 1).font = Font(bold= True)
        sheetResul.cell(k, 1).fill = cinza
        k +=1


#Retira o grid do sheet de resultados        
sheetResul.sheet_view.showGridLines = False
#salva o arquivo output na mesma pasta do arquivo py e abre
wb.save("output.xlsx")
os.system("start EXCEL.EXE output.xlsx")

#https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size



