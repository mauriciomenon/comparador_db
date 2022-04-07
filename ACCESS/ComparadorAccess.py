import subprocess
import pandas as pd
import os
import tkinter as tk
from tkinter.messagebox import showinfo
from tkinter import ttk,Label,messagebox,filedialog as fd
from pandastable import Table
from pandastable import config

def compara():
    
        
    global table_novas
    global table_excluidas
    global table_discrep
    global table1
    global table2
    os.path.dirname(os.path.realpath(__file__))
    
    #file1 = 'ACCESS_ANTIGO.accdb'
    #file2 = 'ACCESS_NOVO.accdb'
    file1 = path1
    file2 = path2
    path = os.path.dirname(os.path.realpath(__file__)) + "\\mdbtools"

    
    global selected_table
    export_command = path + '\\mdb-export.exe ' + file1 
    export_command += ' '
    export_command += selected_table + '  > temp.csv'
    subprocess.run(['cmd.exe', '/c',export_command])
    
    table1 = pd.read_csv('temp.csv',sep=',',encoding='iso-8859-1')
    os.remove("temp.csv")
    
    export_command = path + '\\mdb-export.exe ' + file2
    export_command += ' '
    export_command += selected_table + '  > temp.csv'
    subprocess.run(['cmd.exe', '/c',export_command])
    
    table2 = pd.read_csv('temp.csv',sep=',',encoding='iso-8859-1')
    os.remove("temp.csv")
    
    
    
    table_novas = table2[0:0]
    table_excluidas = table2[0:0]
    table_discrep = table2[0:0]
    
    table_excluidas = table1[~table1.set_index(['RTUNO','PNTNO']).index.isin(table2.set_index(['RTUNO','PNTNO']).index)]
    
    
    
    
    table_novas = table2[~table2.set_index(['RTUNO','PNTNO']).index.isin(table1.set_index(['RTUNO','PNTNO']).index)]
    
    
    
    table_discrep1 = table1[table1.set_index(['RTUNO','PNTNO']).index.isin(table2.set_index(['RTUNO','PNTNO']).index)]
    table_aux = table_discrep1
       
    for col in table1.columns:
       if(col != 'RTUNO' and col!= 'PNTNO'):
          col_test = ['RTUNO','PNTNO']
          col_test.append(col)
          table_aux = table_aux[table_aux.set_index(col_test).index.isin(table2.set_index(col_test).index)]
    
    table_discrep1 = table_discrep1[~table_discrep1.set_index(['RTUNO','PNTNO']).index.isin(table_aux.set_index(['RTUNO','PNTNO']).index)]
    
    
    table_discrep2 = table2[table2.set_index(['RTUNO','PNTNO']).index.isin(table1.set_index(['RTUNO','PNTNO']).index)]
    table_aux = table_discrep2
    
    for col in table2.columns:
        if(col != 'RTUNO' and col!= 'PNTNO'):
            col_test = ['RTUNO','PNTNO']
            col_test.append(col)
            table_aux = table_aux[table_aux.set_index(col_test).index.isin(table1.set_index(col_test).index)]
    
    table_discrep2 = table_discrep2[~table_discrep2.set_index(['RTUNO','PNTNO']).index.isin(table_aux.set_index(['RTUNO','PNTNO']).index)]
    
    
    
    if(table_discrep1.shape[0] == table_discrep2.shape[0]):
        table_discrep1.insert(loc=0, column='Arquivo', value='path1')
        table_discrep2.insert(loc=0, column='Arquivo', value='path2')
        table_discrep = table_discrep2[0:0]
        for i in range(table_discrep2.shape[0]):
            table_discrep = pd.concat([table_discrep,table_discrep1.iloc[[i]]],ignore_index = False) 
            for j in range(table_discrep2.shape[0]):
                if(table_discrep2.iat[j,1] == table_discrep1.iat[i,1] and table_discrep2.iat[j,2] == table_discrep1.iat[i,2]):
                    table_discrep = pd.concat([table_discrep,table_discrep2.iloc[[j]]],ignore_index = False) 
                    break
    else:
        print("ACONTECEU ALGUMA COISA ERRADA NA PARTE DAS LINHAS DISCREPANTES")

    pt1.model.df =table1
    pt1.autoResizeColumns()
    pt1.redraw()
    pt2.model.df =table2
    pt2.autoResizeColumns()
    pt2.redraw()
    pt_resul_discrep.model.df = table_discrep
    pt_resul_discrep.autoResizeColumns()
    pt_resul_discrep
    pt_resul_novas.model.df = table_novas
    pt_resul_novas.autoResizeColumns()
    pt_resul_novas.redraw()
    pt_resul_excluidas.model.df = table_excluidas
    pt_resul_excluidas.autoResizeColumns()
    pt_resul_excluidas.redraw()

    
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################
########################################################################################

df = pd.DataFrame({
    'A': ['','','','','','',],
    'B': ['','','','','','',],
    'C': ['','','','','','',],
    'D': ['','','','','','',],
})


#TK janela principal
def option_changed(self, *args):
    
    print('asaasdasdasd')

#getting screen width and height of display
root=tk.Tk()
width= root.winfo_screenwidth() 
height= root.winfo_screenheight()
#setting tkinter root size
root.geometry("%dx%d" % (width, height))
root.title("COMPARADOR ACCESS v0.1")
root.state("zoomed") 

def select_table():
    global output_tables
    path = os.path.dirname(os.path.realpath(__file__)) + "\\mdbtools"
    output_tables = subprocess.check_output([path + '\\mdb-tables.exe', path1]).decode()
    output_tables = output_tables.split()
    
# =============================================================================
    #########################################################################################
    #TK escolhe a tabela
    
    label = ttk.Label(text="Selecione a tabela para comparar:")
    label.place(x=(width/2)-100, y = 0,height = 30, width = 200)
    #label.pack(fill=tk.X, padx=5, pady=5)
    selected_month = tk.StringVar()
    month_cb = ttk.Combobox(root,width = 50,textvariable=selected_month)
    month_cb['values'] = output_tables
    month_cb['state'] = 'readonly'
    month_cb.pack(fill=tk.X, padx=5, pady=5)
    month_cb.place(x=(width/2)-110, y = 30,height = 30, width = 200)
    def month_changed(event):
        global selected_table
        selected_table = selected_month.get()
        compara()
    
    month_cb.bind('<<ComboboxSelected>>', month_changed)
    
     

# =============================================================================

#################################################################################################

def select_file():
    file_types = (('Access Files', '*.accdb'),('All files', '*.*'))
    file_name = fd.askopenfilename(title='Selecionar Banco antigo',filetypes=file_types)
    global path1
    path1 = file_name
#    lbl1.configure(text=file_name)
    
def select_file2():
    file_types = (('Access Files', '*.accdb'),('All files', '*.*'))
    file_name = fd.askopenfilename(title='Selecionar Banco antigo',filetypes=file_types)
    global path1
    global path2
    path2 = file_name
    if(path1 != "" and path2 != ""):
        select_table()
#    lbl2.configure(text=file_name)
        

#lbl1 = ttk.Label(root, text="dsfsdfsfsfdsfsf")
#lbl1.configure(text="")
#lbl1.place(x=(width/2)-200-50, y =80,height = 40, width = 220)

#lbl2 = Label(root, text="")
#lbl2.configure(text="")
#lbl2.place(x=(width/2)+50, y = 80,height = 40, width = 220)


menubar = tk.Menu(root)

filemenu = tk.Menu(menubar,tearoff=0)
filemenu.add_command(label="ABRIR ARQUIVO ANTIGO",command=select_file)
filemenu.add_command(label="ABRIR ARQUIVO NOVO",command=select_file2)
filemenu.add_command(label="SAIR")
helpmenu = tk.Menu(menubar,tearoff=0)
helpmenu.add_command(label="Como usar")
helpmenu.add_command(label="Sobre o programa")
exportmenu = tk.Menu(menubar,tearoff=0)
exportmenu.add_command(label="Exportar CSV tabela antiga")
exportmenu.add_command(label="Exportar CSV tabela nova")
exportmenu.add_command(label="Exportar CSV relatório")
exportmenu.add_command(label="Exportar CSV Completo")


menubar.add_cascade(label="Arquivo", menu=filemenu)
menubar.add_cascade(label="Exportar", menu=exportmenu)
menubar.add_cascade(label="Ajuda", menu=helpmenu)
root.config(menu=menubar)



tabControl = ttk.Notebook(root)
tabControl.place(x=0, y =70,height = height, width = width)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tab3 = ttk.Frame(tabControl)
tabControl.add(tab1, text ='RELATÓRIO')
tabControl.add(tab2, text ='ARQUIVO ANTIGO')
tabControl.add(tab3, text ='ARQUIVO NOVO')

frame1 = tk.Frame(tab2)
frame1.place(x=0, y =0,height = height-178, width = width)
pt1 = Table(frame1)
pt1.model.df =df
pt1.autoResizeColumns()
pt1.show()
pt1.autoResizeColumns()
pt1.redraw()

frame2 = tk.Frame(tab3)
frame2.place(x=0, y =0,height = height-178, width = width)
pt2 = Table(frame2)
pt2.model.df =df
pt2.autoResizeColumns()
pt2.show()
pt2.autoResizeColumns()
#pt2.columncolors['RTUNO'] = '#ff0000'
#pt2.drawRect(4, 4, color='#ff0000', tag=None, delete=0)
#pt2.setRowColors(rows=2, clr='#ff0000',  cols=[1,3,5])
pt2.autoResizeColumns()
pt2.redraw()
#xx = pt2.colorRows()


lbl_discrep = ttk.Label(tab1, text="LINHAS DISCREPANTES:",
                          font='Helvetica 12 bold')
#lbl1.configure(text="")
lbl_discrep.place(x=0, y =0,height = 22, width = width)

frame_resul_discrep = tk.Frame(tab1)
frame_resul_discrep.place(x=0, y =20,height = (height/1.7)-178, width = width)
pt_resul_discrep = Table(frame_resul_discrep)
pt_resul_discrep.model.df =df
options = {
 'cellbackgr': '#f7f6dc',
 'rowselectedcolor': '#f7f6dc',
 'textcolor': 'black'}
config.apply_options(options, pt_resul_discrep)
pt_resul_discrep.show()
pt_resul_discrep.autoResizeColumns()
pt_resul_discrep.redraw()


######################################
lbl_novas = ttk.Label(tab1, text="LINHAS ADICIONADAS (presentes somente no arquivo novo):",
                          font='Helvetica 12 bold')
#lbl1.configure(text="")
lbl_novas.place(x=0, y =(height/1.7)-178+25,height = 22, width = width)

frame_resul_novas = tk.Frame(tab1)
frame_resul_novas.place(x=0, y =(height/1.7)-178+45,height = (height/2.86)-178, width = width)
pt_resul_novas = Table(frame_resul_novas)
pt_resul_novas.model.df =df
options = {
 'cellbackgr': '#98faa7',
 'colheadercolor': '#16f747',
 'rowselectedcolor': '#98faa7',
 'textcolor': 'black'}
config.apply_options(options, pt_resul_novas)
pt_resul_novas.show()
# pt_resul_novas.setRowColors(rows=0, clr='#00ff08',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_novas.setRowColors(rows=1, clr='#00ff08',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_novas.setRowColors(rows=2, clr='#00ff08',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_novas.setRowColors(rows=3, clr='#00ff08',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_novas.setRowColors(rows=4, clr='#00ff08',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
pt_resul_novas.autoResizeColumns()
pt_resul_novas.redraw()

###########################
lbl_excluidas = ttk.Label(tab1, text="LINHAS EXCLUIDAS (presentes somente no arquivo antigo):",
                          font='Helvetica 12 bold')
#lbl1.configure(text="")
lbl_excluidas.place(x=0, y =((height/1.26)-178),height = 22, width = width)

frame_resul_excluidas = tk.Frame(tab1)
frame_resul_excluidas.place(x=0, y =(height/1.26)-178+25,height = (height/2.86)-178, width = width)
pt_resul_excluidas = Table(frame_resul_excluidas)

pt_resul_excluidas.model.df =df
options = {
 'cellbackgr': '#fa9898',
 'colheadercolor': '#f71616',
 'rowselectedcolor': '#fa9898',
 'textcolor': 'black'}
config.apply_options(options, pt_resul_excluidas)
pt_resul_excluidas.autoResizeColumns()
pt_resul_excluidas.show()
# pt_resul_excluidas.setRowColors(rows=1, clr='#ff0000',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_excluidas.setRowColors(rows=2, clr='#ff0000',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_excluidas.setRowColors(rows=3, clr='#ff0000',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_excluidas.setRowColors(rows=4, clr='#ff0000',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
# pt_resul_excluidas.setRowColors(rows=0, clr='#ff0000',  cols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
pt_resul_excluidas.redraw()


root.mainloop()
