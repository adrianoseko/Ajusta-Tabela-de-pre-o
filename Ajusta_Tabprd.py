
import os
import sys
import pandas as pd
import pyodbc 
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

bc = sys.argv[1]
cl = sys.argv[3]
us = sys.argv[2]
coligada = cl.split('/c:')
banco = bc.split('/d:')
user = us.split('/u:')
drive = "{SQL Server}"
 

root = Tk()
root.geometry("800x600")
root.title("CGA.NET - Tabela Preços")
cnxn = pyodbc.connect(f"DRIVER={drive};SERVER= Seu Servidor;DATABASE={banco[1]};UID=login;PWD=senha;")

sql = f"""SELECT ZTC.IDTABPRECO, NOME from ZA_TTABPRECO ZTC 
LEFT JOIN TTABPRECO TTP (NOLOCK)	        ON TTP.CODCOLIGADA = ZTC.CODCOLIGADA AND TTP.IDTABPRECO = ZTC.IDTABPRECO 
where USADEFAULTABELA = 'N' and 
TTP.IDTABPRECO > 3 AND TTP.ATIVA = 1 AND 
CONVERT(VARCHAR(10) , GETDATE() , 126) >= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAINI , 126) AND
CONVERT(VARCHAR(10) , GETDATE() , 126) <= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAFIM , 126)
"""

#select * from ZA_TABCOMISSAO WHERE idtabcomis = 16 

cursor = cnxn.cursor()
tb = pd.read_sql(sql, cnxn)
nome = tb.NOME.values
idt = tb.IDTABPRECO.values
var = StringVar(root)
var.set(nome[0])
menu = OptionMenu(root, var,*nome)
menu.config(width=30, font=('Helvetica', 12))
menu.place(x=20,y=10)

    
def callback(*args):
    tabela = str(var.get())
    table.delete(*table.get_children())
    global index
    index=0
    for i in tabela:
        if tabela == nome[index]:
            break
        index = index+1
    global idtab
    idtab = idt[index] 
    sql2 = f"""SELECT  IDTABPRECO, ZTC.IDPRD , CODIGOPRD, NOMEFANTASIA, PRECO, CUSTO, MARGEM, ADIC_FINANC  from ZA_TTABPRECOITM ZTC
               LEFT JOIN TPRD (NOLOCK)			        ON TPRD.CODCOLIGADA = ZTC.CODCOLIGADA AND TPRD.IDPRD = ZTC.IDPRD 
               where IDTABPRECO = {idtab} ORDER BY NOMEFANTASIA""" 
    
    global tbgrid
    tbgrid = pd.read_sql(sql2, cnxn)
    cod = tbgrid.CODIGOPRD.values
    nomef = tbgrid.NOMEFANTASIA.values 
    preco = tbgrid.PRECO.values
    custo = tbgrid.CUSTO.values
    margem = tbgrid.MARGEM.values
    adc = tbgrid.ADIC_FINANC.values
    idprd = tbgrid.IDPRD.values	
    ind = 0
    for n in cod:
        table.insert('' ,END, values=(idprd[ind], cod[ind], nomef[ind],preco[ind],custo[ind],margem[ind],adc[ind]), tag=f"{ind}")
        ind=ind+1
    filtro_var.trace("w", filtro)
    table.place(x=20,y=50)
    button3.place(x=700, y=500)
    button4.place(x=700, y=550)
    verscrlbar.place(x=783, y=50, height=425)
    fltr.place(x=450, y=15)
    ft.place(x=390, y=15)
    
def selectItem(self):
    
    button4.place_forget()
    button3.place_forget()
    menu.place_forget()
    ft.place_forget()
    fltr.place_forget()
    
    verscrlbar.place(x=783, y=50, height=425, command=None)
    table.configure(yscrollcommand=None)
    
    codp['state'] = NORMAL
    codpl.place(x=50, y=480)
    codp.place(x=50, y=500)
    
    nomep['state'] = NORMAL
    nomepl.place(x=200, y=480)
    nomep.place(x=200, y=500)
    
    precol.place(x=50, y=530)
    E1.place(x=50, y=550)
    
    cust.place(x=200, y=530)
    E2.place(x=200, y=550)
    
    mrg.place(x=350, y=530)
    E3.place(x=350, y=550)
    
    adcf.place(x=500, y=530)
    E4.place(x=500, y=550)
    
    button.place(x=700, y=530)
    button2.place(x=700, y=560)
    
    
    itemselc = table.selection()[0]
    item = table.item(itemselc, "values")

    
    codp.delete(0,"end")
    codp.insert(0, item[1])
    codp['state'] = DISABLED
    
    nomep.delete(0,"end")
    nomep.insert(0, item[2])
    nomep['state'] = DISABLED
                   
    E1.delete(0,"end")
    E1.insert(0, item[3])
    
    E2.delete(0,"end")
    E2.insert(0, item[4])
    
    E3.delete(0,"end")
    E3.insert(0, item[5])
    
    E4.delete(0,"end")
    E4.insert(0, item[6])
    table.state(("disabled",))
    table.bind('<Button-1>', lambda e: 'break') 
    
    

def filtro(*args):    

    try:
        itens = table.get_children()
        item = filtro_var.get().upper()

        for n in itens:
            if item in table.item(n)['values'][2]:
                item_achado = table.item(n)['values']
                table.delete(n)
                table.insert("", 0 ,  values=item_achado)
        
    
    except:
        messagebox.showerror("ERRO!", """"Não foi possível obter filtro!
Entre em contato com o suporte!""")
    

table = ttk.Treeview(root, selectmode = 'extended', 
                     column=('Column1','Column2','Column3','Column4','Column5','Column6','Column7'), 
                     show='headings',  height=20)


verscrlbar = ttk.Scrollbar(root, orient ="vertical",command=table.yview) 
                           
table.configure(yscrollcommand = verscrlbar.set)                      

table.column('Column1',width=80, minwidth=50, stretch=NO)
table.heading("#1", text="ID Produto")

table.column('Column2',width=100, minwidth=50, stretch=NO)
table.heading("#2", text="Código Produto")

table.column('Column3',width=300, minwidth=100, stretch=NO)
table.heading("#3", text="Nome Produto")

table.column('Column4',width=60, minwidth=50, stretch=NO)
table.heading("#4", text="Preço")

table.column('Column5',width=60, minwidth=50, stretch=NO)
table.heading("#5", text="Custo")

table.column('Column6',width=60, minwidth=50, stretch=NO)
table.heading("#6", text="Margem")

table.column('Column7',width=100, minwidth=50, stretch=NO)
table.heading("#7", text="Adcional Financeiro")

table.place_forget()

table.bind ( "<Double-1>" , selectItem)

def record():
        
    try:
        
        itemselc = table.selection()[0]
        idp = table.item(itemselc, "values")
        idprd = idp[0]
        preco = E1.get()
        custo = E2.get()
        margem = E3.get()
        adc = E4.get()
        E1.delete(0,"end")
        E2.delete(0,"end")
        E3.delete(0,"end")
        E4.delete(0,"end")
            
        
        update = f"""UPDATE ZA_TTABPRECOITM SET PRECO = {preco}, CUSTO = {custo}, MARGEM = {margem}, ADIC_FINANC = {adc}
                   WHERE CODCOLIGADA = 5 AND IDPRD = {idprd} and IDTABPRECO = {idtab}"""
        cursor.execute(update)
        cursor.commit()
       
        codp.place_forget()
        nomep.place_forget()
        
        precol.place_forget()
        E1.place_forget()
        
        cust.place_forget()
        E2.place_forget()
        
        mrg.place_forget()
        E3.place_forget()
        
        adcf.place_forget()
        E4.place_forget()
        
        verscrlbar.place_forget()
        
        button.place_forget()
        button2.place_forget()
        messagebox.showinfo("Sucesso!", "Alterações salvas com sucesso!")
        callback()
    except:
        messagebox.showerror("Erro Sistema", """"Não foi possível efetuar as alterações \b
                             Entre em contato com o suporte!""")

def exportar():
    try:
        tbgrid.to_excel(f"{nome[index]}.xlsx", index=False)
        messagebox.showinfo("Exportação Excel", "Tabela exportada com sucesso!")
        os.startfile(f"C:/Users/admin/Desktop/Adriano/Projetos_SW/TABELA DE PREÇO/{nome[index]}.xlsx")  
    except:
        messagebox.showerror("Exportação Excel", """"Não foi possível exportar a tabela \b
                             Entre em contato com o suporte!""")
def cancela():
    try:
        
        button3.place(x=700, y=500)
        button4.place(x=700, y=550)
        verscrlbar.place(x=783, y=50, height=425)
        
        fltr.place(x=450, y=15)
        ft.place(x=390, y=15)
        
        codpl.place_forget()
        codp.place_forget()
        
        nomepl.place_forget()
        nomep.place_forget()
        
        precol.place_forget()
        E1.place_forget()
        
        cust.place_forget()
        E2.place_forget()
        
        mrg.place_forget()
        E3.place_forget()
        
        adcf.place_forget()
        E4.place_forget()
        
        button.place_forget()
        button2.place_forget()
    
        codp.delete(0,"end")
        nomep.delete(0,"end")
        E1.delete(0,"end")
        E2.delete(0,"end")
        E3.delete(0,"end")
        E4.delete(0,"end")
    
        table.state(("!disabled",))
        table.unbind('<Button-1>')
      
    
    except:
        messagebox.showerror("Erro!", "Entre em contato com o suporte!")


def cancela2():
    menu.place(x=20,y=10)
    table.delete(*table.get_children())
    table.place_forget()
    button3.place_forget()
    button4.place_forget()
    verscrlbar.place_forget()
    fltr.place_forget()
    ft.place_forget()
    
button = Button(root, text="Gravar", command=record, height=1, width=10)
button.pack_forget()

button2 = Button(root, text="Cancelar", command=cancela, height=1, width=10)
button2.pack_forget()

button3 = Button(root, text="Exportar", command=exportar, height=1, width=10)
button2.pack_forget()

button4 = Button(root, text="Cancelar", command=cancela2, height=1, width=10)
button4.pack_forget()

ft = Label(root, text="Pesquisar:")
ft.place_forget()
filtro_var = StringVar()
fltr= Entry(textvariable=filtro_var, width=50)
fltr.place_forget()
codpl = Label(root, text="Código do Produto")
codpl.pack_forget()
codp = Entry(bd =5)
codp.pack_forget()

nomepl = Label(root, text="Nome do Produto")
nomepl.pack_forget()
nomep = Entry(bd =5, width=70)
nomep.pack_forget()

precol = Label(root, text="Preço")
precol.pack_forget()
E1 = Entry(bd =5)
E1.pack_forget()

cust = Label(root, text="Custo")
cust.pack_forget()
E2 = Entry(bd =5)
E2.place_forget()

mrg = Label(root, text="Margem")
mrg.pack_forget()
E3 = Entry(bd =5)
E3.place_forget()

adcf = Label(root, text="Adicional Financeiro")
adcf.pack_forget()
E4 = Entry(bd =5)
E4.place_forget()

label1 = Label(root, text=f"Banco: {banco[1]}")
label1.place(x=50, y=580)

label2 = Label(root, text=f"Coligada: {coligada[1]}")
label2.place(x=200, y=580)

label3 = Label(root, text=f"Usuario: {user[1]}")
label3.place(x=280, y=580)

label3 = Label(root, text=f"Versão: 1.0")
label3.place(x=500, y=580)



var.trace("w", callback)

root.mainloop()
