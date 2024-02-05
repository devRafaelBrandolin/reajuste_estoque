#-------------------------------------
#------IMPORTANDO BIBLIOTECAS---------
#-------------------------------------
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess


#-------------------------------------
#---------DEFININDO FUNÇÕES-----------
#-------------------------------------

def button_refresh():
    pass

def button_entrada():
    pass
    
def button_saida():
    pass

def sair():
    janela.destroy()

def sobre():
    tk.messagebox.showinfo("Sobre", "Esta aplicação foi desenvolvida para suprir a necessidade de manter atualizado o estoque da empresa.\nO objetivo é trabalhar com as planilhas de relatorio dos estoques e gerenciar os dados para importar um novo relatorio com o estoque atualizado!\nPara saber mais acesse o meu GitHub.")

def github():
    import webbrowser
    webbrowser.open("https://github.com/devRafaelBrandolin")

def abrir_pasta_projeto():
    try:
        subprocess.Popen(['explorer', '.'])  # No Windows
    except OSError:
        try:
            subprocess.Popen(['xdg-open', '.'])  # No Linux
        except OSError:
            tk.messagebox.showerror("Erro", "Não foi possível abrir a pasta do projeto.")

#-------------------------------------
#---------CRIANDO JANELA--------------
#-------------------------------------
janela = tk.Tk()
#DEFININDO UM TITULO PARA JANELA
janela.title("Reajuste - Wsac")
#DEFININDO O TAMANHO DA JANELA
janela.geometry('300x200')
#IMPEDINDO QUE A JANELA SEJE REDIMENSIONÁVEL
janela.resizable(width=False, height=False)

#-------------------------------------
#-----CRIANDO WIDGETS DA JANELA-------
#-------------------------------------
#CRIANDO UM MENU PARA A JANELA
# Menu
barra_menu = tk.Menu(janela)
janela.config(menu=barra_menu)
# Menu "Arquivo"
menu_arquivo = tk.Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Arquivo", menu=menu_arquivo)
menu_arquivo.add_command(label="Abrir Pasta", command=abrir_pasta_projeto)
menu_arquivo.add_command(label="Sair", command=sair)
# Menu "Ajuda"
menu_ajuda = tk.Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Ajuda", menu=menu_ajuda)
menu_ajuda.add_command(label="Sobre", command=sobre)
menu_ajuda.add_command(label="Github", command=github)

#CRIANDO UM CABEÇALHO
#DEFININDO A COR DE FUNDO DO CABEÇALHO
cabecalho_frame = tk.Frame(janela, bg="#4CAF50")
#ESTENDENDO O CABEÇALHO À LARGURA DA JANELA
cabecalho_frame.grid(row=0, column=0, pady=10, columnspan=3, sticky="nsew")
#CRIANDO UM TITULO PARA O CABEÇALHO
label_criativo = tk.Label(cabecalho_frame, text="Bem-vindo ao REAJUSTE!", font=("Arial", 16, "bold"), fg="white", bg="#4CAF50")
label_criativo.grid(row=0, column=0, padx=10, pady=10)

#-------------------------------------
#------CRIANDO BOTÕES DA JANELA-------
#-------------------------------------
botao1 = tk.Button(janela, text="REFRESH", font=("Arial", 10, "bold"),  width='10', command=button_refresh)
botao1.grid(row=1, column=0, pady=5)

botao2 = tk.Button(janela, text="Entrada", font=("Arial", 10, "bold"), width='10', command=button_entrada)
botao2.grid(row=1, column=1, pady=5)

botao3 = tk.Button(janela, text="Saída", font=("Arial", 10, "bold"), width='10', command=button_saida)
botao3.grid(row=1, column=2, pady=5)

botao6 = tk.Button(janela, text="SAIR", font=("Arial", 10, "bold"), width='10', command=sair)
botao6.grid(row=2, column=2, pady=8)

#-------------------------------------
# Rodapé
rodape_frame = tk.Frame(janela)
rodape_frame.grid(row=3, column=0, pady=10, columnspan=3)

rodape_label = tk.Label(rodape_frame, text="Desenvolvido por: Rafael Brandolin")
rodape_label.grid(row=0, column=0, pady=10)

janela.mainloop()