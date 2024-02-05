#-------------------------------------
#------IMPORTANDO BIBLIOTECAS---------
#-------------------------------------
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import openpyxl
import datetime

#-------------------------------------
#---------DEFININDO VARIAVEIS-----------
#-------------------------------------
nome_arquivo1 = 'rhp.xlsx'
nome_arquivo2 = 'silvestre.xlsx'
nome_arquivo3 = 'real.xlsx'
version = '1.0'

#-------------------------------------
#---------DEFININDO FUNÇÕES-----------
#-------------------------------------

def button_refresh():
    if os.path.isfile(nome_arquivo1) and os.path.isfile(nome_arquivo2):
        print(f'O arquivo existe na pasta do projeto.')
    else:
        messagebox.showerror("Error", f"O arquivo {nome_arquivo1} e ou {nome_arquivo2} não existe!!!\nExporte os relatórios do Wsac para continuar.")

def coluna_para_indice(coluna):
    return ord(coluna.upper()) - ord('A')

def entrada():
    #VERIFICANDO SE OS ARQUIVOS EXISTEM
    if os.path.isfile(nome_arquivo1) and os.path.isfile(nome_arquivo2):
        #ABRINDO UMA PLANILHA DO EXCEL
        planilha = openpyxl.load_workbook(nome_arquivo1)
        #SELECIONANDO A PRIMEIRA ABA DA PLANILHA
        aba = planilha.active
        #PERCORRENDO TODAS AS LINHAS DA PLANILHA, COMEÇANDO PELA SEGUNDA LINHA
        for row in aba.iter_rows(min_row=2):
            #OBTENDO O VALOR DA CÉLULA NA COLUNA 2
            valor = row[1].value
            #VERIFICA SE O VALOR É MENOR DO QUE ZERO
            if valor < 0:
                #DEFINE O VALOR PARA 0
                row[1].value = 0
        #DELETA A PLANILHA ANTIGA
        os.remove(nome_arquivo1)
        #SALVA A PLANILHA ATUALIZADA E CRIA UMA NOVA PLANILHA PARA USAR DE BASE
        planilha.save(nome_arquivo1)
        planilha.save(nome_arquivo3)
        
        #zREPETINDO OS PASSOS COM A PLANILHA 2
        planilha = openpyxl.load_workbook(nome_arquivo2)
        aba = planilha.active
        for row in aba.iter_rows(min_row=2):
            valor = row[2].value
            if valor < 0:
                row[2].value = 0
        os.remove(nome_arquivo2)
        planilha.save(nome_arquivo2)
#----------------------------------------------------------------------------------
#---PLANILHAS ZERADAS, VAMOS IGUALAR OS VALORES------------------------------------
#----------------------------------------------------------------------------------
        #LENDO OS DADOS DAS PLANILHAS
        real_workbook = openpyxl.load_workbook(nome_arquivo3)
        silvestre_workbook = openpyxl.load_workbook(nome_arquivo2)

        #CRIANDO UMA NOVA PLANILHA - QUE RECEBERA OS DADOS 
        nova_planilha = openpyxl.Workbook()
        nova_planilha_ativa = nova_planilha.active

        #OBTENDO AS FOLHAS ATIVAS
        estoque_real_sheet = real_workbook.active
        silvestre_sheet = silvestre_workbook.active

        #OBTENDO O NUMERO MAXIMO DE LINHAS
        num_linhas_silvestre = silvestre_sheet.max_row

        #PERCORRENDO CADA LINHA DA PLANILHA
        for linha_silvestre in range(2, num_linhas_silvestre + 1):
            codigo_silvestre = silvestre_sheet.cell(row=linha_silvestre, column=coluna_para_indice('B') + 1).value
            #VERIFICA SE O CAMPO NÃO ESTA VAZIO 
            if codigo_silvestre is not None:

                #PROCURANDO O CÓDIGO NA COLUNA A
                for linha_estoque_real in range(1, estoque_real_sheet.max_row + 1):
                    codigo_estoque_real = estoque_real_sheet.cell(row=linha_estoque_real, column=coluna_para_indice('A') + 1).value

                    #VERIFICA SE O CODIGO DA PLANILHA É IGUAL O DA OUTRA
                    if codigo_silvestre == codigo_estoque_real:
                        #OBTENDO OS VALORES DA CELULA A DIREITA DOS CODIGOS "CODIGO | VALOR"
                        valor_estoque_real = estoque_real_sheet.cell(row=linha_estoque_real, column=coluna_para_indice('B') + 1).value
                        valor_cod_silvestre = silvestre_sheet.cell(row=linha_silvestre, column=coluna_para_indice('A') + 1).value

                        #PREENCHENDO A NOVA PLANILHA COM ESSES VALORES
                        nova_planilha_ativa.append([valor_cod_silvestre, valor_estoque_real])

        #DELETA A PLANILHA ANTIGA
        os.remove(nome_arquivo2)
        #SALVA A PLANILHA NOVA
        nova_planilha.save(nome_arquivo2)
#----------------------------------------------------------------------------------
#---CRIANDO ARQUIVO PARA IMPORTAR OS NOVOS DADOS-----------------------------------
#----------------------------------------------------------------------------------
        #CARREGANDO A PLANILHA
        workbook1 = openpyxl.load_workbook(nome_arquivo2)

        #SELECIONA PLANILHA
        sheet1 = workbook1.active
        #OBTENDO DATA ATUAL
        data_atual = datetime.datetime.now().strftime("%d-%m-%Y")

        #ABRE O ARQUIVO DE TEXTO PARA COLOCARMOS OS DADOS
        with open(f'estoque_silvestre {data_atual}.txt', 'w') as txt_file2:
            
            #ESCREVE OS CABEÇALHOS DAS COLUNAS NO ARQUIVO DE TEXTO
            headers2 = [coluna.value for coluna in sheet1[1]]
            txt_file2.write(';'.join(map(str, headers2)) + '\n')

            #ITERA SOBRE AS LINHAS DA PLANILHA (começando da segunda linha, já que a primeira contém os cabeçalhos)
            for row in sheet1.iter_rows(min_row=2, values_only=True):
                #ESCREVE OS VALORES DE CADA COLUNA NO ARQUIVO DE TEXTO, SEPARANDOS POR ";"
                txt_file2.write(';'.join(map(str, row)) + '\n')
        
        #DELETA AS PLANILHAS ANTIGAS
        os.remove(nome_arquivo1)
        os.remove(nome_arquivo2)
    else:
        messagebox.showerror("Error", f"O arquivo {nome_arquivo1} e ou {nome_arquivo2} não existe!!!\nExporte os relatórios do Wsac para continuar.")

def button_entrada():
    entrada()
    
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
janela.title(f"Reajuste - Wsac {version}")
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