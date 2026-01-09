import gspread
import pandas as pd
import random
import datetime
import openpyxl
import tkinter as tk
from tkinter import OptionMenu
from tkinter import Tk, StringVar
from tkinter import *
from tkinter import ttk

#Definindo variaveis
consultor = ""
qtd_clientes = ""
vendedores = ["VENDEDOR1", "VENDEDOR2", "VENDEDOR3", "VENDEDOR4", "VENDEDOR5"]

lista_vendedores_com_leads = []
# Define a coluna "VALIDADE DO LEAD"
coluna_validade_lead = "VALIDADE DO LEAD"
# Define a coluna "ATUAL"
coluna_atual = "ATUAL"
# Define a coluna "ATUAL"
coluna_lead = "LEAD"
# Definindo a conluna CNPJ
coluna_cnpj = "CNPJ"
# Definindo as colunas de origem:                                                                                                                                                                                        
colunas_origem = ["ESTRATÉGIA", "QUALIFICAÇÃO", "CNPJ", "RAZÃO SOCIAL", "ENRIQUECIMENTO", "SERVIÇOS ATUAIS"]

# Lista Planilhas de destino:
planilhas_destino = {
    "VENDEDOR1": "https://docs.google.com/spreadsheets/d/1Q0jLzU9uXiJ3X-VyLdS88tcwTbuDC6bKQxh_GUERaEU/edit?gid=0#gid=0",
    "VENDEDOR2": "https://docs.google.com/spreadsheets/d/1WGqsu3poRadEnS9oSPkl6IKXznYaZjuvdEYhFHOD5hQ/edit?gid=0#gid=0",
    "VENDEDOR3": "https://docs.google.com/spreadsheets/d/1Rgtc_vUQ9Ic2vSZp_ySbScpmRWjCs_v_thmMwViyb-s/edit?gid=0#gid=0",
    "VENDEDOR4": "https://docs.google.com/spreadsheets/d/1QQLvhyrmch4TGH2luOzAO65SD3exxG014Ilo8YirXj0/edit?gid=0#gid=0",
    "VENDEDOR5": "https://docs.google.com/spreadsheets/d/1059Es5L0JPtdbWfuq0ZR5vGDbXxcYyzwzWOYAZyc2Dc/edit?gid=0#gid=0"
    }
    
def copiar(): 
    # Apresentando credencial:
    gc = gspread.service_account("PROJETO AUTOMATIZAÇÃO INTERFACE/CREDENCIAL/project-f5-411718-d5ecbc1f2057.json")
    # Abrindo a planilha
    planilha_origem = gc.open_by_url('https://docs.google.com/spreadsheets/d/1eUFNbZkO-1UrmWgxaqFWfLIGMXn9XUYaXQJPv5EC1qE/edit?gid=0#gid=0')
     # Obtendo a aba 0 da spreadsheet. https://docs.google.com/spreadsheets/d/1copGvgG5O50xgLZRiFYa9g6cCbm8qkS0fOLaSjwJub8/edit#gid=1602287841&fvid=299385111
    aba = planilha_origem.get_worksheet(0) 
    
    link_planilha = link_planilha_selecionado.get()
    consultor = (link_planilha)
    print(consultor)
    qtd_clientes = int(entry_qtd_clientes.get())

    # Abrindo a planilha de destino
    gc = gspread.service_account("PROJETO AUTOMATIZAÇÃO INTERFACE/CREDENCIAL/project-f5-411718-d5ecbc1f2057.json")
 
    # Obterndo link a partir do nome do consultor
    url_planilha_destino = planilhas_destino[consultor]
    
    planilha_dest = gc.open_by_url(url_planilha_destino)
     # Obtendo a aba 0 da spreadsheet.
    aba_dest = planilha_dest.worksheet("MIG VOZ 1P - REC") 
    
    # Obtém a linha de cabeçalho
    linha_cabecalho = aba.row_values(1)

    # Obtém os índices das colunas
    indice_coluna_validade_lead = linha_cabecalho.index(coluna_validade_lead)
    indice_coluna_atual = linha_cabecalho.index(coluna_atual)
    indice_coluna_lead = linha_cabecalho.index(coluna_lead)



    # Filtra as linhas por "DISPONÍVEL" na coluna "VALIDADE DO LEAD" e "consultor" na coluna "ATUAL" - USAR / APÓS TERMINAR A CONSULTA VIABILIDADE
    #linhas_filtradas = list(filter(lambda linha: linha[indice_coluna_validade_lead] == "DISPONIVEL" and linha[indice_coluna_atual] != consultor and linha[indice_coluna_lead] in ("BL", "HIBRIDO/END DIFERENTE"), aba.get_all_values()[1:]))  # Exclui a linha de cabeçalho
    linhas_filtradas = list(filter(lambda linha: linha[indice_coluna_validade_lead] == "DISPONIVEL" and linha[indice_coluna_atual] != consultor and linha[indice_coluna_lead] in ("BL", "HIBRIDO/END DIFERENTE"), aba.get_all_values()[1:]))  # Exclui a linha de cabeçalho

    # Mistura as linhas aleatoriamente
    random.shuffle(linhas_filtradas)

    # Seleciona os qtd_clientes primeiros clientes
    clientes_selecionados = linhas_filtradas[:qtd_clientes]
    
    # Obtém os índices das colunas desejadas
    indices_colunas = [linha_cabecalho.index(coluna) for coluna in colunas_origem]

    # Extrai os dados das colunas desejadas das linhas filtradas
    dados_filtrados = []
    for linha in clientes_selecionados:
        dados_linha = []
        for indice_coluna in indices_colunas:
            valor_celula = linha[indice_coluna]
            dados_linha.append(valor_celula)
        dados_filtrados.append(dados_linha)
    
    #Transofrma a lista em df
    df_dados_filtrados = pd.DataFrame(dados_filtrados, columns=colunas_origem)
    
    #insere uma coluna na primeira posição do df com o nome DATA e a data atual como seu conteúdo
    df_dados_filtrados.insert(0, "VALOR ATUAL", None)
    df_dados_filtrados.insert(0, "VALOR NOVO", None)
    df_dados_filtrados.insert(0, "NV PLANO", None)    
    df_dados_filtrados.insert(7, "VALIDADE DO LEAD", None)
    df_dados_filtrados.insert(7, "COMENTÁRIOS", None)
    df_dados_filtrados.insert(7, "PROX CONTATO", None)
    df_dados_filtrados.insert(11, 'DATA', datetime.datetime.today().strftime('%d/%m/%Y'))
    df_dados_filtrados.to_excel("ESTRUTURA.xlsx")
    
    #Cola as informações na aba_dest    
    aba_dest.append_rows(df_dados_filtrados.values.tolist())
    
    def atualizar():
        # Apresentando credencial:
        gc = gspread.service_account('PROJETO AUTOMATIZAÇÃO INTERFACE/CREDENCIAL/project-f5-411718-d5ecbc1f2057.json')
        #planilha de origem
        sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1eUFNbZkO-1UrmWgxaqFWfLIGMXn9XUYaXQJPv5EC1qE/edit?gid=0#gid=0')
        #Obtendo a aba 0 da spreadsheet.
        worksheet = sh.get_worksheet(0) 
        # Obtendo todos os dados da planilha, incluindo cabeçalhos
        all_data = worksheet.get_all_values()

        # Criando o DataFrame
        df2 = pd.DataFrame(all_data)

        # Atualizar o valor da coluna 12 para as linhas que possuem CNPJs presentes no df
        df2.loc[df2.iloc[:, 2].isin(df_dados_filtrados.iloc[:, 5]), 7] = consultor 
        df2.loc[df2.iloc[:, 2].isin(df_dados_filtrados.iloc[:, 5]), 8]= "EM_AÇÃO"
        df2.to_excel("teste.xlsx")
        # Atualizar a planilha Google Sheets
        worksheet.update(range_name='', values=df2.values.tolist())
        print("Compartilhamento concluido!")
        lista_vendedores_com_leads.append(consultor)
        # Atualizar a listbox
        lista_vendedores.delete(0, END)
        for vendedor in lista_vendedores_com_leads:
            lista_vendedores.insert(END, vendedor)
        return 
    atualizar()
    return df_dados_filtrados


#criando interface:
#Janela principal:
janela = tk.Tk()
estilo = ttk.Style()
estilo.configure("meu_estilo", background="#4B0082")
# Definindo a cor de fundo da janela
janela.option_add("*Frame.style", "meu_estilo")
janela.geometry("400x400")  # Define o tamanho da janela (largura x altura)
janela.title("Compartilhamento de Leads")


#Criando frame principal
frame = tk.Frame(janela)
frame.option_add("*Frame.style", "meu_estilo")
frame.pack(expand=YES, fill=BOTH)
#Criando frame para o primeiro label
frame_label1 = tk.Frame(frame)
frame_label1.pack(side=TOP, expand=YES)
#Criando frame para o segundo label
frame_label2 = tk.Frame(frame)
frame_label2.pack(expand=YES)
#Criando frame para o botão:
frame_botao = tk.Frame(frame)
frame_botao.pack(expand=YES)

# Label "Vendedor"
label_vendedor = tk.Label(frame_label1, text="Selecione o consultor(a):")
label_vendedor.config(width=20, font="ArialBlack")
label_vendedor.pack(side=LEFT)
# Cria o OptionMenu (VENDEDOR)
link_planilha_selecionado = tk.StringVar()
link_planilha_selecionado.set("Vendedor 1")
menu_vendedores = OptionMenu(frame_label1, link_planilha_selecionado, *planilhas_destino.keys())
menu_vendedores.config(width=10)
menu_vendedores.pack(side=RIGHT)

# Criando o label "Quantidade de clientes"
rotulo_qtd_cli = tk.Label(frame_label2, text="Quantidade de clientes:")
rotulo_qtd_cli.config(width=20, font="ArialBlack")
rotulo_qtd_cli.pack(side=LEFT)
# Criando o entry para a quantidade de clientes
entry_qtd_clientes = tk.Entry(frame_label2, textvariable=qtd_clientes)
entry_qtd_clientes.config(width=10  )
entry_qtd_clientes.pack(side=RIGHT, padx=10)

# Criando o botão para copiar os dados
botao_copiar = tk.Button(frame_botao, text="COMPARTILHAR", command=copiar)
botao_copiar.config(font="System", pady=5)
botao_copiar.grid()
#Validadando operação:

    
frame_msg = tk.Frame(janela)
frame_msg.pack(side=BOTTOM)

lista_vendedores = tk.Listbox(frame_msg)
lista_vendedores.config(width=400, height=300, font="System")
lista_vendedores.pack(pady=50, padx=100) 

# Executa a janela principal
janela.mainloop()


