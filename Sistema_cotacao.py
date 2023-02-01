import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
from datetime import datetime
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())

def extrair_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_textocotacao['text'] = f"A Cotação do {moeda} no dia {data_cotacao} foi de: R${valor_moeda}"

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'

def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_cotacao = calendario_moeda.get()

        ano = data_cotacao[-4:]
        mes = data_cotacao[3:5]
        dia = data_cotacao[:2]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel("Teste.xlsx")
        label_atualizarcotacoes['text'] = "Arquivo Atualizado com Sucesso"
    except:
        label_atualizarcotacoes['text'] = "Selecione um arquivo Excel no Formato Correto"

janela = tk.Tk()

janela.title("Sistema de Cotação de Moedas")

label_cotacaomoeda = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionarmoeda = tk.Label(text="Selecionar Moeda", anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_selecionardata = tk.Label(text="Selecionar a data que deseja extrair a cotação", anchor='e')
label_selecionardata.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moeda = DateEntry(year=2023, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nswe',  columnspan=2)

botao_extraircotacao = tk.Button(text="Extrair Cotação", command=extrair_cotacao)
botao_extraircotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')

label_cotacaomultiplasmoeda = tk.Label(text="Cotação Múltiplas Moedas", borderwidth=2, relief='solid')
label_cotacaomultiplasmoeda.grid(row=4, column=0, padx=10, pady=10, sticky='nswe',  columnspan=3)

label_selecionararquivo = tk.Label(text="Selecione um arquivo em Exel com as Moedas na Coluna A")
label_selecionararquivo.grid(row=5, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text="Clique para Selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')

label_arquivoselecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
label_arquivoselecionado.grid(row=6, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_data = tk.Label(text="Data da Cotação", anchor='e')
label_data.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')

calendario_data = DateEntry(year=2023, locale='pt_br')
calendario_data.grid(row=7, column=2, padx=10, pady=10, sticky='nsew')

botao_atualizarcotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')

label_atualizarcotacoes = tk.Label(text="")
label_atualizarcotacoes.grid(row=8, column=1, padx=10, pady=10, sticky='nswe', columnspan=2)

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=9, column=2, padx=10, pady=10, sticky='nswe')

janela.mainloop()