#!/usr/bin/env python3
import pandas as pd
import xlwings as xw
from auxiliares import *


class gera_calendario:
    linhas = [7,9,11,13,15]
    colunas = ['B', 'C', 'D', 'E', 'F', 'G', 'H']

    def __init__(self, nome_arquivo, mes, ano, inicio, fim, extras, nro_capitulo):
        self.nome_arquivo = nome_arquivo
        self.nro_capitulo = nro_capitulo
        self.mes = mes
        self.ano = ano
        self.inicio = inicio
        self.fim = fim
        self.extras = extras


    def escrever_xls(self):

        wb = xw.Book(self.nome_arquivo) # wb = xw.Book(filename) would open an existing file

        ws = wb.sheets["Setembro"]
        
        total_dias = total_dias_mes(self.mes, self.ano) + self.extras

        dias = 0

        primeiro = primeiro_dia_semana(self.mes, self.ano)

        for coluna in self.colunas:
            if(ord(coluna) >= primeiro):
                dias += 1
                if(dias >= self.inicio):
                    celula = coluna+str(5)
                    print(celula)
                    ws.range(celula).value = '2'

        #print(wb.sheets)

        for linha in self.linhas:
            for coluna in self.colunas:
                dias+=1
                if(dias > total_dias):
                    break
                if(dias >= self.inicio):
                    celula = coluna+str(linha)
                    #print(celula)
                    ws.range(celula).value = '2'


dia = 6
mes = 9
ano = 2021
novo_calendario = gera_calendario("./calendario_clube.xlsx", mes, ano, 13, 30, 2, 15)
novo_calendario.escrever_xls()