#!/usr/bin/env python3
import pandas as pd
import xlwings as xw
from datetime import date, timedelta
from auxiliares import *


class Capitulo:
    def __init__(self, inicio, nome=None, fim=None) -> None:
        self.inicio = inicio
        self.nome = nome
        self.fim = fim


class gera_calendario:
    colunas = ['C', 'D', 'E', 'F', 'G', 'H', 'I']

    def __init__(self, nome_arquivo, mes, ano, inicio, qtd_dias):
        self.nome_arquivo = nome_arquivo
        self.mes = mes
        self.ano = ano
        self.inicio = inicio
        self.qtd_dias = qtd_dias
        self.domingo, self.espera = self.encontra_domingo()
        self.capitulos = []
        self.le_txt()
        

    def encontra_domingo(self):
        data = date(year=self.ano, month=self.mes, day=self.inicio)

        indice_da_semana = data.isoweekday()%7

        domingo = data - timedelta(days=indice_da_semana)

        return (int(domingo.strftime("%d")), indice_da_semana)
    
    def le_txt(self):
        arquivo = open('capitulos.txt', 'r')

        for linha in arquivo:
            if "->" in linha:
                pass
                
            if ":" in linha:
                separados = linha.split(":")

                inicio = int(separados[0])
                fim = int(separados[1])

                self.capitulos.append(Capitulo(inicio=inicio, fim=fim))
            else:
                self.capitulos.append(Capitulo(inicio= int(linha)))

    def gera_dias(self):
        total_paginas = self.capitulos[-1].fim
        pagina_dia = total_paginas/(self.qtd_dias)
        captiulo_dia = []
        prox_cap = 0
        paginas = self.capitulos[0].inicio
        
        #Divide os capítulos por dia
        for i in range(1, self.qtd_dias+1+self.espera):
            paginas += pagina_dia

            captiulos_atuais = []

            for cap in range(prox_cap, len(self.capitulos)):
                if(paginas >= self.capitulos[cap].inicio):
                    captiulos_atuais.append(str(cap+1))
                    prox_cap = cap
                else:
                    break
            
            captiulo_dia.append(captiulos_atuais)

        #Capítulo com nome
        j = 0
        for capitulos in captiulo_dia:
            captiulo_dia[j] = " e ".join(capitulos)
            j+=1

        return captiulo_dia

    def atualiza_data_calendario(self, ws):
        #Seleciona o dia e o mês
        celula_dia = "R5"
        ws.range(celula_dia).value = self.domingo
        celula_mes = "S5"
        ws.range(celula_mes).value = self.mes
        celula_ano = "M1"
        ws.range(celula_ano).value = self.ano

    def escreve_calendario(self, ws):
        i = 0
        dias = 0
        valores = self.gera_dias()

        linha = 4
        qtd_maxima = self.qtd_dias + self.espera
        while(dias <= qtd_maxima):
            for coluna in self.colunas:
                if(dias > qtd_maxima):
                    break
                if(dias < self.espera):
                    dias+=1
                    continue
                dias+=1
                celula = coluna+str(linha)
                ws.range(celula).value = "Capítulo "+valores[i]
                ws.range(celula).api.Font.Size = 14
                i+=1
            linha += 2


    def escrever_xls(self):

        wb = xw.Book(self.nome_arquivo) # wb = xw.Book(filename) would open an existing file
        ws = wb.sheets[0]

        self.atualiza_data_calendario(ws)

        #Escreve
        self.escreve_calendario(ws)


dia = 4
mes = 10
ano = 2021
qtd_dias = 40
novo_calendario = gera_calendario("./novo_calendario.xlsx", mes, ano, dia, qtd_dias)
#novo_calendario.gera_dias()
novo_calendario.escrever_xls()