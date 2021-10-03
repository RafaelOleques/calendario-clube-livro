from cronograma import Cronograma
from editor_xls import Editor_xls


dia = 11
mes = 10
ano = 2021
qtd_dias = 40

nome_txt = './arquivos/capitulos.txt'
nome_xls = "./arquivos/novo_calendario.xlsx"

cronograma = Cronograma(mes, ano, dia, qtd_dias)
editor = Editor_xls(nome_xls, nome_txt, cronograma)
editor.escreve_xls()