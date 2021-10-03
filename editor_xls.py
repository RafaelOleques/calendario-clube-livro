from capitulo import Capitulo
from xls_lib_interface import Xls_lib_interface

class Editor_xls:
    colunas = ['C', 'D', 'E', 'F', 'G', 'H', 'I']

    def __init__(self, nome_xls, nome_txt, cronograma):
        self.nome_xls = nome_xls
        self.cronograma = cronograma
        self.nome_txt = nome_txt
        self.le_txt()    
    
    def le_txt(self):
        arquivo = open(self.nome_txt, 'r')

        for linha in arquivo:
            if "->" in linha:
                pass
            
            #Pagina inicial e final do capítulo
            if ":" in linha:
                separados = linha.split(":")

                inicio = int(separados[0])
                fim = int(separados[1])

                self.cronograma.capitulos.append(Capitulo(inicio=inicio, fim=fim))
            else:
                #Somente a página inicial do capítulo
                self.cronograma.capitulos.append(Capitulo(inicio= int(linha)))

    def atualiza_data_xls(self, xls):
        #Seleciona o dia e o mês
        celula_dia = "R5"
        xls.edit_val_cell(celula_dia, self.cronograma.domingo)
        celula_mes = "S5"
        xls.edit_val_cell(celula_mes, self.cronograma.mes)
        celula_ano = "M1"
        xls.edit_val_cell(celula_ano, self.cronograma.ano)

    def escreve_calendario(self, xls):
        i = 0
        dias = 0
        valores = self.cronograma.gera_cap_dias()

        linha = 4
        qtd_maxima = self.cronograma.qtd_dias + self.cronograma.espera
        while(dias <= qtd_maxima):
            for coluna in self.colunas:
                if(dias > qtd_maxima):
                    break
                celula = coluna+str(linha)
                if(dias < self.cronograma.espera):
                    dias+=1
                    xls.edit_val_cell(celula, "")
                    continue
                dias+=1
                xls.edit_val_cell(celula, valores[i])
                xls.edit_val_cell(celula, valores[i])
                xls.edit_font_size_cell(celula, 14)
                i+=1
            linha += 2

    def escreve_xls(self):
        xls = Xls_lib_interface(self.nome_xls)

        self.atualiza_data_xls(xls)
        self.escreve_calendario(xls)

