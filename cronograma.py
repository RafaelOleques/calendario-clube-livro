from datetime import date, timedelta

class Cronograma:
    def __init__(self, mes, ano, inicio, qtd_dias):
        self.mes = mes
        self.ano = ano
        self.inicio = inicio
        self.qtd_dias = qtd_dias
        self.capitulos = []
        self.domingo, self.espera = self.encontra_domingo()

    def encontra_domingo(self):
        data = date(year=self.ano, month=self.mes, day=self.inicio)

        indice_da_semana = data.isoweekday()%7
        domingo = data - timedelta(days=indice_da_semana)

        return (int(domingo.strftime("%d")), indice_da_semana)
    
    def cap_por_dia(self):
        prox_cap = 0
        total_paginas = self.capitulos[-1].fim
        paginas = self.capitulos[0].inicio
        pagina_dia = total_paginas/(self.qtd_dias)
        capitulo_dia = []

        for i in range(1, self.qtd_dias+1+self.espera):
            paginas += pagina_dia

            captiulos_atuais = []

            for cap in range(prox_cap, len(self.capitulos)):
                if(paginas >= self.capitulos[cap].inicio):
                    captiulos_atuais.append(str(cap+1))
                    prox_cap = cap
                else:
                    break
            
            capitulo_dia.append(captiulos_atuais)
        
        return capitulo_dia

    def gera_cap_dias(self):
        #Divide os capítulos por dia
        capitulo_dia = self.cap_por_dia()
        string_capitulos = []

        #String para cada capítulo
        j = 0
        for capitulos in capitulo_dia:
            string_capitulos.append(", ".join(capitulos))
            
            ultima_virgula = string_capitulos[j].rfind(",")

            if(ultima_virgula != -1):
                str_aux = string_capitulos[j].rsplit(",", 1)
                string_capitulos[j] = ' e'.join(str_aux)
            
            if(ultima_virgula != -1):
                string_capitulos[j] = "Capítulos "+string_capitulos[j]
            else:
                string_capitulos[j] = "Capítulo "+string_capitulos[j]

            j+=1

        return string_capitulos