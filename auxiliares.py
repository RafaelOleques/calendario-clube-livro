from datetime import datetime
import calendar

def primeiro_dia_semana(mes, ano):
    #Pega o dia da semana
    dia = datetime(ano, mes, 1).weekday() + 2

    c = ord("A") + dia
    return(c)

def total_dias_mes(mes, ano):
    _, num_dias = calendar.monthrange(ano,mes)
    return(num_dias)

