from datetime import date
import datetime

DIA = 29
MES = 9
ANO = 2021 
data = date(year=ANO, month=MES, day=DIA)
print(data)

indice_da_semana = data.isoweekday()%7
print(indice_da_semana)

domingo = data - datetime.timedelta(days=indice_da_semana)
print(domingo.strftime("%d"))