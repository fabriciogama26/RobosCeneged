import pandas as pd

# Caminho do arquivo Excel
file_path = "C:\\Users\\fabriciogama\\OneDrive - CENEGED - COMPANHIA ELETROMECANICA E GERENCIAMENTO DE DADOS\\Leonardo\\Relatório de Excecução Manutenção.xlsx"
file2_path = "C:\\Users\\fabriciogama\\Downloads\\PLANILHA CONSOLIDADO MEDIÇÃO.xlsx"

# Ler o arquivo Excel
data1 = pd.read_excel(file_path, skiprows=1)  # Começa a partir da segunda linha
data2 = pd.read_excel(file2_path)

# Forçar a conversão das colunas 'Hora-Inicio' e 'Hora-Fim' para datetime
data1['Hora-Inicio'] = pd.to_datetime(data1['Hora-Inicio'], errors='coerce', format='%H:%M:%S').dt.time
data1['Hora-Fim'] = pd.to_datetime(data1['Hora-Fim'], errors='coerce', format='%H:%M:%S').dt.time

# Diagnóstico após o tratamento
print("Tipos de dados após tratamento:")
print(data2.dtypes)

print("\nValores das colunas 'Hora-Inicio' e 'Hora-Fim':")
print(data1[['Hora-Inicio', 'Hora-Fim']].head(20))
