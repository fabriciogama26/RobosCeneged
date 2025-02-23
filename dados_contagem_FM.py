import pandas as pd

# Carregar o arquivo Excel
file_path = "C:\\Users\\fabriciogama\\Downloads\\Folha_Medicao_janeiro.xlsx"
df = pd.read_excel(file_path)

# Converter a coluna 'Data' para o formato datetime
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')

# Contar as ocorrências de 'Folha de Medição' para cada data
contagem_fm_por_data = df.groupby(['Data', 'Folha de Medição']).size().reset_index(name='Quantidade')

# Exibir o resultado
print(contagem_fm_por_data)
