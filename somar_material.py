import pandas as pd

# Ler o arquivo Excel
caminho_arquivo = r"C:\Users\fabriciogama\Downloads\Material.xlsx"  # Substitua pelo caminho do seu arquivo
df = pd.read_excel(caminho_arquivo)

# Agrupar por 'Material' e somar as quantidades
resultado = df.groupby(['Material', 'Descrição Material', 'UDM'], as_index=False)['Qtde'].sum()

# Salvar o resultado em um novo arquivo Excel (opcional)
resultado.to_excel(r"C:\Users\fabriciogama\Downloads\resultado_soma.xlsx", index=False)

# Exibir o resultado no console
print(resultado)