import pandas as pd

# Carregar o arquivo Excel
file_path = "C:\\Users\\fabriciogama\\Downloads\\Dash\\FOLHA MEDIÇÃO.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Converter a coluna "Data" para datetime
df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

# Ordenar os dados cronologicamente
df = df.sort_values(by=["Data"], ascending=True)

# Unificar as colunas de equipe para contar corretamente a quantidade de equipes
df["Equipes"] = df[["Equipe", "Equipe 2", "Equipe 3", "Equipe 4", "Equipe 5", "Equipe 6"]].apply(lambda x: list(set(x.dropna())), axis=1)

# Separar dados com e sem OS/OM
df_com_osom = df[df["OS/OM"].notna()]
df_sem_osom = df[df["OS/OM"].isna()]

# Agrupar por projeto, contrato, OS/OM e Base, calculando os valores desejados
def processar_dados(dataframe):
    return dataframe.groupby(["Projeto", "Contrato", "OS/OM", "Servico", "Base"], dropna=False).agg(
        Total_de_Dias=("Data", lambda x: x.nunique()),  # Contar dias únicos corretamente
        Quantidade_Equipes=("Equipes", lambda x: len(set(x.sum()))),  # Contar equipes únicas corretamente
        Data_inicio=("Data", "min"),
        Data_fim=("Data", "max"),  # Garantir que a última data do projeto ou OS/OM seja a Data_fim
        Total_Valor=("Total Valor", "sum")
    ).reset_index()

# Processar os dois conjuntos de dados e unir os resultados
df_result_com_osom = processar_dados(df_com_osom)
df_result_sem_osom = processar_dados(df_sem_osom)
df_result = pd.concat([df_result_com_osom, df_result_sem_osom], ignore_index=True)

# Ajustar a formatação da moeda
df_result["Total_Valor"] = df_result["Total_Valor"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

# Salvar o arquivo processado
output_file = "Tabela_Processada.xlsx"
df_result.to_excel(output_file, index=False)

# Exibir os resultados
print(f"Arquivo salvo como {output_file}")

