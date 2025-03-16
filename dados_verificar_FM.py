import pandas as pd

# Arquivos
output_data = "Folha_Medicao_Organizada.xlsx"

output_data2 = "FOLHA MEDIÇÃO CONSOLIDADO.xlsx"

# Salvar o DataFrame final em um novo arquivo Excel com o nome personalizado
output_path = f"C:\\Users\\fabriciogama\\Downloads\\{output_data}"

# Salvar o DataFrame final em um novo arquivo Excel com o nome personalizado
output_path2 = f"C:\\Users\\fabriciogama\\Downloads\\{output_data2}"

# Carregar os arquivos Excel em DataFrames
df_output_data = pd.read_excel(output_path)
df_output_data2 = pd.read_excel(output_path2)

# Verificar se a coluna "Folha de Medição" existe em ambos os DataFrames
if 'Folha de Medição' in df_output_data.columns and 'Folha de Medição' in df_output_data2.columns:
    # Extrair os valores da coluna "FM" de ambos os DataFrames
    fm_output_data = df_output_data['Folha de Medição'].dropna().unique()  # Valores únicos da coluna "FM" no output_data
    fm_output_data2 = df_output_data2['Folha de Medição'].dropna().unique()  # Valores únicos da coluna "Folha de Medição" no output_data2

    # Comparar os valores
    # Valores que estão em output_data mas não em output_data2
    valores_faltantes_em_output_data2 = set(fm_output_data) - set(fm_output_data2)

    # Valores que estão em output_data2 mas não em output_data
    valores_faltantes_em_output_data = set(fm_output_data2) - set(fm_output_data)

    # Exibir resultados 
    if valores_faltantes_em_output_data2:
        print(f"Valores na coluna 'Folha de Medição' de {output_data} que não estão em {output_data2}:")
        for valor in valores_faltantes_em_output_data2:
            print(valor)
    else:
        print(f"Todos os valores da coluna 'Folha de Medição' de {output_data} estão presentes em {output_data2}.")

    # if valores_faltantes_em_output_data:
    #     print("\nValores na coluna 'Folha de Medição' de output_data2 que não estão em output_data:")
    #     for valor in valores_faltantes_em_output_data:
    #         print(valor)
    # else:
    #     print("\nTodos os valores da coluna 'Folha de Medição' de output_data2 estão presentes em output_data.")
else:
    print("A coluna 'Folha de Medição' não foi encontrada em um ou ambos os arquivos.")