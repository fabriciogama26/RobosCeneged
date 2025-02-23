import os
import pandas as pd

# Caminho da pasta
folder_path = "C:\\Users\\fabriciogama\\Downloads\\FOLHA DE MEDIÇÃO - JANEIRO 2025"

# Listar todos os arquivos na pasta
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

# Inicializar lista para armazenar os dados organizados
final_data = []

# Iterar sobre os arquivos Excel
for file in excel_files:
    file_path = os.path.join(folder_path, file)  # Construir o caminho completo
    print(f"Lendo arquivo: {file}")
    
    try:
        # Ler a aba "Folha de Medição"
        data = pd.read_excel(file_path, sheet_name="Folha de Medição")
        data['Arquivo Origem'] = file  # Adicionar uma coluna indicando o arquivo de origem

        # Extrair "Projeto"
        data['Projeto'] = data.apply(
            lambda row: row['Unnamed: 11'] if 'Projeto:' in str(row['Unnamed: 10']) else None, axis=1
        )
        data['Projeto'] = data['Projeto'].ffill().astype(str).str.replace('.0', '', regex=False).str.strip()

        # Extrair "Registro ( Nº FM)"
        data['Registro ( Nº FM)'] = data.apply(
            lambda row: row['Unnamed: 3'] if 'Registro ( Nº FM):' in str(row['Unnamed: 2']) else None, axis=1
        )
        data['Registro ( Nº FM)'] = data['Registro ( Nº FM)'].ffill().astype(str).str.replace('.0', '', regex=False).str.strip()

        # print("Diagnóstico para verificar valores:"
        #       "\Registro ( Nº FM):", data['Registro ( Nº FM)'].dropna().tolist())

        # Extrair "Data"
        data['Data'] = data.apply(
            lambda row: row['Unnamed: 8'] if 'Data:' in str(row['Unnamed: 7']) else None, axis=1
        )
        data['Data'] = data['Data'].ffill()

        # Localizar início das seções "Descrição do Serviço" , "Serviço" e "Total Valor"
        descricao_start = data[data['Unnamed: 2'].astype(str).str.contains('Descrição do Serviço', na=False)].index
        servico_start = data[data['Unnamed: 7'].astype(str).str.contains('Serviço', na=False)].index
        total_valor_start = data[data['Unnamed: 12'].astype(str).str.contains('Total Valor', na=False)].index

        # Verificar se as seções foram encontradas

        if len(descricao_start) > 0 and len(servico_start) > 0 and len(total_valor_start) > 0:
            descricao_values = data.loc[descricao_start[0] + 1:, 'Unnamed: 2'].dropna().reset_index(drop=True)
            servico_values = data.loc[servico_start[0] + 1:, 'Unnamed: 7'].dropna().reset_index(drop=True)
            total_valor_values = data.loc[total_valor_start[0] + 1:, 'Unnamed: 12'].dropna().reset_index(drop=True)

            # Processar dados por "Arquivo Origem"
            for origem in data['Arquivo Origem'].unique():
                origem_data = data[data['Arquivo Origem'] == origem]

                # print("Diagnóstico para verificar valores:")
                # print("Projeto:", origem_data['Projeto'].dropna().tolist())
                # print("Data:", origem_data['Data'].dropna().tolist())
                # print("Total Valor:", origem_data['Unnamed: 12'].dropna().tolist())
                # print("Descrição do Serviço:", descricao)
                # print("Serviço:", servico)
                # print("Origem:", origem)
                # print("--------------------------------------------------")

                # Iterar sobre os valores extraídos
                for i, (descricao, servico, total_valor) in enumerate(zip(descricao_values, servico_values, total_valor_values)):
                    # Interromper processamento atual se encontrar células em branco
                    if pd.isnull(descricao) or pd.isnull(servico) or pd.isnull(total_valor) or descricao == "" or servico == "" or total_valor == "":
                        break
                        
                    # Adicionar os dados processados à lista final
                    final_data.append({
                        'Projeto': origem_data['Projeto'].dropna().iloc[6],  # Primeiro valor válido de "Projeto"
                        'Data': origem_data['Data'].dropna().iloc[6],  # Primeiro valor válido de "Data"
                        'Folha de Medição': origem_data['Registro ( Nº FM)'].dropna().iloc[7],  # Primeiro valor válido de "Folha de Medição"
                        'Total Valor': total_valor,
                        'Descrição do Serviço': descricao,
                        'Serviço': servico,
                        'Origem': origem
                    })

    except Exception as e:
        print(f"Erro ao ler ou processar o arquivo '{file}': {e}")

# Filtrar linhas com "SEM Restrição" em "Descrição do Serviço"
final_data = [row for row in final_data if row['Descrição do Serviço'] != "SEM Restrição"]

# Converter os dados processados em um DataFrame final
if final_data:
    final_df = pd.DataFrame(final_data)

    # Salvar o DataFrame final em um novo arquivo Excel
    output_path = "C:\\Users\\fabriciogama\\Downloads\\Folha_Medicao_Organizada_Final.xlsx"
    final_df.to_excel(output_path, index=False)

    print(f"Dados organizados salvos em: {output_path}")
else:
    print("Nenhum dado válido encontrado ou processado.")
