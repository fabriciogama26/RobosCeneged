import os
import pandas as pd

# Caminho da pasta
folder_path = "C:\\Users\\fabriciogama\\OneDrive - CENEGED - COMPANHIA ELETROMECANICA E GERENCIAMENTO DE DADOS\\Leonardo\\Folha de Medição"

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

        # Extrair "Data"
        data['Data'] = data.apply(
            lambda row: row['Unnamed: 8'] if 'Data:' in str(row['Unnamed: 7']) else None, axis=1
        )
        data['Data'] = data['Data'].ffill()

        # Localizar início das seções "Descrição do Serviço" e "Serviço"
        descricao_start = data[data['Unnamed: 2'].astype(str).str.contains('Descrição do Serviço', na=False)].index
        servico_start = data[data['Unnamed: 7'].astype(str).str.contains('Serviço', na=False)].index

        if len(descricao_start) > 0 and len(servico_start) > 0:
            descricao_values = data.loc[descricao_start[0] + 1:, 'Unnamed: 2'].dropna().reset_index(drop=True)
            servico_values = data.loc[servico_start[0] + 1:, 'Unnamed: 7'].dropna().reset_index(drop=True)

            # Processar dados por "Arquivo Origem"
            for origem in data['Arquivo Origem'].unique():
                origem_data = data[data['Arquivo Origem'] == origem]

                # Iterar sobre os valores extraídos
                for i, (descricao, servico) in enumerate(zip(descricao_values, servico_values)):
                    # Interromper processamento atual se encontrar células em branco
                    if pd.isnull(descricao) or pd.isnull(servico) or descricao == "" or servico == "":
                        break

                    # Adicionar os dados processados à lista final
                    final_data.append({
                        'Projeto': origem_data['Projeto'].dropna().iloc[6],  # Primeiro valor válido de "Projeto"
                        'Data': origem_data['Data'].dropna().iloc[6],  # Primeiro valor válido de "Data"
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
