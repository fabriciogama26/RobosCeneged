import os
import pandas as pd

# Caminho da pasta
folder_path = r"C:\Users\fabriciogama\Downloads\Mediçao"

# Listar todos os arquivos na pasta
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

# Inicializar lista para armazenar os dados organizados
final_data = []

# Inicializar lista para armazenar os dados organizados
final_data_2 = []

# Iterar sobre os arquivos Excel
for file in excel_files:
    file_path = os.path.join(folder_path, file)  # Construir o caminho completo do arquivo
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

        # Extrair "Registro (Nº FM)"
        data['Registro ( Nº FM)'] = data.apply(
            lambda row: row['Unnamed: 3'] if 'Registro ( Nº FM):' in str(row['Unnamed: 2']) else None, axis=1
        )
        data['Registro ( Nº FM)'] = data['Registro ( Nº FM)'].ffill().astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # Extrair "Equipe"
        data['Encarregado'] = data.apply(
            lambda row: row['Unnamed: 3'] if 'Encarregado:' in str(row['Unnamed: 2']) else None, axis=1
        )
        data['Encarregado'] = data['Encarregado'].ffill()

        # Extrair "Data"
        data['Data'] = data.apply(
            lambda row: row['Unnamed: 8'] if 'Data:' in str(row['Unnamed: 7']) else None, axis=1
        )
        data['Data'] = data['Data'].ffill()

        # Filtrar o DataFrame para incluir apenas as linhas a partir da linha 55
        # data_filtered = data.iloc[55:]

        # Localizar início das seções "Quantidade", "Descrição Material" , "Lote" , "Quantidade Retirado" , "Descrição Material Retirado" e "Lote Retirado"
        quantidade_material_start = data[data['Unnamed: 6'].astype(str).str.contains('Quantidade', na=False)].index
        descricao_material_start = data[data['Unnamed: 1'].astype(str).str.contains('Descrição Material', na=False)].index
        lote_material_start = data[data['Unnamed: 5'].astype(str).str.contains('Lote', na=False)].index
        quantidade_material_retirado_start = data[data['Unnamed: 13'].astype(str).str.contains('Quantidade', na=False)].index
        descricao_material_retirado_start = data[data['Unnamed: 10'].astype(str).str.contains('Descrição Material', na=False)].index
        lote_material_retirado_start = data[data['Unnamed: 12'].astype(str).str.contains('Lote', na=False)].index

        # print(data.iloc[50:60])

        # print(quantidade_material_start)
        # print(descricao_material_start)
        # print(lote_material_start)
        # print(quantidade_material_retirado_start)
        # print(descricao_material_retirado_start)

        # Verificar se as seções foram encontradas e extrair os valores correspondentes
        if (len(quantidade_material_start) > 0 and len(descricao_material_start) > 0 and len(lote_material_start) > 0 and len(quantidade_material_retirado_start)> 0 
        and len(descricao_material_retirado_start) > 0 and len(lote_material_retirado_start) > 0):
            # Extrair os valores das seções "Quantidade", "Descrição Material" , "Lote" , "Quantidade Retirado" , "Descrição Material Retirado" e
            #  "Lote Retirado" a partir do início encontrado e remover linhas vazias
            quantidade_material_values = data.loc[quantidade_material_start[0] + 1:, 'Unnamed: 6'].dropna().reset_index(drop=True)
            descricao_material_values = data.loc[descricao_material_start[0] + 1:, 'Unnamed: 1'].dropna().reset_index(drop=True)
            lote_material_values = data.loc[lote_material_start[0] + 1:, 'Unnamed: 5'].dropna().reset_index(drop=True)
            quantidade_material_retirado_values = data.loc[quantidade_material_retirado_start[0] + 1:, 'Unnamed: 13'].dropna().reset_index(drop=True)
            descricao_material_retirado_values = data.loc[descricao_material_retirado_start[0] + 1:, 'Unnamed: 10'].dropna().reset_index(drop=True)
            lote_material_retirado_values = data.loc[lote_material_retirado_start[0] + 1:, 'Unnamed: 12'].dropna().reset_index(drop=True)

            # Processar dados por "Arquivo Origem" 
            for origem in data['Arquivo Origem'].unique():
                origem_data = data[data['Arquivo Origem'] == origem]
                
                # print("--------------------------------------------------")
                # print("Diagnóstico para verificar valores:")
                # print("Projeto:", origem_data['Projeto'].dropna().tolist())
                # print("Data:", origem_data['Data'].dropna().tolist())
                # print("Quantidade Material:", quantidade_material_values)
                # print("Descrição Material:", descricao_material_values)
                # print("Lote Material:", lote_material_values)
                # print("Quantidade Retirado:", quantidade_material_retirado_values)
                # print("Descrição Material Retirado:", descricao_material_retirado_values)
                # print("Lote Retirado:", lote_material_retirado_values)
                # print("Origem:", origem)
                # print("--------------------------------------------------")

                # Iterar sobre os valores extraídos e processar os dados
                for i, (quantidade_material, descricao_material, lote_material, quantidade_retirado, descricao_retirado, lote_retirado) in enumerate(zip( quantidade_material_values, descricao_material_values, 
                lote_material_values, quantidade_material_retirado_values, descricao_material_retirado_values, lote_material_retirado_values)):
                    # Interromper processamento atual se encontrar células em branco
                    if (pd.isnull(quantidade_material) or pd.isnull(descricao_material) or pd.isnull(lote_material) or pd.isnull(quantidade_retirado) or pd.isnull(descricao_retirado)
                    or pd.isnull(lote_retirado) or quantidade_material == "" or descricao_material == "" or lote_material == "" or quantidade_retirado == "" 
                    or descricao_retirado == "" or lote_retirado == ""):
                        break
                        
                    # Adicionar os dados processados à lista final
                    final_data.append({
                        'Projeto': origem_data['Projeto'].dropna().iloc[6],  # Primeiro valor válido de "Projeto"
                        'Data': origem_data['Data'].dropna().iloc[6],  # Primeiro valor válido de "Data"
                        'Folha de Medição': origem_data['Registro ( Nº FM)'].dropna().iloc[7],  # Primeiro valor válido de "Folha de Medição"
                        'Equipe': origem_data['Encarregado'].dropna().iloc[6],  # Primeiro valor valido de "Equipe"
                        'Quantidade Material': quantidade_material, # Primeiro valor valido de "Quantidade Material"
                        'descricao_material': descricao_material, # Primeiro valor valido de "Descrição Material"
                        'Lote Material': lote_material, # Primeiro valor valido de "Lote Material"
                        'quantidade_retirado': quantidade_retirado, # Primeiro valor valido de "Quantidade Retirado"
                        'descricao_retirado': descricao_retirado, # Primeiro valor valido de "Descrição Material Retirado"
                        'Lote Retirado': lote_retirado, # Primeiro valor valido de "Lote Retirado"
                        'Origem': origem # Nome do arquivo
                    })
                    

    except Exception as e:
        print(f"Erro ao ler ou processar o arquivo '{file}': {e}")

# Filtrar linhas com "SEM Restrição" em "Descrição do Serviço"
final_data = [row for row in final_data if row['descricao_material'] != "SEM Restrição"]

    # Converter os dados processados em um DataFrame final
if final_data:
    final_df = pd.DataFrame(final_data)

    folder_path2 = r"C:\Users\fabriciogama\Downloads"

    # Criar DataFrame para material aplicado
    material_aplicado_df = pd.DataFrame({
        'cod_material*': final_df['Lote Material'],
        'situacao': 'NV',
        'local_aplicado*': 'GERAL',
        'qtd_material': final_df['Quantidade Material'],
        'rastro_material': ''  # Deixar vazio ou preencher conforme necessário
    })

    # Criar DataFrame para material retirado
    material_retirado_df = pd.DataFrame({
        'cod_material*': final_df['Lote Retirado'],
        'situacao': 'SC',
        'local_aplicado*': 'GERAL',
        'qtd_material': final_df['Quantidade Retirado'],
        'rastro_material': ''  # Deixar vazio ou preencher conforme necessário
    })

    # Salvar os DataFrames em arquivos CSV com codificação UTF-8
    output_folder = folder_path2  # Pasta de saída (mesma pasta de entrada)
    material_aplicado_path = os.path.join(output_folder, "modelo_importacao_serv_material_lote_aplicado.csv")
    material_retirado_path = os.path.join(output_folder, "modelo_importacao_serv_material_lote_retirado.csv")

    material_aplicado_df.to_csv(material_aplicado_path, index=False, encoding='utf-8')
    material_retirado_df.to_csv(material_retirado_path, index=False, encoding='utf-8')

    print(f"Arquivo de material aplicado salvo em: {material_aplicado_path}")
    print(f"Arquivo de material retirado salvo em: {material_retirado_path}")
else:
    print("Nenhum dado válido encontrado ou processado.")
