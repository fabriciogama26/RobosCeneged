import os
import pandas as pd
from datetime import datetime

# Obter a data e hora atual no formato desejado (por exemplo: YYYY-MM-DD_HH-MM-SS)
current_time = datetime.now().strftime("%d-%m-%Y_%Hh%M")

# Caminho da pasta
folder_path = r"C:\Users\fabriciogama\Downloads\Mediçao"

# Listar todos os arquivos na pasta
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

# Inicializar lista para armazenar os dados organizados
final_data = []

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

        # Localizar início das seções "QTD", "Descrição do Serviço" , "Serviço" e "Total Valor"
        quantidade_start = data[data['Unnamed: 11'].astype(str).str.contains('QTD', na=False)].index
        descricao_start = data[data['Unnamed: 2'].astype(str).str.contains('Descrição do Serviço', na=False)].index
        servico_start = data[data['Unnamed: 7'].astype(str).str.contains('Serviço', na=False)].index
        total_valor_start = data[data['Unnamed: 12'].astype(str).str.contains('Total Valor', na=False)].index

        # Verificar se as seções foram encontradas e extrair os valores correspondentes
        if len(quantidade_start) > 0 and len(descricao_start) > 0 and len(servico_start) > 0 and len(total_valor_start) > 0:
            # Extrair os valores das seções "QTD", "Descrição do Serviço" , "Serviço" e "Total Valor" a partir do início encontrado e remover linhas vazias
            quantidade_values = data.loc[quantidade_start[0] + 1:, 'Unnamed: 11'].dropna().reset_index(drop=True)
            descricao_values = data.loc[descricao_start[0] + 1:, 'Unnamed: 2'].dropna().reset_index(drop=True)
            servico_values = data.loc[servico_start[0] + 1:, 'Unnamed: 7'].dropna().reset_index(drop=True)
            total_valor_values = data.loc[total_valor_start[0] + 1:, 'Unnamed: 12'].dropna().reset_index(drop=True)

            # if contrato_values == "36-lv":
            #     contrato_values == "Linha Viva"
            # elif contrato_values == "36-lm":
            #     contrato_values == "Linha Morta"
            # elif contrato_values == "spot-manutencao":
            #     contrato_values == "Manutencao"
            # elif contrato_values == "spot-expansao":
            #     contrato_values == "Expansao"

            # Processar dados por "Arquivo Origem" 
            for origem in data['Arquivo Origem'].unique():
                origem_data = data[data['Arquivo Origem'] == origem]

                
                # print("--------------------------------------------------")
                # print("Diagnóstico para verificar valores:")
                # print("Projeto:", origem_data['Projeto'].dropna().tolist())
                # print("Data:", origem_data['Data'].dropna().tolist())
                # print("Total Valor:", origem_data['Unnamed: 12'].dropna().tolist())
                # print("Descrição do Serviço:", descricao)
                # print("Quantidade:", quantidade)
                # print("Serviço:", servico)
                # print("Origem:", origem)
                # print("--------------------------------------------------")

                # Iterar sobre os valores extraídos e processar os dados
                try:
                    for i, (quantidade, descricao,servico, total_valor) in enumerate(zip( quantidade_values, descricao_values, servico_values, total_valor_values)):
                        # Interromper processamento atual se encontrar células em branco
                        if pd.isnull(quantidade) or pd.isnull(descricao) or pd.isnull(servico) or pd.isnull(total_valor) or quantidade == "" or descricao == "" or servico == "" or total_valor == "":
                            break


                        # Adicionar os dados processados à lista final
                        final_data.append({
                            'Projeto': origem_data['Projeto'].dropna().iloc[6],  # Primeiro valor válido de "Projeto"
                            'Data': origem_data['Data'].dropna().iloc[6],  # Primeiro valor válido de "Data"
                            'Folha de Medição': origem_data['Registro ( Nº FM)'].dropna().iloc[7],  # Primeiro valor válido de "Folha de Medição"
                            'Equipe': origem_data['Encarregado'].dropna().iloc[6],  # Primeiro valor valido de "Equipe"
                            'Total Valor': total_valor,
                            'Descrição do Serviço': descricao,
                            'Quantidade': quantidade,
                            'Serviço': servico,
                            'Origem': origem
                        })

                except Exception as e:
                    print(f"Erro ao processar os dados do arquivo '{file}': {e}")

    except Exception as e:
        print(f"Erro ao ler ou processar o arquivo '{file}': {e}")

# Filtrar linhas com "SEM Restrição" em "Descrição do Serviço"
final_data = [row for row in final_data if row['Descrição do Serviço'] != "SEM Restrição"]

# Converter os dados processados em um DataFrame final
if final_data:
    final_df = pd.DataFrame(final_data)

    # Salvar o DataFrame final em um novo arquivo Excel
    output_data = F"Folha_Medicao_Organizada_{current_time}.xlsx"

    # Salvar o DataFrame final em um novo arquivo Excel com o nome personalizado
    output_path = f"C:\\Users\\fabriciogama\\Downloads\\{output_data}"
    # index=False para evitar a criação de uma coluna de index
    final_df.to_excel(output_path, index=False)

    print(f"Dados organizados salvos em: {output_path}_{current_time}")
else:
    print("Nenhum dado válido encontrado ou processado.")


# Contar a quantidade de arquivos Excel na pasta
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
num_arquivos_pasta = len(excel_files)

# Contar a quantidade de arquivos processados no arquivo salvo e identificar arquivos faltantes
try:
    df = pd.read_excel(output_path)
    num_arquivos_salvos = df['Origem'].nunique()  # Contar quantos arquivos diferentes aparecem na coluna "Origem"
    
    # Converter as listas para conjuntos
    arquivos_pasta_set = set(excel_files)  # Conjunto de arquivos na pasta
    arquivos_processados_set = set(df['Origem'].unique())  # Conjunto de arquivos processados

    # Encontrar a diferença entre os conjuntos (arquivos na pasta que não foram processados)
    arquivos_faltantes_set = arquivos_pasta_set - arquivos_processados_set

    # Converter o conjunto de volta para uma lista (se necessário)
    arquivos_faltantes = list(arquivos_faltantes_set)
    
    # Exibir resultados de contagem de arquivos
    print(f"Quantidade de arquivos Excel na pasta: {num_arquivos_pasta}")
    print(f"Quantidade de arquivos registrados no arquivo {output_data}: {num_arquivos_salvos}")
    
    # Exibir arquivos faltantes
    if arquivos_faltantes:
        print("Arquivos faltantes (não processados ou não registrados):")
        # Exibir os arquivos faltantes dentro da pasta
        for arquivo in arquivos_faltantes:
            print(f"- {arquivo}")
    else:
        print("Todos os arquivos foram processados e registrados.")
        
except Exception as e:
    print(f"Erro ao ler o arquivo salvo: {e}")
