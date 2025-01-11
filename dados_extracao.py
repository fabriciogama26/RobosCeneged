import pandas as pd

# Caminhos dos arquivos Excel
file1_path = "C:\\Users\\fabriciogama\\OneDrive - CENEGED - COMPANHIA ELETROMECANICA E GERENCIAMENTO DE DADOS\\Leonardo\\Relatório de Excecução Manutenção.xlsx"
file2_path = "C:\\Users\\fabriciogama\\Downloads\\PLANILHA CONSOLIDADO MEDIÇÃO.xlsx"

# Ler as planilhas do Excel
data1 = pd.read_excel(file1_path, skiprows=1)  # Começa a partir da segunda linha
data2 = pd.read_excel(file2_path)

# Converter para datetime, tratando valores inválidos
data1['Hora-Inicio'] = pd.to_datetime(data1['Hora-Inicio'], errors='coerce', format='%H:%M:%S').dt.strftime('%H:%M')
data1['Hora-Fim'] = pd.to_datetime(data1['Hora-Fim'], errors='coerce', format='%H:%M:%S').dt.strftime('%H:%M')

# Renomear colunas da Planilha 2 para facilitar
data2 = data2.rename(columns=lambda x: x.strip())  # Remove espaços extras nos nomes das colunas

# Substituir valores 'NaT' por strings vazias ou valores padrão
data1['Hora-Inicio'] = data1['Hora-Inicio'].fillna('')
data1['Hora-Fim'] = data1['Hora-Fim'].fillna('')

# Criar a nova tabela organizada
result = []

# Iterar sobre as linhas da primeira planilha
for _, row1 in data1.iterrows():
    pep = row1['PEP']  # Coluna PEP na primeira planilha
    # Filtrar linhas na segunda planilha que têm o mesmo valor em PROJETO
    matches = data2[data2['PROJETO'] == pep]
    
    if not matches.empty:
        for _, row2 in matches.iterrows():
            # Converter PROJETO para string antes de verificar
            projeto_str = str(row2['PROJETO'])
            
            # Determinar o valor de contrato_chosen
            contrato_chosen = "expansão" if "OII" in projeto_str else "manutenção"
            
            # Formatar data e hora
            dat_srv = pd.to_datetime(row2['DATA DE EXECUÇÃO']).strftime('%d/%m/%Y')
            dat_srv2 = pd.to_datetime(row2['DATA DE EXECUÇÃO']).strftime('%d/%m/%Y')
            
            # Extrair valores de 'Hora-Inicio' e 'Hora-Fim'
            hr_inic = row1['Hora-Inicio'] if row1['Hora-Inicio'] else ''
            hr_fim = row1['Hora-Fim'] if row1['Hora-Fim'] else ''
            
            # Montar a nova linha com os dados combinados
            result.append({
                'obras_chosen': row2['PROJETO'],  # PROJETO na segunda planilha
                'contrato_chosen': contrato_chosen,
                'equipe_chosen': row1['Encarregado'],
                'tip_srv_chosen': '',  # Em branco
                'inputString': row1['Cidade'],
                'cod_irr_chosen': 801,  # Valor fixo
                'dat_srv': dat_srv,
                'hr_inic': hr_inic,
                'dat_srv2': dat_srv2,
                'hr_fim': hr_fim,
                'serv_chosen': row2['LOTE'],
                'qtd': row2['VALIDAÇÃO LIGHT']
            })

# Converter o resultado em um DataFrame
result_df = pd.DataFrame(result)

# Garantir a ordem das colunas na nova planilha
final_columns = [
    'obras_chosen', 'contrato_chosen', 'equipe_chosen', 'tip_srv_chosen',
    'inputString', 'cod_irr_chosen', 'dat_srv', 'hr_inic', 'dat_srv2',
    'hr_fim', 'serv_chosen', 'qtd'
]

# Reorganizar as colunas na ordem desejada
result_df = result_df.reindex(columns=final_columns)

# Salvar o resultado em um novo arquivo Excel
output_path = "C:\\Users\\fabriciogama\\Downloads\\Nova pasta\\dados_organizados.xlsx"
result_df.to_excel(output_path, index=False)

print(f"Dados organizados salvos em: {output_path}")
