import pandas as pd

# Carregar os arquivos
df_nomes_abreviados = pd.read_excel('nomes_abreviados.csv')  # Arquivo com os nomes abreviados
df_nomes_completos = pd.read_excel('nomes_completos.csv')    # Arquivo com os nomes completos

# Criar um dicionário para mapear os nomes abreviados para os nomes completos
mapeamento_nomes = dict(zip(df_nomes_completos['Nome_Abreviado'], df_nomes_completos['Nome_Completo']))

# Substituir os nomes abreviados pelos nomes completos
df_nomes_abreviados['Nome'] = df_nomes_abreviados['Nome'].map(mapeamento_nomes).fillna(df_nomes_abreviados['Nome'])

# Salvar o resultado em um novo arquivo
df_nomes_abreviados.to_csv('nomes_atualizados.csv', index=False)

print("Substituição concluída. Resultado salvo em 'nomes_atualizados.csv'.")