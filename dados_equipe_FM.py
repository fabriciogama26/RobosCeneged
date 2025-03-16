import pandas as pd

# Carregar os dados da Planilha1
df1 = pd.read_excel('equipe.xlsx', sheet_name='Planilha1', header=None)

# Carregar os dados da Planilha2
df2 = pd.read_excel('equipe.xlsx', sheet_name='Planilha2', header=None)

# Criar um dicionário para mapear nomes para códigos de equipe
equipe_map = {}
for index, row in df2.iterrows():
    if pd.notna(row[0]) and pd.notna(row[1]):
        codigo = row[0]
        nome = row[1].split('-')[0].strip()  # Remover o número após o traço e espaços extras
        equipe_map[nome] = codigo

# Função para mapear nomes para códigos de equipe
def mapear_equipe(nomes):
    equipes = []
    for nome in nomes.split('/'):
        nome = nome.strip()  # Remover espaços extras
        # Tentar encontrar correspondência exata ou parcial
        encontrado = False
        for chave in equipe_map:
            if nome in chave or chave in nome:
                equipes.append(equipe_map[chave])
                encontrado = True
                break
        if not encontrado:
            equipes.append('N/A')  # Caso o nome não seja encontrado
    return '/'.join(equipes)

# Aplicar a função à coluna A da Planilha1
df1['Equipe'] = df1[0].apply(mapear_equipe)

# Salvar o DataFrame modificado em um novo arquivo .xlsx
df1.to_excel('equipe_modificado.xlsx', index=False, header=False)