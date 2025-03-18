import pandas as pd

# Carregar os dados da Planilha1
df1 = pd.read_excel('equipe.xlsx', sheet_name='Planilha1', header=None)

# # Carregar os dados da Planilha2
# df2 = pd.read_excel('equipe.xlsx', sheet_name='Planilha2', header=None)

# # Criar um dicionário para mapear nomes para códigos de equipe
# equipe_map = {}
# for index, row in df2.iterrows():
#     if pd.notna(row[0]) and pd.notna(row[1]):
#         codigo = row[0]
#         nome = row[1].split('-')[0].strip()  # Remover o número após o traço e espaços extras
#         equipe_map[nome] = codigo

# Adicionar a lista personalizada de mapeamentos
mapeamento_personalizado = {
    "MAURICIO NASCIMENTO": "EX004",
    "NATHAN PEREIRA": "MT009",
    "MAURICIO MARQUES": "MT006",
    "JEFFERSON JUNIOR": "EX017",
    "JEFFERSON  JUNIOR": "EX017",
    "CAIQUE BRAZIEL TEODORO ALVES": "AV001",
    "DANIEL ALVES DA SILVA CARDOSO": "AV002",
    "MARCIO JUNIOR": "AV003",
    "DEIVID LUIS": "EX013",
    "DEIVID LUIZ": "EX013",
    "LUCILIO RODRIGUES": "EX015",
    "ROBERTO NASSAR": "EX011",
    "ROBSON MARQUES": "EX012",
    "FLAVIO GARCIA": "EX018",
    "THIAGO LIMA": "EX001",
    "LUIS CARLOS": "EX002",
    "MARCELO DA ROCHA": "EX003",
    "RAFAEL VICTOR": "EX005",
    "JOSE HIPOLITO": "EX007",
    "RAPHAEL CASEMIRO": "EX008",
    "MAX PAULO": "EX010",
    "ROBERTO NASSAR": "EX011",
    "ROBSON DA SILVA MARQUES": "EX012",
    "DEIVID LUIZ": "EX013",
    "JOSE ROBERTO": "EX014",
    "LUCILIO RODRIGUES": "EX015",
    "REGINALDO DE ANDRADE": "LV001",
    "RICARDO ROBERTO": "LV003",
    "WALBER SPINDOLA": "LV004",
    "JORGE LEANDRO": "MT001",
    "JOSE EDSON": "MT002",
    "MARCEL DO CARMO": "MT003",
    "EDUARDO MARTINS FERREIRA": "MT004",
    "VINICIO JOSE": "MT005",
    "MAURILIO MARQUES": "MT007",
    "WELLINGTON BARCELOS": "MT008",
    "OTAVIO VELOSO ": "EX016",
}

# # Combinar o mapeamento personalizado com o mapeamento da Planilha2
# equipe_map.update(mapeamento_personalizado)

# Função para mapear nomes para códigos de equipe
def mapear_equipe(nomes):
    equipes = []
    for nome in nomes.split('/'):  # Dividir os nomes por barra
        nome = nome.strip()  # Remover espaços extras
        
        encontrado = False # Variável para indicar se o nome foi encontrado
        # Iterar sobre o mapeamento personalizado para encontrar a correspondência
        for chave in mapeamento_personalizado:
            if nome in chave or chave in nome:  # Verificar se o nome contém a chave
                equipes.append(mapeamento_personalizado[chave])  # Adicionar o código de equipe correspondente
                encontrado = True  # Marcar que o nome foi encontrado
                break
        if not encontrado:
            equipes.append('N/A')  # Caso o nome não seja encontrado
    return '/'.join(equipes)  # Juntar os códigos de equipe com barra

# Aplicar a função à coluna A da Planilha1
df1['Equipe'] = df1[0].apply(mapear_equipe)

# Salvar o DataFrame modificado em um novo arquivo .xlsx
df1.to_excel('equipe_modificado.xlsx', index=False, header=False)