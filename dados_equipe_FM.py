import pandas as pd
from thefuzz import process

# Carregar os dados da Planilha1
df1 = pd.read_excel(r"C:\Users\fabriciogama\Downloads\Equipe.xlsx", sheet_name='Planilha1', header=None)

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
    "JEFFERSON": "EX017",
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
    "HIPÓLITO": "EX007",
    "RAPHAEL CASEMIRO": "EX008",
    "RAFAEL CASEMIRO": "EX008",
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
    "J. EDSON": "MT002",
    "MARCEL DO CARMO": "MT003",
    "EDUARDO MARTINS FERREIRA": "MT004",
    "VINICIO JOSE": "MT005",
    "VINICIO JOSÉ": "MT005",
    "MAURILIO MARQUES": "MT007",
    "WELLINGTON BARCELOS": "MT008",
    "OTAVIO VELOSO ": "EX016",
    "MAURICIO SILVA": "EX004",
}

# # Combinar o mapeamento personalizado com o mapeamento da Planilha2
# equipe_map.update(mapeamento_personalizado)

# Função para encontrar a melhor correspondência aproximada
def encontrar_correspondencia(nome, mapeamento, limite=50):

    # Usar fuzzywuzzy para encontrar a melhor correspondência
    melhor_correspondencia, pontuacao = process.extractOne(nome, mapeamento.keys())
    # Retornar o código se a pontuação for maior que o limite
    if pontuacao >= limite:
        return mapeamento[melhor_correspondencia]
    else:
        return 'N/A'  # Caso não encontre uma correspondência válida

# Função para mapear nomes para códigos de equipe
def mapear_equipe(nomes):
    equipes = []
    for nome in nomes.split('/'):
        nome = nome.strip()  # Remover espaços extras
        # Encontrar a correspondência aproximada
        codigo = encontrar_correspondencia(nome, mapeamento_personalizado)
        equipes.append(codigo)
    return '/'.join(equipes)

# Aplicar a função à coluna A da Planilha1
df1['Equipe'] = df1[0].apply(mapear_equipe)

# Salvar o DataFrame final em um novo arquivo Excel
output_data = 'equipe_modificado.xlsx'

# Salvar o DataFrame final em um novo arquivo Excel com o nome personalizado
output_path = f"C:\\Users\\fabriciogama\\Downloads\\{output_data}"
# index=False para evitar a criação de uma coluna de index

# Salvar o DataFrame modificado em um novo arquivo .xlsx
df1.to_excel(output_path, index=False, header=False)