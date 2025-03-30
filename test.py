def extrair_linhas_com_palavras_chave(nome_arquivo, palavras_chave):
    linhas_encontradas = []

    with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
        for linha in arquivo:
            if any(palavra.lower() in linha.lower() for palavra in palavras_chave):
                linhas_encontradas.append(linha.strip())

    return linhas_encontradas

# Nome do arquivo
nome_arquivo = 'robo_apontamento_log.txt'

# Palavras-chave que você quer buscar
palavras_chave = ['erro', 'tentativa']

# Extrair as linhas que contêm as palavras-chave
linhas_encontradas = extrair_linhas_com_palavras_chave(nome_arquivo, palavras_chave)

# Exibir as linhas encontradas
for linha in linhas_encontradas:
    print(linha)

