import pandas as pd
from difflib import SequenceMatcher
import spacy
import os
import logging

# Configurar o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Instalar spaCy e xlrd explicitamente
os.system("pip install spacy")
os.system("pip install torch")
os.system("pip install xlrd")
os.system("python -m spacy download pt_core_news_sm")

logging.info("Instalação de pacotes concluída")

nlp = spacy.load('pt_core_news_sm')
logging.info("Modelo de linguagem carregado")

def ler_xls(nome_arquivo):
    logging.info(f"Lendo o arquivo {nome_arquivo}")
    if nome_arquivo.endswith('.xls') or nome_arquivo.endswith('.xlsx'):
        df = pd.read_excel(nome_arquivo)
        return df
    else:
        logging.error("Formato de arquivo não suportado. Por favor, insira um arquivo .xls ou .xlsx")
        return None

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def remover_verbos_e_preposicoes(frase):
    doc = nlp(frase)
    nova_frase = ' '.join([token.text for token in doc if token.pos_ not in ['VERB', 'ADP']])
    return nova_frase

def encontrar_strings_similares(df1, df2, limite_similaridade=0.6):
    logging.info("Iniciando a busca por strings similares")
    similaridades = {}

    for index1, row1 in df1.iterrows():
        frase1 = str(row1['NOME DAS ATIVIDADES']).lower()
        frase1_sem_verbos_preposicoes = remover_verbos_e_preposicoes(frase1)
        similaridades_frase = []
        similaridade_dominante = 0
        string_similar_sem_verbos_preposicoes = ""

        for index2, row2 in df2.iterrows():
            frase2 = str(row2['ATRIBUIÇÕES']).lower()
            frase2_sem_verbos_preposicoes = remover_verbos_e_preposicoes(frase2)
            sim = similar(frase1_sem_verbos_preposicoes, frase2_sem_verbos_preposicoes)

            if sim > limite_similaridade:
                similaridades_frase.append(sim)
                if sim > similaridade_dominante:
                    similaridade_dominante = sim
                    string_similar = row2['ATRIBUIÇÕES']
                    string_similar_sem_verbos_preposicoes = frase2_sem_verbos_preposicoes

        if similaridade_dominante > 0:
            if string_similar_sem_verbos_preposicoes in similaridades:
                similaridades[string_similar_sem_verbos_preposicoes]['contagem'] += 1
                similaridades[string_similar_sem_verbos_preposicoes]['frases_parametro'].append(frase1_sem_verbos_preposicoes)
                similaridades[string_similar_sem_verbos_preposicoes]['similaridades_parametro'].append(max(similaridades_frase)) # Similaridade dominante para esta frase
            else:
                similaridades[string_similar_sem_verbos_preposicoes] = {'contagem': 1, 'frases_parametro': [frase1_sem_verbos_preposicoes], 'similaridades_parametro': [max(similaridades_frase)], 'similaridade_dominante': max(similaridades_frase)}

    logging.info("Busca por strings similares concluída")
    return similaridades

def calcular_similaridade(frase_parametro, frase_similar):
    return similar(frase_parametro.lower(), frase_similar.lower()) * 100

def salvar_resultados_em_txt(similaridades, nome_arquivo):
    logging.info(f"Salvando resultados em {nome_arquivo}")
    with open(nome_arquivo, 'w') as arquivo:
        arquivo.write("Strings similares, suas contagens, as frases correspondentes e a similaridade dominante:\n")
        for string, info in similaridades.items():
            string_sem_verbos_preposicoes = remover_verbos_e_preposicoes(string)
            arquivo.write(f"{string_sem_verbos_preposicoes}, Contagem: {info['contagem']}, Similaridade Dominante: {info['similaridade_dominante']:.2f}%\n")
            arquivo.write("Frases utilizadas como parâmetro:\n")
            for frase, similaridade in zip(info['frases_parametro'], info['similaridades_parametro']):
                frase_sem_verbos_preposicoes = remover_verbos_e_preposicoes(frase)
                arquivo.write(f"- {frase_sem_verbos_preposicoes}: Similaridade Dominante: {similaridade:.2f}%\n")
            arquivo.write("\n")

    logging.info("Resultados salvos com sucesso em arquivo de texto")

def salvar_resultados_em_excel(similaridades, nome_arquivo):
    logging.info(f"Salvando resultados em {nome_arquivo}")
    data = {'Frase Similar': [], 'Frase Parâmetro': [], 'Similaridade Dominante (%)': []}
    for string, info in similaridades.items():
        string_sem_verbos_preposicoes = remover_verbos_e_preposicoes(string)
        for frase, similaridade in zip(info['frases_parametro'], info['similaridades_parametro']):
            frase_sem_verbos_preposicoes = remover_verbos_e_preposicoes(frase)
            data['Frase Similar'].append(string_sem_verbos_preposicoes)
            data['Frase Parâmetro'].append(frase_sem_verbos_preposicoes)
            data['Similaridade Dominante (%)'].append(similaridade * 100)  # Multiplicar por 100 para obter a porcentagem

    df = pd.DataFrame(data)
    df.to_excel(nome_arquivo, index=False)
    logging.info("Resultados salvos com sucesso em arquivo Excel")

# Carregar os arquivos XLS
df1 = ler_xls('ATIVIDADES EM PRODUCAO.xls')
df2 = ler_xls('2024-04-24 Atribuicoes - SUPPP.xls')

if df1 is not None and df2 is not None:
    # Encontrar todas as strings similares e suas contagens
    similaridades = encontrar_strings_similares(df1, df2)

    # Salvar os resultados em um arquivo Excel
    salvar_resultados_em_excel(similaridades, 'resultados.xlsx')
    logging.info("Resultados salvos com sucesso em 'resultados.xlsx'")

    # Salvar os resultados em um arquivo de texto
    salvar_resultados_em_txt(similaridades, 'resultados.txt')
    logging.info("Resultados salvos com sucesso em 'resultados.txt'")
else:
    logging.error("Erro ao carregar os arquivos XLS")
