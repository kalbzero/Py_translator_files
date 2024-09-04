import openpyxl
from googletrans import Translator
from tqdm import tqdm
import json
import os
import time
from bs4 import BeautifulSoup

# Inicializa o tradutor e o cache
tradutor = Translator()
cache_traducoes = {}
cache_arquivo = 'cache_traducoes.json'

def salvar_cache():
    with open(cache_arquivo, 'w', encoding='utf-8') as f:
        json.dump(cache_traducoes, f, ensure_ascii=False, indent=4)

def is_number(texto):
    # Remove todos os pontos, vírgulas e traços
    texto_limpo = texto.replace('.', '').replace(',', '').replace('-', '')
    return texto_limpo.isdigit()

def traduzir_texto(texto):
    if texto.strip() in cache_traducoes:
        return cache_traducoes[texto.strip()]

    # Verifica se o texto é um número, um link ou começa com "Image"
    if is_number(texto) or texto.startswith('http') or texto.startswith('Image'):
        return texto  # Não traduz números, links ou textos que começam com "Image"

    # Substitui ponto e vírgula por espaço antes de traduzir
    texto_sem_pontovirgula = texto.replace(';', ' ')

    tentativas = 3
    for _ in range(tentativas):
        try:
            # Remover tags HTML e obter o texto limpo
            soup = BeautifulSoup(texto_sem_pontovirgula, "html.parser")
            texto_limpo = soup.get_text()

            # Traduzir o texto limpo
            traducao = tradutor.translate(texto_limpo, src='es', dest='pt').text

            # Restaurar as tags HTML no texto traduzido
            traducao_html = BeautifulSoup(traducao, "html.parser").prettify()
            texto_traduzido = traducao_html

            # Armazenar a tradução no cache
            cache_traducoes[texto.strip()] = texto_traduzido
            return texto_traduzido
        except Exception as e:
            print(f"Erro ao tentar traduzir '{texto}': {e}")
            # Verifica se a mensagem de erro indica o limite de traduções
            if "AVAILABLE FREE TRANSLATIONS" in str(e):
                print("Limite de traduções atingido. Encerrando o programa.")
                salvar_cache()
                exit(1)
            time.sleep(5)  # Espera antes de tentar novamente

    print("Máximo de tentativas atingido. Pulando para a próxima palavra.")
    return texto

def traduzir_arquivo_xlsx(nome_arquivo):
    print(f"Iniciando tradução do arquivo: {nome_arquivo}")

    try:
        # Carrega o arquivo XLSX
        workbook = openpyxl.load_workbook(nome_arquivo)
        sheet = workbook.active

        # Coleta os textos para traduzir
        textos_para_traduzir = []
        textos_original = {}  # Mapeia os textos originais para suas células

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value not in textos_original:
                        textos_original[cell.value] = cell
                        textos_para_traduzir.append(cell.value)

        # Tradução dos textos
        textos_traduzidos = {}
        for texto in tqdm(textos_para_traduzir, desc="Traduzindo textos"):
            traduzido = traduzir_texto(texto)
            textos_traduzidos[texto] = traduzido

        for texto_original, texto_traduzido in textos_traduzidos.items():
            if texto_original in textos_original:
                textos_original[texto_original].value = texto_traduzido

        nome_base, extensao = os.path.splitext(nome_arquivo)
        novo_nome = f"{nome_base}_pt.xlsx"
        workbook.save(novo_nome)
        salvar_cache()
        
        print(f"Tradução concluída. Arquivo salvo como: {novo_nome}")
        print(f"Total de traduções únicas: {len(cache_traducoes)}")

    except KeyboardInterrupt:
        print("\nInterrupção detectada! Salvando o progresso...")
        salvar_cache()
        nome_base, extensao = os.path.splitext(nome_arquivo)
        novo_nome = f"{nome_base}_parcial_pt.xlsx"
        workbook.save(novo_nome)
        print(f"Progresso salvo no arquivo: {novo_nome}")
        print(f"Total de traduções únicas: {len(cache_traducoes)}")
        raise

# Traduzir o arquivo XLSX
arquivo_para_traduzir = 'arquivo_entrada.xlsx'

if os.path.exists(arquivo_para_traduzir):
    traduzir_arquivo_xlsx(arquivo_para_traduzir)
    print("O arquivo foi traduzido com sucesso!")
else:
    print(f"O arquivo {arquivo_para_traduzir} não foi encontrado.")
