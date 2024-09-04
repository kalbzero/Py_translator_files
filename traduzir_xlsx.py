import openpyxl
from googletrans import Translator
from tqdm import tqdm
import json
import os
import time
from bs4 import BeautifulSoup
import re

# Inicializa o tradutor e o cache
tradutor = Translator()
cache_traducoes = {}
cache_arquivo = 'cache_traducoes.json'

def salvar_cache():
    """Salva o cache de traduções em um arquivo JSON."""
    with open(cache_arquivo, 'w', encoding='utf-8') as f:
        json.dump(cache_traducoes, f, ensure_ascii=False, indent=4)

def carregar_cache():
    """Carrega o cache de traduções do arquivo JSON."""
    if os.path.exists(cache_arquivo):
        with open(cache_arquivo, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def is_number(texto):
    """Verifica se o texto é um número, removendo pontos, vírgulas e traços."""
    texto_limpo = texto.replace('.', '').replace(',', '').replace('-', '')
    return texto_limpo.isdigit()

def contains_html(texto):
    """Verifica se o texto contém tags HTML básicas."""
    return bool(re.search(r'<[a-z][\s\S]*>', texto))

def traduzir_texto(texto):
    """Traduz o texto para o português, utilizando cache para evitar traduções repetidas."""
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
            # Remover tags HTML se o texto contiver HTML
            if contains_html(texto_sem_pontovirgula):
                soup = BeautifulSoup(texto_sem_pontovirgula, "html.parser")
                texto_limpo = soup.get_text()
            else:
                texto_limpo = texto_sem_pontovirgula

            # Traduzir o texto limpo
            traducao = tradutor.translate(texto_limpo, src='es', dest='pt').text

            # Restaurar as tags HTML no texto traduzido, se necessário
            texto_traduzido = traducao
            if contains_html(texto):
                texto_traduzido = f"<html>{texto_traduzido}</html>"

            # Armazenar a tradução no cache
            cache_traducoes[texto.strip()] = texto_traduzido
            return texto_traduzido
        except Exception as e:
            print(f"Erro ao tentar traduzir '{texto}': {e}")
            if "AVAILABLE FREE TRANSLATIONS" in str(e):
                print("Limite de traduções atingido. Encerrando o programa.")
                salvar_cache()
                exit(1)
            time.sleep(5)  # Espera antes de tentar novamente

    print("Máximo de tentativas atingido. Pulando para a próxima palavra.")
    return texto

def traduzir_arquivo_xlsx(nome_arquivo):
    """Traduz o conteúdo de um arquivo XLSX e salva a tradução em um novo arquivo."""
    print(f"Iniciando tradução do arquivo: {nome_arquivo}")

    try:
        # Carrega o cache de traduções
        global cache_traducoes
        cache_traducoes = carregar_cache()

        # Carrega o arquivo XLSX
        workbook = openpyxl.load_workbook(nome_arquivo)
        sheet = workbook.active

        # Tradução dos textos
        for row in tqdm(sheet.iter_rows(), desc="Traduzindo textos"):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = traduzir_texto(cell.value)

        # Salva o arquivo traduzido
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
