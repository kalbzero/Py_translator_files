import csv
import os
import json
import time
from concurrent.futures import ThreadPoolExecutor
from googletrans import Translator
from tqdm import tqdm

tradutor = Translator()
cache_arquivo = 'cache_traducoes.json'

# Carregar cache de traduções existente
if os.path.exists(cache_arquivo):
    with open(cache_arquivo, 'r', encoding='utf-8') as f:
        cache_traducoes = json.load(f)
else:
    cache_traducoes = {}

def salvar_cache():
    with open(cache_arquivo, 'w', encoding='utf-8') as f:
        json.dump(cache_traducoes, f, ensure_ascii=False, indent=4)

def is_number(texto):
    # Remove todos os pontos, vírgulas e traços
    texto_limpo = texto.replace('.', '').replace(',', '').replace('-', '')
    return texto_limpo.isdigit()


def is_url(texto):
    return texto.lower().startswith("http")

def traduzir_texto(texto):
    texto_strip = texto.strip()
    
    # Verificar se é um número ou um link
    if is_number(texto_strip) or is_url(texto_strip):
        return texto_strip
    
    if texto_strip in cache_traducoes:
        return cache_traducoes[texto_strip]
    
    tentativas = 0
    max_tentativas = 5
    while tentativas < max_tentativas:
        try:
            traducao = tradutor.translate(texto, src='es', dest='pt').text
            cache_traducoes[texto_strip] = traducao
            return traducao
        except Exception as e:
            print(f"Erro ao traduzir: {e}")
            tentativas += 1
            if tentativas < max_tentativas:
                print(f"Tentando novamente em 5 segundos... ({tentativas}/{max_tentativas})")
                time.sleep(5)
            else:
                print("Máximo de tentativas atingido. Pulando para a próxima palavra.")
                return texto  # Retorna o texto original em caso de falha

def traduzir_arquivo(nome_arquivo):
    print(f"Iniciando tradução do arquivo: {nome_arquivo}")
    
    try:
        with open(nome_arquivo, 'r', encoding='utf-8-sig') as arquivo:
            leitor = csv.reader(arquivo, delimiter=';')
            linhas = list(leitor)
        
        textos_para_traduzir = [celula for linha in linhas for celula in linha if celula.strip()]
        
        with ThreadPoolExecutor() as executor:
            textos_traduzidos = list(tqdm(
                executor.map(traduzir_texto, textos_para_traduzir),
                total=len(textos_para_traduzir),
                desc="Traduzindo textos"
            ))
        
        indice_traduzido = 0
        for i, linha in enumerate(linhas):
            for j, celula in enumerate(linha):
                if celula.strip():
                    # Verifica novamente se é um número ou um link para não alterar o original
                    if is_number(celula.strip()) or is_url(celula.strip()):
                        linhas[i][j] = celula
                    else:
                        linhas[i][j] = textos_traduzidos[indice_traduzido]
                        indice_traduzido += 1
        
        nome_base, extensao = os.path.splitext(nome_arquivo)
        novo_nome = f"{nome_base}_pt{extensao}"
        
        with open(novo_nome, 'w', newline='', encoding='utf-8-sig') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            escritor.writerows(linhas)
        
        print(f"Tradução concluída. Arquivo salvo como: {novo_nome}")
        salvar_cache()
        print(f"Total de traduções únicas: {len(cache_traducoes)}")

    except KeyboardInterrupt:
        print("\nInterrupção detectada! Salvando o progresso e cache...")
        salvar_cache()
        
        nome_base, extensao = os.path.splitext(nome_arquivo)
        novo_nome = f"{nome_base}_parcial_pt{extensao}"
        
        with open(novo_nome, 'w', newline='', encoding='utf-8-sig') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            escritor.writerows(linhas)
        
        print(f"Progresso salvo no arquivo: {novo_nome}")
        print(f"Total de traduções únicas: {len(cache_traducoes)}")
        raise  # Relevanta a exceção para encerrar o programa

# Traduzir apenas o arquivo arquivo_entrada.csv
arquivo_para_traduzir = 'arquivo_entrada.csv'

if os.path.exists(arquivo_para_traduzir):
    traduzir_arquivo(arquivo_para_traduzir)
    print("O arquivo foi traduzido com sucesso!")
else:
    print(f"O arquivo {arquivo_para_traduzir} não foi encontrado.")
