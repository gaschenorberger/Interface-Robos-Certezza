import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tqdm import tqdm  # Para barra de progresso
from pathlib import Path
import pyautogui
import os
import shutil
import subprocess
import sys

# Inicia o Chrome
# "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfile"

# Site de Download
# https://meudanfe.com.br/ver-danfe

def mover_arquivos_xml():

    pasta_origem = r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS"
    pasta_destino = fr"{caminhoPasta}\XML"

    # Garante que o destino existe
    os.makedirs(pasta_destino, exist_ok=True)

    # Percorre os arquivos da pasta de origem
    for arquivo in os.listdir(pasta_origem):
        if arquivo.lower().endswith('.xml'):
            caminho_origem = os.path.join(pasta_origem, arquivo)
            caminho_destino = os.path.join(pasta_destino, arquivo)
            try:
                shutil.move(caminho_origem, caminho_destino)
            except Exception as e:
                print(f'Erro ao mover {arquivo}: {e}')

    print("Todos os arquivos movidos")

def abrir_chrome():
    chrome_options = webdriver.ChromeOptions()

    chrome_options.debugger_address = "localhost:9222"
    # Inicializa o driver conectado ao Chrome aberto
    driver = webdriver.Chrome(options=chrome_options)
    driver.get('https://meudanfe.com.br/')

    print("Chrome conectado com sucesso!")

    return driver

def caminho_script():

    global caminhoPasta
    caminhoPasta = Path(__file__).resolve().parent

caminho_script()

# Caminho para o arquivo de notas
# caminho_arquivo = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\NOTAS.txt'

if len(sys.argv) < 2:
    print("❌ Caminho do arquivo de notas não fornecido.")
    exit()

caminho_arquivo = sys.argv[1]

arquivo_erros = fr'{caminhoPasta}\erros_download.txt'

# Velocidade de execução em segundos (ajuste conforme necessário)
tempo_espera = 1.5  # Tempo de espera entre ações principais

# Definir intervalo de chaves a processar
inicio = 0  # A partir de qual chave começar (0 é a primeira)
fim = None   # Até qual chave processar (None para todas)

# Configurar o Chrome com o caminho de download correto

driver = abrir_chrome()

# Lê as chaves do arquivo de notas
with open(caminho_arquivo, 'r') as file:
    chaves = [linha.strip() for linha in file.readlines()]

# Definir intervalo de chaves a processar
chaves = chaves[inicio:fim]

# Lista para armazenar chaves com erro
chaves_com_erro = []

# Configura espera explícita
wait = WebDriverWait(driver, 10)

erros_seguidos = 0
limite_erros = 5

chaves_pendentes = chaves.copy()

for chave in tqdm(chaves, desc="Processando notas fiscais", unit="nota"):
    try:

        campo_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite a CHAVE DE ACESSO"]')))
        campo_input.clear()
        campo_input.send_keys(chave)
        time.sleep(tempo_espera)

        botao_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Buscar DANFE/XML")]')))
        botao_buscar.click()
        time.sleep(tempo_espera)

        print(f'Chave processada: {chave}')


        botao_baixar_xml = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Baixar XML")]')))
        botao_baixar_xml.click()
        time.sleep(tempo_espera)

        print(f'Download iniciado para a chave: {chave}')

        # Se chegou até aqui, remove dos pendentes
        if chave in chaves_pendentes:
            chaves_pendentes.remove(chave)


        botao_nova_consulta = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]')))
        botao_nova_consulta.click()
        time.sleep(tempo_espera)

        print("Retornando para nova consulta")

        erros_seguidos = 0

    except Exception as e:
        print(f'❌ Erro ao processar chave {chave}: {e}')
        chaves_com_erro.append(chave)

        erros_seguidos += 1

        if erros_seguidos >= limite_erros:
            print(f"Foram detectados {limite_erros} erros seguidos")
            print("A rede foi bloqueada. Troque a VPN e execute novamente o robô")

            # Salvar chaves restantes (incluindo a atual e todas pendentes)
            chaves_restantes = [chave] + chaves_pendentes

            caminho_chaves_restantes = fr'{caminhoPasta}\chaves_restantes.txt'
            with open(caminho_chaves_restantes, 'w') as file_restantes:
                for chave_restante in chaves_restantes:
                    file_restantes.write(chave_restante + '\n')

            print(f"Chaves restantes foram salvas em: {caminho_chaves_restantes}")
            break


# Salvar as chaves com erro em um arquivo de texto
if chaves_com_erro:
    with open(arquivo_erros, 'w') as erro_file:
        for chave_errada in chaves_com_erro:
            erro_file.write(chave_errada + '\n')
    print(f"\nAs seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:")
    print("\n".join(chaves_com_erro))
else:


    print("\nNenhuma chave apresentou erro.")

print("Processamento concluído!")
# mover_arquivos_xml()

# Fechar o navegador manualmente após a execução
