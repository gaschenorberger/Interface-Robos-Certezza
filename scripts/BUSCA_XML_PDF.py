import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tqdm import tqdm   # Para barra de progresso

import pyautogui        # pip install pyautogui

import random           # (biblioteca padrão do Python)
import mousekey         # pip install mousekey - https://github.com/hansalemaos/mousekey
import pyscreeze        # pip install pyscreeze
import pyperclip
import sys

def procurarImagem(nome_arquivo, confidence=0.6, region=None, maxTentativas=60, horizontal=0, vertical=0, dx=0, dy=0, acao='clicar', clicks=1, ocorrencia=1, delay_tentativa=1):
    mkey = mousekey.MouseKey()

    def click(x, y):
        pyautogui.click(x, y)

    def doubleClick(x, y):
        pyautogui.doubleClick(x, y)

    def coordenada(x, y):
        print(f'Coordenadas da imagem: ({x}, {y})')

    def moveMouse(x, y, variationx=(-5, 5), variationy=(-5, 5), up_down=(0.2, 0.3), min_variation=-10, max_variation=10, use_every=4, sleeptime=(0.009, 0.019), linear=90):
        mkey.left_click_xy_natural(
            int(x) - random.randint(*variationx),
            int(y) - random.randint(*variationy),
            delay=random.uniform(*up_down),
            min_variation=min_variation,
            max_variation=max_variation,
            use_every=use_every,
            sleeptime=sleeptime,
            print_coords=True,
            percent=linear,
        )

    def clickDrag(x, y, dx, dy):
        pyautogui.moveTo(x, y)
        pyautogui.mouseDown()
        pyautogui.moveTo(x + dx, y + dy, duration=0.5)
        pyautogui.mouseUp()


    acoesValidas = ['clicar', 'mover clicar', 'clicar arrastar']

    if acao not in acoesValidas:
        raise ValueError(f"Ação inválida: '{acao}'. Escolha entre {acoesValidas}.")

    tentativas = 0
    while tentativas < maxTentativas:
        tentativas += 1
        try:
            imag = list(pyautogui.locateAllOnScreen(nome_arquivo, confidence=confidence, region=region))
            
            if imag:
                if len(imag) >= ocorrencia:
                    img = imag[ocorrencia - 1]  
                    x, y = pyautogui.center(img) 
                    x += horizontal
                    y += vertical

                    match acao:
                        case 'clicar':
                            match clicks:
                                case 0:
                                    coordenada(x, y)
                                case 1:
                                    click(x, y)
                                case 2:
                                    doubleClick(x, y)

                        case 'mover clicar':
                            moveMouse(x, y)
                        
                        case 'clicar arrastar':
                            clickDrag(x, y, dx, dy)

                    return True
                else:
                    print(f'A ocorrência {ocorrencia} não foi encontrada.')
                    return False

        except pyscreeze.ImageNotFoundException:
            pass
        time.sleep(delay_tentativa)

    print(f'Imagem não encontrada após {maxTentativas} tentativas.')
    return False

# Inicia o Chrome
# "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfile"

# Site de Download
# https://meudanfe.com.br/ver-danfe


def mover_arquivos_xml():

    pasta_origem = r"C:\VS_CODE\DOWNLOAD'S ARQUIVOS"
    pasta_destino = r"C:\VS_CODE\BUSCA_XML\PDF"

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

# Rodar Erros

# caminho_arquivo = r"C:\VS_CODE\BUSCA_XML\erros_download2.txt"

# Caminho para o arquivo de notas
if len(sys.argv) < 2:
    print("❌ Caminho do arquivo de notas não fornecido.")
    exit()

caminho_arquivo = sys.argv[1]

# caminho_arquivo = r"C:\VS_CODE\BUSCA_XML\XML'S J MONTE\filial 0011.txt"
arquivo_erros = r'C:\VS_CODE\BUSCA_XML\erros_download.txt'

# Velocidade de execução em segundos (ajuste conforme necessário)
tempo_espera = 1.8  # Tempo de espera entre ações principais

# Definir intervalo de chaves a processar
inicio = 0  # A partir de qual chave começar (0 é a primeira)
fim = None   # Até qual chave processar (None para todas)

# Conectar ao Chrome já aberto (lembre-se de abrir com --remote-debugging-port=9222)
chrome_options = webdriver.ChromeOptions()
chrome_options.debugger_address = "localhost:9222"

# Inicializa o driver conectado ao Chrome aberto
driver = webdriver.Chrome(options=chrome_options)

print("Chrome conectado com sucesso!")

# Lê as chaves do arquivo de notas
with open(caminho_arquivo, 'r') as file:
    chaves = [linha.strip() for linha in file.readlines()]

# Definir intervalo de chaves a processar
chaves = chaves[inicio:fim]

# Lista para armazenar chaves com erro
chaves_com_erro = []

# Configura espera explícita
wait = WebDriverWait(driver, 10)

# Barra de progresso para o processamento das chaves
for chave in tqdm(chaves, desc="Processando notas fiscais", unit="nota"):
    try:
        # Insere a chave no campo correto
        campo_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite a CHAVE DE ACESSO"]')))
        campo_input.clear()
        campo_input.send_keys(chave)
        time.sleep(tempo_espera)

        # Clica no botão "Buscar DANFE/XML"
        botao_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Buscar DANFE/XML")]')))
        botao_buscar.click()
        time.sleep(tempo_espera)

        print(f' Chave processada: {chave}')

        # Clica no botão "Baixar PDF"
        botao_baixar_xml = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/button[2]')))
        botao_baixar_xml.click()
        time.sleep(tempo_espera)


        procurarImagem(r"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\download.png")
        
        time.sleep(tempo_espera)
        pyperclip.copy(chave)
        caminhoPdf = fr'C:\VS_CODE\BUSCA_XML\PDF\{chave}'
        pyautogui.write(caminhoPdf), time.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl','w')


        print(f'Download iniciado para a chave: {chave}')

        # Clica no botão "Nova consulta" para voltar à página inicial
        botao_nova_consulta = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]')))
        botao_nova_consulta.click()
        time.sleep(tempo_espera)

        print("Retornando para nova consulta...")

    except Exception as e:
        print(f'Erro ao processar chave {chave}: {e}')
        chaves_com_erro.append(chave)

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
mover_arquivos_xml()


