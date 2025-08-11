import customtkinter as ctk
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog
import os
import subprocess
import pygetwindow as gw
import time
import sys
from pathlib import Path
import comtypes
from multiprocessing import Process, Queue
import multiprocessing
import threading

# BIBLIOTECAS ROBÔS

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from screeninfo import get_monitors
from tqdm import tqdm  # Para barra de progresso
import pyautogui        # pip install pyautogui
import random           # (biblioteca padrão do Python)
import mousekey         # pip install mousekey - https://github.com/hansalemaos/mousekey
import pyscreeze        # pip install pyscreeze
import pyperclip
import sys
import shutil
import getpass
import platform




processo_robo = None
caminho_arquivo = ""


""" 
==============================================================================
                ADICIONAR LOGS NA AREA DE TEXTO DA INTERFACE
==============================================================================
"""

log_textbox = None
fila_logs = Queue() 

def monitorar_logs():
    while not fila_logs.empty():
        mensagem = fila_logs.get()

        if mensagem.startswith("SINAL:"):
            comando = mensagem.split("SINAL:")[1]

            if comando == "CONCLUIDO":
                botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
                botao_iniciar.configure(state="disabled")
                botao_chrome.configure(state='normal')
                botao_upload.configure(state='normal')
                botao_vpn.configure(state='normal')
                opcaoXml.configure(state="normal")
                opcaoPdf.configure(state="normal")
                opcaoPdfXml.configure(state="normal")
                atualizar_status_robo("CONCLUIDO")

        else:
            adicionar_log(mensagem)

    root.after(500, monitorar_logs)

def adicionar_log(texto):
    global log_textbox
    if log_textbox is None:
        print(texto)
        return

    
    log_textbox.configure(state="normal")
    log_textbox.insert("end", texto + "\n")
    log_textbox.configure(state="disabled")
    log_textbox.see("end")
    

""" 
==============================================================================
                                FUNÇÕES ROBOS
==============================================================================
"""



def robo_pdf(caminho_arquivo, fila_logs):


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

    def log(msg):
        fila_logs.put(msg)

    def caminho_script():

        global caminhoPasta
        caminhoPasta = Path(__file__).resolve().parent

    def iniciarNavegador():
        chrome_options = webdriver.ChromeOptions()
        chrome_options.debugger_address = "localhost:9222"  # conecta no Chrome já aberto
        driver = webdriver.Chrome(options=chrome_options)
        log("Chrome conectado com sucesso!")
        return driver

    def ativar_pagina(nome_pag): #ATIVA A PAGINA
        window = gw.getWindowsWithTitle(nome_pag)[0]
        window.activate()

    caminho_script()
    
    log("⚠️ NÃO MEXA NO COMPUTADOR DURANTE A EXECUÇÃO DESSE ROBÔ")

    usuario = getpass.getuser()
    caminhoNotas = caminho_arquivo

    tempo_espera = 1.8  # Tempo de espera entre ações principais
    inicio = 0  # A partir de qual chave começar (0 é a primeira)
    fim = None   # Até qual chave processar (None para todas)
    is_child = multiprocessing.current_process().name != "MainProcess"

    navegador = None


    try:

        # Inicializa o driver conectado ao Chrome aberto
        navegador = iniciarNavegador()


        # Lê as chaves do arquivo de notas
        with open(caminhoNotas, 'r') as file:
            chaves = [linha.strip() for linha in file.readlines()]

        # Definir intervalo de chaves a processar
        chaves = chaves[inicio:fim]

        # Lista para armazenar chaves com erro
        chaves_com_erro = []

        # Configura espera explícita
        wait = WebDriverWait(navegador, 10)

        # Barra de progresso para o processamento das chaves
        erros_seguidos = 0
        limite_erros = 5
        chaves_pendentes = chaves.copy()
        total = len(chaves)

        for i, chave in enumerate(tqdm(chaves, desc="Processando notas fiscais", unit="nota", disable=is_child), start=1):
            log(f"Processando nota {i} de {total} - Chave: {chave}\n")
            try:
                ativar_pagina('Consulte e Gere DANFE Online - Emissão Simples de NF-e')

                # CHECAGEM PARA VOLTAR A TELA INICIAL------------------------------------------------------------
                wait_curto = WebDriverWait(navegador, 3)

                try:
                    botao_nova_consulta = wait_curto.until(
                        EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]'))
                    )
                    botao_nova_consulta.click()
                    # print("Botão 'Nova consulta' clicado.")
                except TimeoutException:
                    # print("Botão 'Nova consulta' não encontrado em 3 segundos. Seguindo...")
                    print()
                #-------------------------------------------------------------------------------------------------


                # Insere a chave no campo correto
                campo_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite a CHAVE DE ACESSO"]')))
                campo_input.clear()
                campo_input.send_keys(chave)
                time.sleep(tempo_espera)

                # Clica no botão "Buscar DANFE/XML"
                botao_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Buscar DANFE/XML")]')))
                botao_buscar.click()
                time.sleep(tempo_espera)

                log(f' Chave processada: {chave}\n')

                # Clica no botão "Baixar PDF"
                botao_baixar_xml = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/button[2]')))
                botao_baixar_xml.click()
                time.sleep(tempo_espera)


                for nome in consultar_janelas_abertas():
                    if ".pdf" in nome.lower(): 
                        maximizarJanela(nome)
                        moverJanela(nome)
                        ativar_pagina(nome)
                        break 

                time.sleep(0.5)


                btnDownloadImg = fr'{caminhoPasta}\scripts\download.png'
                procurarImagem(btnDownloadImg, ocorrencia=3)

                pdf_folder = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\PDF"
                os.makedirs(pdf_folder, exist_ok=True)
                
                time.sleep(tempo_espera)
                
                pyperclip.copy(chave)
                caminhoPdf = fr'{pdf_folder}\{chave}'
                log(f"CAMINHO: {caminhoPdf}")

                time.sleep(2.1)

                pyautogui.write(caminhoPdf), time.sleep(0.5)
                pyautogui.press('enter')
                pyautogui.hotkey('ctrl','w')

                log(f'\n✅ Download iniciado para a chave: {chave}')

                time.sleep(2)  

                log(f"\n📁 PDF SALVO NO CAMINHO: {pdf_folder}\n")


                # Clica no botão "Nova consulta" para voltar à página inicial
                botao_nova_consulta = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]')))
                botao_nova_consulta.click()
                time.sleep(tempo_espera)

                log("\nRetornando para nova consulta...")
                erros_seguidos = 0

            except Exception as e:
                log(f'Erro ao processar chave {chave}: {e}')
                chaves_com_erro.append(chave)
                erros_seguidos += 1

                if erros_seguidos >= limite_erros:
                    log(f"❌ Foram detectados {limite_erros} erros seguidos\n")
                    log("⚠️ A rede foi bloqueada. Troque a VPN e execute novamente o robô")

                    chaves_restantes = [chave] + chaves_pendentes
                    caminho_chaves_restantes = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\chaves_restantes.txt"

                    with open(caminho_chaves_restantes, 'w') as file_restantes:
                        for chave_restante in chaves_restantes:
                            file_restantes.write(chave_restante + '\n')

                    log(f"⚠️ Chaves restantes foram salvas em: {caminho_chaves_restantes}")
                    break

        # Salvar as chaves com erro em um arquivo de texto
        arquivo_erros = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"

        if chaves_com_erro:
            with open(arquivo_erros, 'w') as erro_file:
                for chave_errada in chaves_com_erro:
                    erro_file.write(chave_errada + '\n')
            log(f"\nAs seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:")
            log("\n".join(chaves_com_erro), "\n")
        else:
            log("\n✅ Nenhuma chave apresentou erro.\n")

        log("✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")
        fila_logs.put("SINAL:CONCLUIDO")

    except Exception as e:
        log(f"❌ Erro inesperado no robô: {e}")
        fila_logs.put("SINAL:CONCLUIDO")

    finally:
        # FECHAR JANELA AO FINAL DA EXECUÇÃO
        janelas = gw.getWindowsWithTitle('Consulte e Gere DANFE Online - Emissão Simples de NF-e')

        for janela in janelas:
            try:
                janela.close()
            except Exception as e:
                log(f"Erro ao fechar a janela: {e}")

        log("✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")
        fila_logs.put("SINAL:CONCLUIDO")

def robo_xml(caminho_arquivo, fila_logs):

    def log(msg):
        fila_logs.put(msg)

    def iniciarNavegador():
        chrome_options = webdriver.ChromeOptions()
        chrome_options.debugger_address = "localhost:9222"  # conecta no Chrome já aberto
        driver = webdriver.Chrome(options=chrome_options)
        log("Chrome conectado com sucesso!")
        return driver

    def caminho_script():
        global caminhoPasta
        caminhoPasta = Path(__file__).resolve().parent

    def ativar_pagina(nome_pag): #ATIVA A PAGINA
        window = gw.getWindowsWithTitle(nome_pag)[0]
        window.activate()


    caminho_script()

    usuario = getpass.getuser()
    caminhoNotas = caminho_arquivo


    tempo_espera = 1.5
    inicio = 0
    fim = None
    is_child = multiprocessing.current_process().name != "MainProcess"

    navegador = None  # Definir fora do try para garantir que o finally o veja

    try:
        navegador = iniciarNavegador()

        with open(caminhoNotas, 'r') as file:
            chaves = [linha.strip() for linha in file.readlines()]

        chaves = chaves[inicio:fim]
        chaves_com_erro = []
        wait = WebDriverWait(navegador, 20)
        erros_seguidos = 0
        limite_erros = 5
        chaves_pendentes = chaves.copy()
        total = len(chaves)

        for i, chave in enumerate(tqdm(chaves, desc="Processando notas fiscais", unit="nota", disable=is_child), start=1):
            log(f"Processando nota {i} de {total} - Chave: {chave}\n")
            try:

                # CHECAGEM PARA VOLTAR A TELA INICIAL ---------------------------------------------------------
                wait_curto = WebDriverWait(navegador, 3)

                try:
                    botao_nova_consulta = wait_curto.until(
                        EC.element_to_be_clickable((By.XPATH, '//a[contains(@id, "newSearchBtn")]//span[contains(@class,"sm:block")]'))
                    )
                    botao_nova_consulta.click()
                    # print("Botão 'Nova consulta' clicado.")
                except TimeoutException:
                    # print("Botão 'Nova consulta' não encontrado em 3 segundos. Seguindo...")   
                    print()
                #-------------------------------------------------------------------------------------------------
                ativar_pagina('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')
                campo_input = navegador.find_element(By.XPATH, '//input[contains(@id, "searchTxt")]')
                campo_input.clear()
                campo_input.send_keys(chave)
                time.sleep(tempo_espera)

                ativar_pagina('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')
                botao_buscar = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@id,"searchBtn")]')))
                botao_buscar.click()
                time.sleep(13)


                # while navegador.find_elements(By.XPATH, '//div[contains(@class, "jloading")]'):
                #     print("Loading ainda na tela...")
                #     time.sleep(0.2)


                log(f'Chave processada: {chave}\n')

                ativar_pagina('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')
                botao_baixar_xml = navegador.find_element(By.XPATH, '//span[contains(text(),"Baixar XML")]')
                botao_baixar_xml.click()
                time.sleep(tempo_espera)

                log(f'✅ Download iniciado para a chave: {chave}')

                if chave in chaves_pendentes:
                    chaves_pendentes.remove(chave)

                pastaDownloads = fr"C:\Users\{usuario}\Downloads"
                pastaAlternativa = fr"C:\VS_CODE\DOWNLOAD_ARQUIVOS"

                if verificaPastaXML(pastaDownloads):
                    moverXML(pastaDownloads)
                    print("DONWLOAD")
                elif verificaPastaXML(pastaAlternativa):
                    moverXML(pastaAlternativa)
                    print("ALTERNATIVA")
                else:
                    log("⚠️ Não foi possível mover o XML")

                # botao_nova_consulta = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]')))
                # botao_nova_consulta.click()
                # time.sleep(tempo_espera)

                log("Retornando para nova consulta\n")
                erros_seguidos = 0

            except Exception as e:
                log(f'❌ Erro ao processar chave {chave}: {e}')
                chaves_com_erro.append(chave)
                erros_seguidos += 1

                if erros_seguidos >= limite_erros:
                    log(f"❌ Foram detectados {limite_erros} erros seguidos")
                    log("⚠️ A rede foi bloqueada. Troque a VPN e execute novamente o robô\n")

                    chaves_restantes = [chave] + chaves_pendentes
                    caminho_chaves_restantes = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\chaves_restantes.txt"

                    with open(caminho_chaves_restantes, 'w') as file_restantes:
                        for chave_restante in chaves_restantes:
                            file_restantes.write(chave_restante + '\n')

                    log(f"⚠️ Chaves restantes foram salvas em: {caminho_chaves_restantes}")
                    break

        # Salvar erros
        arquivo_erros = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"
        if chaves_com_erro:
            with open(arquivo_erros, 'w') as erro_file:
                for chave_errada in chaves_com_erro:
                    erro_file.write(chave_errada + '\n')
            log(f"\n❌ As seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:\n")
            log("\n".join(chaves_com_erro))
        else:
            log("\n✅ Nenhuma chave apresentou erro.\n")


        fila_logs.put("SINAL:CONCLUIDO")
        # atualizar_status_robo("CONCLUIDO")

    except Exception as e:
        log(f"❌ Erro inesperado no robô: {e}")
        fila_logs.put("SINAL:ERRO")

    finally:
        # FECHAR JANELA AO FINAL DA EXECUÇÃO
        janelas = gw.getWindowsWithTitle('Consulte e Gere DANFE Online - Emissão Simples de NF-e')

        for janela in janelas:
            try:
                janela.close()
            except Exception as e:
                log(f"Erro ao fechar a janela: {e}")


        log("\n✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")
        # atualizar_status_robo("CONCLUIDO")
        fila_logs.put("SINAL:CONCLUIDO")


def robo_pdf_xml(caminho_arquivo, fila_logs):

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

    def log(msg):
        fila_logs.put(msg)

    def caminho_script():

        global caminhoPasta
        caminhoPasta = Path(__file__).resolve().parent

    def iniciarNavegador():
        chrome_options = webdriver.ChromeOptions()
        chrome_options.debugger_address = "localhost:9222"  # conecta no Chrome já aberto
        driver = webdriver.Chrome(options=chrome_options)
        log("Chrome conectado com sucesso!")
        return driver

    def ativar_pagina(nome_pag): #ATIVA A PAGINA
        window = gw.getWindowsWithTitle(nome_pag)[0]
        window.activate()

    caminho_script()
    
    usuario = getpass.getuser()
    caminhoNotas = caminho_arquivo


    tempo_espera = 1.5  # Tempo de espera entre ações principais
    inicio = 0  # A partir de qual chave começar (0 é a primeira)
    fim = None   # Até qual chave processar (None para todas)
    is_child = multiprocessing.current_process().name != "MainProcess"

    navegador = None

    try:
        navegador = iniciarNavegador()

        # Lê as chaves do arquivo de notas
        with open(caminhoNotas, 'r') as file:
            chaves = [linha.strip() for linha in file.readlines()]

        # Definir intervalo de chaves a processar
        chaves = chaves[inicio:fim]

        # Lista para armazenar chaves com erro
        chaves_com_erro = []

        # Configura espera explícita
        wait = WebDriverWait(navegador, 10)

        # Barra de progresso para o processamento das chaves
        erros_seguidos = 0
        limite_erros = 5
        chaves_pendentes = chaves.copy()
        total = len(chaves)

        for i, chave in enumerate(tqdm(chaves, desc="Processando notas fiscais", unit="nota", disable=is_child), start=1):
            try:
                log(f"Processando nota {i} de {total} - Chave: {chave}\n")
                log("⚠️ NÃO MEXA NO COMPUTADOR DURANTE A EXECUÇÃO DESSE ROBÔ")
                ativar_pagina('Consulte e Gere DANFE Online - Emissão Simples de NF-e')

                # CHECAGEM PARA VOLTAR A TELA INICIAL---------------------------------------------------------------
                wait_curto = WebDriverWait(navegador, 3)

                try:
                    botao_nova_consulta = wait_curto.until(
                        EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]'))
                    )
                    botao_nova_consulta.click()
                    # print("Botão 'Nova consulta' clicado.")
                except TimeoutException:
                    # print("Botão 'Nova consulta' não encontrado em 3 segundos. Seguindo...")
                    print()
                #---------------------------------------------------------------------------------------------------

                # Insere a chave no campo correto
                campo_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Digite a CHAVE DE ACESSO"]')))
                campo_input.clear()
                campo_input.send_keys(chave)
                time.sleep(tempo_espera)

                # Clica no botão "Buscar DANFE/XML"
                botao_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Buscar DANFE/XML")]')))
                botao_buscar.click()
                time.sleep(tempo_espera)

                log(f'Chave processada: {chave}')

                # Clica no botão "Baixar XML"
                botao_baixar_xml = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Baixar XML")]')))
                botao_baixar_xml.click()
                time.sleep(tempo_espera)

                #clicar em salvar
                # pyautogui.press('enter')

                log(f'✅ Download iniciado para a chave: {chave}')

                pastaDownloads = fr"C:\Users\{usuario}\Downloads"
                pastaAlternativa = fr"C:\VS_CODE\DOWNLOAD_ARQUIVOS"

                if verificaPastaXML(pastaDownloads):
                    moverXML(pastaDownloads)
                elif verificaPastaXML(pastaAlternativa):
                    moverXML(pastaAlternativa)
                else:
                    log("⚠️ Não foi possível mover o XML")

                # Clica no botão "Baixar PDF"
                botao_baixar_danfe = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/button[2]')))
                botao_baixar_danfe.click()
                time.sleep(tempo_espera)


                for nome in consultar_janelas_abertas():
                    if ".pdf" in nome.lower(): 
                        maximizarJanela(nome)
                        moverJanela(nome)
                        ativar_pagina(nome)
                        break 

                time.sleep(0.5)


                btnDownloadImg = fr'{caminhoPasta}\scripts\download.png'
                procurarImagem(btnDownloadImg, ocorrencia=3)
                
                pdf_folder = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\PDF"
                os.makedirs(pdf_folder, exist_ok=True)
                
                time.sleep(1)
            
                caminhoPdf = fr'{pdf_folder}\{chave}.pdf'
                log(f"CAMINHO: {caminhoPdf}")

                time.sleep(2.1)

                pyautogui.write(caminhoPdf)
                time.sleep(2)
                pyautogui.press('enter')
                pyautogui.hotkey('ctrl','w')

                log(f'\n✅ Download iniciado para a chave: {chave}')

                time.sleep(2)  

                log(f"\n📁 PDF SALVO NO CAMINHO: {pdf_folder}\n")

                # Clica no botão "Nova consulta" para voltar à página inicial
                botao_nova_consulta = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Nova consulta")]')))
                botao_nova_consulta.click()
                time.sleep(tempo_espera)

                log("Retornando para nova consulta...")
                erros_seguidos = 0

            except Exception as e:
                log(f'Erro ao processar chave {chave}: {e}')
                chaves_com_erro.append(chave)
                erros_seguidos += 1

                # SE DER 5 ERROS SEGUIDOS ELE VAI PARAR O ROBÔ
                if erros_seguidos >= limite_erros:
                    log(f"❌ Foram detectados {limite_erros} erros seguidos")
                    log("⚠️ A rede foi bloqueada. Troque a VPN e execute novamente o robô\n")

                    chaves_restantes = [chave] + chaves_pendentes
                    caminho_chaves_restantes = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\chaves_restantes.txt"

                    with open(caminho_chaves_restantes, 'w') as file_restantes:
                        for chave_restante in chaves_restantes:
                            file_restantes.write(chave_restante + '\n')

                    log(f"⚠️ Chaves restantes foram salvas em: {caminho_chaves_restantes}")
                    break

        # Salvar as chaves com erro em um arquivo de texto
        arquivo_erros = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"

        if chaves_com_erro:
            with open(arquivo_erros, 'w') as erro_file:
                for chave_errada in chaves_com_erro:
                    erro_file.write(chave_errada + '\n')
            log(f"\n❌ As seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:")
            log("\n".join(chaves_com_erro))
        else:
            log("\n✅ Nenhuma chave apresentou erro.")


        log("✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME")
        fila_logs.put("SINAL:CONCLUIDO")

    except Exception as e:
        log(f"❌ Erro inesperado no robô: {e}")
        fila_logs.put("SINAL:ERRO")

    finally:
        # FECHAR JANELA AO FINAL DA EXECUÇÃO
        janelas = gw.getWindowsWithTitle('Consulte e Gere DANFE Online - Emissão Simples de NF-e')

        for janela in janelas:
            try:
                janela.close()
            except Exception as e:
                log(f"Erro ao fechar a janela: {e}")

        log("✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")
        fila_logs.put("SINAL:CONCLUIDO")

    # mover_arquivos_xml()




""" 
==============================================================================
                            RELACIONADO A PASTAS
==============================================================================
"""

def moverXML(caminho_origem, timeout=15):
    usuario = getpass.getuser()
    pasta_destino = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\XML"
    os.makedirs(pasta_destino, exist_ok=True)

    tempo_inicial = time.time()
    arquivo_encontrado = None

    while time.time() - tempo_inicial < timeout:
        arquivos_xml = [
            os.path.join(caminho_origem, f)
            for f in os.listdir(caminho_origem)
            if f.lower().endswith('.xml')
        ]

        if arquivos_xml:
            arquivos_xml.sort(key=os.path.getmtime, reverse=True)
            arquivo_mais_recente = arquivos_xml[0]

            if not arquivo_mais_recente.endswith(".crdownload"):
                arquivo_encontrado = arquivo_mais_recente
                break

        time.sleep(0.5)  

    if not arquivo_encontrado:
        print("⏰ Timeout: Nenhum arquivo XML encontrado a tempo.")
        return

    destino_final = os.path.join(pasta_destino, os.path.basename(arquivo_encontrado))

    try:
        shutil.move(arquivo_encontrado, destino_final)
        print(f"✅ XML movido: {os.path.basename(arquivo_encontrado)}")
    except Exception as e:
        adicionar_log(f"❌ Erro ao mover o arquivo: {e}")

def verificaPastaXML(caminho):
    try:
        arquivos = os.listdir(caminho)
    except FileNotFoundError:
        return False
    return any(f.lower().endswith('.xml') for f in arquivos)

def caminho_script():
    global caminhoPasta

    caminhoPasta = Path(__file__).resolve().parent
    print(caminhoPasta)

def resource_path(relative_path):
    """Pega o caminho absoluto, funcionando no .exe e fora dele"""
    try:
        base_path = sys._MEIPASS  # Quando rodando no .exe
    except Exception:
        base_path = os.path.abspath(".")  # Quando rodando em .py
    return os.path.join(base_path, relative_path)




""" 
==============================================================================
                            MONITORAR CHROME ABERTO

Se o chrome estiver aberto ele deixa o botão "ABRIR CHROME" desativado,
caso contrário, deixa o botão ativado.

==============================================================================
"""

def monitorarChrome():

    janelas = consultar_janelas_abertas()
    meuDanfe = 'Consulte e Gere DANFE Online - Emissão Simples de NF-e'

    if meuDanfe in janelas:
        if botao_chrome.cget('state') != 'disabled':
            botao_chrome.configure(state='disabled')
            
    else:
        if botao_chrome.cget('state') != 'normal':
            botao_chrome.configure(state='normal')
            botao_iniciar.configure(state='disabled')
            adicionar_log("⚠️ PARA INICIAR ABRA O CHROME")

    root.after(500, monitorarChrome)
    #Monitora a cada 500 milissegundos


""" 
==============================================================================
                            VERIFICAÇÕES E STATUS

Verifica se está tudo OK para iniciar o Robô e atualiza o status na interface
==============================================================================
"""

def atualizar_status_robo(status):
    status_label.configure(text=f"STATUS: {status}")
    if status == "EXECUTANDO":
        status_label.configure(text_color="red")
        botao_iniciar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
        botao_upload.configure(state="disabled")
        botao_vpn.configure(state="disabled")
        botao_chrome.configure(state="disabled")
        opcaoXml.configure(state="disabled")
        opcaoPdf.configure(state="disabled")
        opcaoPdfXml.configure(state="disabled")

    elif status in ["PARADO", "PRONTO"]:
        status_label.configure(text_color="red")
        botao_iniciar.configure(state="normal", fg_color="#00CC00", hover_color="#009900")
        botao_upload.configure(state="normal")
        botao_vpn.configure(state="normal")
        botao_chrome.configure(state="normal")
        opcaoXml.configure(state="normal")
        opcaoPdf.configure(state="normal")
        opcaoPdfXml.configure(state="normal")

    elif status == "CONCLUIDO":
        status_label.configure(text_color="green")
        botao_iniciar.configure(state="normal", fg_color="#00CC00", hover_color="#009900")
        botao_upload.configure(state="normal")
        botao_vpn.configure(state="normal")
        botao_chrome.configure(state="normal")
        opcaoXml.configure(state="normal")
        opcaoPdf.configure(state="normal")
        opcaoPdfXml.configure(state="normal")

def verificarIniciar():
    tipo_download = var_opcao.get()  

    janelas = consultar_janelas_abertas()
    meuDanfe = 'Consulte e Gere DANFE Online - Emissão Simples de NF-e'


    if tipo_download != "" and meuDanfe in janelas and caminho_arquivo:
        adicionar_log("✅ PRONTO PARA INICIAR\n")
        atualizar_status_robo("PRONTO")

    elif caminho_arquivo and tipo_download == "":
        adicionar_log("⚠️ SELECIONE UMA OPÇÃO DE DOWNLOAD")

    elif tipo_download != "" and meuDanfe in janelas:
        adicionar_log('⚠️ FAÇA UPLOAD DE UM ARQUIVO PARA INICIAR')

    elif tipo_download != "" and caminho_arquivo:
        adicionar_log('⚠️ ABRA O CHROME PARA LIBERAR O BOTAO INICIAR')
    else:
        pass


    if caminho_arquivo and tipo_download in ['PDF', 'XML', 'PDF_XML'] and meuDanfe in janelas:
        botao_iniciar.configure(state='normal')
    else:
        botao_iniciar.configure(state='disabled')

def selecionar_opcao():
    verificarIniciar()


""" 
==============================================================================
                            BOTÕES INTERFACE
==============================================================================
"""


def upload_arquivo():

    global caminho_arquivo

    tipo_download = var_opcao.get()
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

    if caminho_arquivo:
        adicionar_log(f"Arquivo selecionado: {caminho_arquivo}\n")
        arquivo_label.configure(text=f"Arquivo: {os.path.basename(caminho_arquivo)}")
        
    else:
        adicionar_log("⚠️ Nenhum arquivo selecionado.")

    verificarIniciar()
     
def iniciar_robo():
    global processo_robo

    if processo_robo and processo_robo.is_alive():
        adicionar_log("⚠️ Já existe um robô em execução.")
        return

    tipo_download = var_opcao.get()

    if tipo_download == "":
        adicionar_log("⚠️ Selecione PDF ou XML antes de iniciar.")
        return


    # adicionar_log("Robô iniciado com sucesso!")
    atualizar_status_robo("RODANDO")

    if tipo_download == "XML":
        processo_robo = Process(target=robo_xml, args=(caminho_arquivo, fila_logs))
        processo_robo.start()

    elif tipo_download == "PDF":
        processo_robo = Process(target=robo_pdf, args=(caminho_arquivo, fila_logs))
        processo_robo.start()

    elif tipo_download == "PDF_XML":
        processo_robo = Process(target=robo_pdf_xml, args=(caminho_arquivo, fila_logs))
        processo_robo.start()

    else:
        adicionar_log("❌ Tipo de download inválido.")


    try:
        adicionar_log(f"▶️ Robô {tipo_download} iniciado com sucesso!")
        atualizar_status_robo("RODANDO")
        botao_parar.configure(state="normal", fg_color="red", hover_color="#cc0000")

    except Exception as e:
        adicionar_log(f"❌ Erro ao iniciar o robô: {e}")
        atualizar_status_robo("PRONTO")
        botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")

def parar_robo():
    global processo_robo

    if processo_robo and processo_robo.is_alive():
        try:

            processo_robo.terminate()
            processo_robo = None
            adicionar_log("⛔ Execução interrompida pelo usuário.")
            adicionar_log("⚠️ Para iniciar novamente abra o CHROME")
            atualizar_status_robo("PARADO")
            botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
            botao_iniciar.configure(state="disabled")



            usuario = getpass.getuser()
            pastaDownloads = fr"C:\Users\{usuario}\Downloads"
            pastaAlternativa = fr"C:\VS_CODE\DOWNLOAD_ARQUIVOS"

            if verificaPastaXML(pastaDownloads):
                moverXML(pastaDownloads)
                print("DONWLOAD")
            elif verificaPastaXML(pastaAlternativa):
                moverXML(pastaAlternativa)
                print("ALTERNATIVA")
            else:
                adicionar_log("⚠️ Não foi possível mover o XML")
                


            janelas = gw.getWindowsWithTitle('Consulte e Gere DANFE Online - Emissão Simples de NF-e')

            for janela in janelas:
                try:
                    janela.close()
                except Exception as e:
                    adicionar_log(f"Erro ao fechar a janela: {e}")

        except Exception as e:
            adicionar_log(f"❌ Erro ao tentar parar o robô: {e}")
    else:
        adicionar_log("⚠️ Nenhum robô está em execução.")

def abrir_ajuda():
    try:
        caminho_pdf = fr"{caminhoPasta}\Documentacao_Interface_Robos_Certezza_v1.0.0.pdf"

        if platform.system() == "Windows":
            os.startfile(caminho_pdf)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", caminho_pdf], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", caminho_pdf], check=True)

    except Exception as e:
        adicionar_log(f"❌ Erro ao abrir o PDF da ajuda: {e}")

def abrir_vpn():
    caminho = r"C:\Program Files (x86)\Kaspersky Lab\Kaspersky VPN 5.21\ksdeui.exe"

    if os.path.exists(caminho):
        os.startfile(caminho)
        adicionar_log("Abrindo VPN...")
    else:
        adicionar_log("Instale o Kaspersky VPN")

    tentativas = 0
    while not any('Kaspersky VPN' in janela for janela in consultar_janelas_abertas()):
        if tentativas >= 30:
            adicionar_log("❌ Timeout! Não foi possível abrir o VPN.")
            return
        adicionar_log("Esperando o Kaspersky VPN abrir...")
        time.sleep(3)
        tentativas += 1

    adicionar_log("Kaspersky VPN foi aberto!")

def abrir_vpn_thread():
    threading.Thread(target=abrir_vpn, daemon=True).start()

def abrir_chrome():

    opcaoDownload = var_opcao.get()

    if opcaoDownload == "":
        adicionar_log("⚠️ Selecione PDF ou XML antes de iniciar.")
        return

    elif opcaoDownload in ['PDF', 'XML', 'PDF_XML']:
        comando = r'start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Selenium\ChromeTestProfile" --app=https://meudanfe.com.br/'
        subprocess.Popen(comando, shell=True)

        time.sleep(2)

        maximizarJanela("Consulte e Gere DANFE Online - Emissão Simples de NF-e")
        verificarIniciar()

    botao_chrome.configure(state='disabled')




""" 
==============================================================================
                            FUNÇÕES JANELAS
==============================================================================
"""


def consultar_janelas_abertas():
    janelas = gw.getAllTitles()
    return janelas

def maximizarJanela(nomeJanela):
    while True:
        janelas = gw.getWindowsWithTitle(f'{nomeJanela}')

        if janelas:
            janela = janelas[0]
            janela.maximize()
            break
        
def moverJanela(nomeJanela):
    janela = gw.getWindowsWithTitle(nomeJanela)
    if not janela:
        print("Janela não encontrada.")
        return
    janela = janela[0]

    x, y = janela.left, janela.top

    monitores = get_monitors()

    monitor_atual = None
    for m in monitores:
        if m.x <= x < m.x + m.width and m.y <= y < m.y + m.height:
            monitor_atual = m
            break

    monitor_principal = monitores[0]

    if monitor_atual == monitor_principal:
        print("A janela já está no monitor principal.")
    else:
        print("Movendo a janela para o monitor principal.")
        nova_pos_x = monitor_principal.x + 100
        nova_pos_y = monitor_principal.y + 100
        janela.moveTo(nova_pos_x, nova_pos_y)


""" 
==============================================================================
                        CONFIGURAÇÕES TELA PRINCIPAL
==============================================================================
"""


def telaPrincipal():

    global var_opcao
    global opcaoPdf
    global opcaoXml
    global opcaoPdfXml
    global var_opcao
    global log_textbox
    global botao_parar
    global botao_iniciar
    global botao_upload
    global botao_vpn
    global botao_chrome
    global status_label
    global arquivo_label
    global root

    
    caminho_script()

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Painel de Robôs - Certezza")
    root.iconbitmap(resource_path(fr"{caminhoPasta}\img\logoICO.ico"))
    root.geometry("625x600")
    root.resizable(False, False)

    var_opcao = tk.StringVar(value="")

    # Frame principal
    main_frame = ctk.CTkFrame(master=root, fg_color="#D9D9D9", border_color="#F5C400", border_width=2)
    main_frame.pack(padx=10, pady=(10, 0), fill="x")

    # Cabeçalho com status e arquivo
    header_frame = ctk.CTkFrame(master=main_frame, fg_color="transparent")
    header_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=(5, 0), sticky="ew")
    header_frame.grid_columnconfigure(0, weight=0)  # Status
    header_frame.grid_columnconfigure(1, weight=1)  # Arquivo (vai expandir)
    header_frame.grid_columnconfigure(2, weight=0)  # Botão Ajuda

    status_label = ctk.CTkLabel(header_frame, text="STATUS: PARADO", font=("Arial", 12, "bold"), text_color="red")
    status_label.grid(row=0, column=0, padx=(0, 10), sticky="w")

    arquivo_label = ctk.CTkLabel(header_frame, text="Arquivo: Nenhum arquivo selecionado", font=("Arial", 11), text_color="black")
    arquivo_label.grid(row=0, column=1, sticky="w")

    botao_ajuda = ctk.CTkButton(header_frame, text="AJUDA", command=abrir_ajuda,
                                fg_color="#007B8A", hover_color="#006472", text_color="white", width=100)
    botao_ajuda.grid(row=0, column=2, sticky="e", padx=(10, 10), pady=(5, 0))


    # Radiobuttons
    opcoes_frame = ctk.CTkFrame(master=main_frame, fg_color="transparent")
    opcoes_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0), padx=10, sticky="w")

    opcaoPdf = ctk.CTkRadioButton(opcoes_frame, text="PDF", variable=var_opcao, value="PDF", radiobutton_height=20, radiobutton_width=20, command=selecionar_opcao, hover_color="#ffcc00", fg_color="#ffcc00")
    opcaoPdf.pack(side="left")

    opcaoXml = ctk.CTkRadioButton(opcoes_frame, text="XML", variable=var_opcao, value="XML", radiobutton_height=20, radiobutton_width=20, command=selecionar_opcao, hover_color="#ffcc00", fg_color="#ffcc00")
    opcaoXml.pack(side="left")

    opcaoPdfXml = ctk.CTkRadioButton(opcoes_frame, text="PDF E XML", variable=var_opcao, value="PDF_XML", radiobutton_height=20, radiobutton_width=20, command=selecionar_opcao, hover_color="#ffcc00", fg_color="#ffcc00")
    opcaoPdfXml.pack(side="left")

    # try:
    #     logo = resource_path(fr"{caminhoPasta}\img\logo.png")
    #     logoOpen = Image.open(logo).resize((160, 25))
    #     logo_image = ImageTk.PhotoImage(logoOpen)
    #     logo_label = ctk.CTkLabel(main_frame, image=logo_image, text="")
    #     logo_label.place(x=410, y=10)
    # except:
    #     pass

    instrucoes = (
        "1- ESCOLHER O TIPO DE DOWNLOAD (PDF/XML/AMBOS)\n"
        "2- FAZER UPLOAD DE ARQUIVO TXT CONTENDO AS NOTAS UMA EMBAIXO DA OUTRA\n"
        "3- ATIVAR A VPN\n"
        "4- ABRIR CHROME\n"
        "5- CLICAR NO BOTÃO DE INICIAR"
    )
    ctk.CTkLabel(main_frame, text=instrucoes, justify="left", font=("Arial", 11)).grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="w")

    botoes_linha_frame = ctk.CTkFrame(master=main_frame, fg_color="transparent")
    botoes_linha_frame.grid(row=3, column=0, columnspan=2, padx=(10), pady=(5, 10), sticky="w")

    botao_upload = ctk.CTkButton(botoes_linha_frame, text="1- UPLOAD", command=upload_arquivo, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    botao_upload.grid(row=0, column=0, padx=(0, 10))

    botao_vpn = ctk.CTkButton(botoes_linha_frame, text="2- ABRIR VPN", command=abrir_vpn_thread, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    botao_vpn.grid(row=0, column=1, padx=(0, 10))

    botao_chrome = ctk.CTkButton(botoes_linha_frame, text="3- ABRIR CHROME", command=abrir_chrome, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    botao_chrome.grid(row=0, column=2, padx=(0, 10))

    botao_iniciar = ctk.CTkButton(botoes_linha_frame, text="4- INICIAR", command=iniciar_robo, fg_color="#00CC00", hover_color="#009900", text_color="black", width=100, state="disabled")
    botao_iniciar.grid(row=0, column=3, padx=(0, 10))

    botao_parar = ctk.CTkButton(botoes_linha_frame, text="PARAR EXECUÇÃO", command=parar_robo, fg_color="#A9A9A9", hover_color="#808080", text_color="white", width=120, state="disabled")
    botao_parar.grid(row=0, column=4, padx=(0, 10))


    log_frame = ctk.CTkFrame(master=root, fg_color="#EDEDED", border_color="#F5C400", border_width=2)
    log_frame.pack(padx=10, pady=(0, 10), fill="both", expand=True)

    log_textbox = tk.Text(log_frame, bg="#EDEDED", font=("Courier New", 10), relief="flat", state="disabled", wrap="word", height=15)
    log_textbox.pack(padx=10, pady=10, fill="both", expand=True)

    adicionar_log("⚠️ CASO HOUVER ALGUMA DÚVIDA EM RELAÇÃO AS FUNCIONALIDADES DO SISTEMA, CLIQUE NO BOTÃO 'AJUDA'\n")

    monitorar_logs()
    monitorarChrome()
    root.mainloop()

    return caminho_arquivo



if __name__ == "__main__":

    from multiprocessing import freeze_support
    freeze_support()
    root = telaPrincipal()










