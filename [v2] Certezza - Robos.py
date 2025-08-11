import customtkinter as ctk
from tkinter import *
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import os
import subprocess
import pygetwindow as gw
import time
import sys
from pathlib import Path
import comtypes
from multiprocessing import Process, Queue
from multiprocessing import Manager
import multiprocessing
import threading
from ttkbootstrap import Style
from ttkbootstrap.widgets import Notebook
from tkinter import messagebox

# BIBLIOTECAS ROBÔS

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from screeninfo import get_monitors
from datetime import datetime
from tqdm import tqdm  # Para barra de progresso
import undetected_chromedriver as uc
import pyautogui        # pip install pyautogui
import random           # (biblioteca padrão do Python)
import mousekey         # pip install mousekey - https://github.com/hansalemaos/mousekey
import pyscreeze        # pip install pyscreeze
import pyperclip
import sys
import shutil
import getpass
import platform
import traceback



""" 
|===================================================================================================================|
|                                                    INTERFACE XML                                                  |
|===================================================================================================================|
"""


processo_robo = None
caminho_arquivo = ""

usuario = getpass.getuser()
pasta_arquivos = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS"
os.makedirs(pasta_arquivos, exist_ok=True)



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
                # botao_chrome.configure(state='normal')
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

#---MEU DANFE---

def robo_pdf(caminho_arquivo, fila_logs, chaves_com_erro):

    def log(msg):
        fila_logs.put(msg)

    usuario = getpass.getuser()

    def iniciarNavegador():
        # chrome_options = webdriver.ChromeOptions()
        # chrome_options.debugger_address = "localhost:9222"  # conecta no Chrome já aberto
        # # driver = webdriver.Chrome(options=chrome_options)
        # log("Chrome conectado com sucesso!")

        print("[LOG] Iniciando navegador...")

        pasta_destino = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\PDF"
        os.makedirs(pasta_destino, exist_ok=True)

        options = uc.ChromeOptions()
        prefs = {
            "download.default_directory": pasta_destino,  
            "download.prompt_for_download": False,         # não perguntar onde salvar
            "directory_upgrade": True,                     # substituir diretórios existentes
            "safebrowsing.enabled": True,                    # habilitar safe browsing (pode ajudar em downloads de arquivos)
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1
        }

        options = uc.ChromeOptions()
        options.add_argument("https://meudanfe.com.br/")
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')

        options.add_experimental_option("prefs", prefs)

        driver = uc.Chrome(options=options, headless=False)

        time.sleep(1)
        return driver

    def caminho_script():
        global caminhoPasta
        caminhoPasta = Path(__file__).resolve().parent
    

    caminho_script()

    caminhoNotas = caminho_arquivo

    tempo_espera = 1.8  # Tempo de espera entre ações principais
    inicio = 0  # A partir de qual chave começar (0 é a primeira)
    fim = None   # Até qual chave processar (None para todas)
    is_child = multiprocessing.current_process().name != "MainProcess"

    navegador = None


    try:

        try:
            navegador = iniciarNavegador()
            print("[LOG] Navegador retornado pela função iniciarNavegador()")
        except Exception as e:
            print(f"[ERRO] Deu ruim logo após iniciar o navegador: {e}")


        with open(caminhoNotas, 'r') as file:
            chaves = [linha.strip() for linha in file.readlines()]

        chaves = chaves[inicio:fim]
        wait = WebDriverWait(navegador, 40)
        erros_seguidos = 0
        limite_erros = 5
        chaves_pendentes = chaves.copy()
        total = len(chaves)

        for i, chave in enumerate(tqdm(chaves, desc="Processando notas fiscais", unit="nota", disable=is_child), start=1):
            log(f"Processando nota {i} de {total} - Chave: {chave}\n")
            try:

                campo_input = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, '//input[contains(@id, "searchTxt")]')
                ))
                campo_input.clear()
                campo_input.send_keys(chave)

                time.sleep(tempo_espera)


                botao_buscar = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//a[contains(@id,"searchBtn")]')
                ))
                botao_buscar.click()

                time.sleep(1)


                try:
                    if navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]"):

                        adicionar_log("⚠️ Esperando Carregamento")
                        divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                        while divLoading:
                            time.sleep(1)
                            divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                            if not divLoading:
                                break
                    else:
                        pass
                except NoSuchElementException:
                    pass


                parar, erros_seguidos = verificar_erro(navegador, chave, chaves_pendentes, chaves_com_erro, erros_seguidos, limite_erros, usuario, log)

                if parar:
                    continue
                else:
                    print("✅ Nenhum erro detectado.")
                    log(f'Chave processada: {chave}\n')


                    botao_baixar_danfe = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//span[contains(text(),"Baixar DANFE")]')
                    ))                
                    botao_baixar_danfe.click()

                    
                    time.sleep(tempo_espera)
                    log(f'✅ Download iniciado para a chave: {chave}')

                    if chave in chaves_pendentes:
                        chaves_pendentes.remove(chave)

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

        compararPastaLista("pdf", caminhoNotas, log)

    except Exception as e:
        log(f"❌ Erro inesperado no robô: {e}")
        fila_logs.put("SINAL:ERRO")

    finally:
        # Salvar erros
        if chaves_com_erro:
            arquivo_erros = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"
            with open(arquivo_erros, 'w') as erro_file:
                for chave_errada in chaves_com_erro:
                    erro_file.write(chave_errada + '\n')
            log(f"\n❌ As seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:\n")
            log("\n".join(chaves_com_erro))
        else:
            log("\n✅ Nenhuma chave apresentou erro.\n")

        # FECHAR JANELA AO FINAL DA EXECUÇÃO
        janelas = gw.getWindowsWithTitle('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')
        for janela in janelas:
            try:
                janela.close()
            except Exception as e:
                log(f"Erro ao fechar a janela: {e}")

        log("\n✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")

        

        fila_logs.put("SINAL:CONCLUIDO")

def robo_xml(caminho_arquivo, fila_logs, chaves_com_erro):

    def log(msg):
        fila_logs.put(msg)

    usuario = getpass.getuser()

    def iniciarNavegador():
        # chrome_options = webdriver.ChromeOptions()
        # chrome_options.debugger_address = "localhost:9222"  # conecta no Chrome já aberto
        # # driver = webdriver.Chrome(options=chrome_options)
        # log("Chrome conectado com sucesso!")

        print("[LOG] Iniciando navegador...")

        pasta_destino = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\XML"
        os.makedirs(pasta_destino, exist_ok=True)

        options = uc.ChromeOptions()
        prefs = {
            "download.default_directory": pasta_destino,  
            "download.prompt_for_download": False,         # não perguntar onde salvar
            "directory_upgrade": True,                     # substituir diretórios existentes
            "safebrowsing.enabled": True,                    # habilitar safe browsing (pode ajudar em downloads de arquivos)
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1
        }

        options = uc.ChromeOptions()
        options.add_argument("https://meudanfe.com.br/")
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')

        options.add_experimental_option("prefs", prefs)

        driver = uc.Chrome(options=options, headless=False)

        time.sleep(1)
        return driver

    def caminho_script():
        global caminhoPasta
        caminhoPasta = Path(__file__).resolve().parent
    

    caminho_script()

    caminhoNotas = caminho_arquivo


    tempo_espera = 1.5
    inicio = 0
    fim = None
    is_child = multiprocessing.current_process().name != "MainProcess"

    navegador = None  # Definir fora do try para garantir que o finally o veja

    try:

        try:
            navegador = iniciarNavegador()
            print("[LOG] Navegador retornado pela função iniciarNavegador()")
        except Exception as e:
            print(f"[ERRO] Deu ruim logo após iniciar o navegador: {e}")


        with open(caminhoNotas, 'r') as file:
            chaves = [linha.strip() for linha in file.readlines()]

        chaves = chaves[inicio:fim]
        # chaves_com_erro = []
        wait = WebDriverWait(navegador, 40)
        erros_seguidos = 0
        limite_erros = 5
        chaves_pendentes = chaves.copy()
        total = len(chaves)

        for i, chave in enumerate(tqdm(chaves, desc="Processando notas fiscais", unit="nota", disable=is_child), start=1):

            log(f"Processando nota {i} de {total} - Chave: {chave}\n")

            try:

                campo_input = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, '//input[contains(@id, "searchTxt")]')
                ))
                campo_input.clear()
                campo_input.send_keys(chave)

                time.sleep(tempo_espera)


                botao_buscar = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//a[contains(@id,"searchBtn")]')
                ))
                botao_buscar.click()

                time.sleep(1)


                try:
                    if navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]"):

                        log("⚠️ Esperando Carregamento")
                        divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                        while divLoading:
                            time.sleep(1)
                            divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                            if not divLoading:
                                break
                    else:
                        pass
                except NoSuchElementException:
                    pass

                
                parar, erros_seguidos = verificar_erro(navegador, chave, chaves_pendentes, chaves_com_erro, erros_seguidos, limite_erros, usuario, log)

                if parar:
                    continue
                else:
                    print("✅ Nenhum erro detectado.")
                    log(f'Chave processada: {chave}\n')


                    botao_baixar_xml = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//span[contains(text(),"Baixar XML")]')
                    ))                
                    botao_baixar_xml.click()

                    time.sleep(tempo_espera)
                    

                    log(f'✅ Download iniciado para a chave: {chave}')

                    if chave in chaves_pendentes:
                        chaves_pendentes.remove(chave)

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

        compararPastaLista("xml", caminhoNotas, log)

    except Exception as e:
        log(f"❌ Erro inesperado no robô: {e}")
        fila_logs.put("SINAL:ERRO")

    finally:
        # Salvar erros
        if chaves_com_erro:
            arquivo_erros = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"
            with open(arquivo_erros, 'w') as erro_file:
                for chave_errada in chaves_com_erro:
                    erro_file.write(chave_errada + '\n')
            log(f"\n❌ As seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:\n")
            log("\n".join(chaves_com_erro))
        else:
            log("\n✅ Nenhuma chave apresentou erro.\n")

        # FECHAR JANELA AO FINAL DA EXECUÇÃO
        janelas = gw.getWindowsWithTitle('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')
        for janela in janelas:
            try:
                janela.close()
            except Exception as e:
                log(f"Erro ao fechar a janela: {e}")

        log("\n✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")

        fila_logs.put("SINAL:CONCLUIDO")

def robo_pdf_xml(caminho_arquivo, fila_logs, chaves_com_erro):

    def log(msg):
        fila_logs.put(msg)

    usuario = getpass.getuser()

    def iniciarNavegador():
        # chrome_options = webdriver.ChromeOptions()
        # chrome_options.debugger_address = "localhost:9222"  # conecta no Chrome já aberto
        # # driver = webdriver.Chrome(options=chrome_options)
        # log("Chrome conectado com sucesso!")

        print("[LOG] Iniciando navegador...")

        pasta_destino = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\PDF_XML"
        os.makedirs(pasta_destino, exist_ok=True)

        options = uc.ChromeOptions()
        prefs = {
            "download.default_directory": pasta_destino,  
            "download.prompt_for_download": False,         # não perguntar onde salvar
            "directory_upgrade": True,                     # substituir diretórios existentes
            "safebrowsing.enabled": True,                    # habilitar safe browsing (pode ajudar em downloads de arquivos)
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1
        }

        options = uc.ChromeOptions()
        options.add_argument("https://meudanfe.com.br/")
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')

        options.add_experimental_option("prefs", prefs)

        driver = uc.Chrome(options=options, headless=False)

        time.sleep(1)
        return driver

    def caminho_script():
        global caminhoPasta
        caminhoPasta = Path(__file__).resolve().parent
    

    caminho_script()

    caminhoNotas = caminho_arquivo


    tempo_espera = 1.5
    inicio = 0
    fim = None
    is_child = multiprocessing.current_process().name != "MainProcess"

    navegador = None  # Definir fora do try para garantir que o finally o veja

    try:

        try:
            navegador = iniciarNavegador()
            print("[LOG] Navegador retornado pela função iniciarNavegador()")
        except Exception as e:
            print(f"[ERRO] Deu ruim logo após iniciar o navegador: {e}")


        with open(caminhoNotas, 'r') as file:
            chaves = [linha.strip() for linha in file.readlines()]

        chaves = chaves[inicio:fim]
        wait = WebDriverWait(navegador, 40)
        erros_seguidos = 0
        limite_erros = 5
        chaves_pendentes = chaves.copy()
        total = len(chaves)

        for i, chave in enumerate(tqdm(chaves, desc="Processando notas fiscais", unit="nota", disable=is_child), start=1):
            log(f"Processando nota {i} de {total} - Chave: {chave}\n")
            try:

                #------------BAIXAR XML-----------------

                campo_input = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, '//input[contains(@id, "searchTxt")]')
                ))
                campo_input.clear()
                campo_input.send_keys(chave)

                time.sleep(tempo_espera)


                botao_buscar = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//a[contains(@id,"searchBtn")]')
                ))
                botao_buscar.click()

                time.sleep(1)


                try:
                    if navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]"):

                        log("⚠️ Esperando Carregamento")
                        divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                        while divLoading:
                            time.sleep(1)
                            divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                            if not divLoading:
                                break
                    else:
                        pass
                except NoSuchElementException:
                    pass


                parar, erros_seguidos = verificar_erro(navegador, chave, chaves_pendentes, chaves_com_erro, erros_seguidos, limite_erros, usuario, log)

                if parar:
                    continue

                else:
                    print("✅ Nenhum erro detectado")
                    log(f'Chave processada: {chave}\n')


                    botao_baixar_xml = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//span[contains(text(),"Baixar XML")]')
                    ))                
                    botao_baixar_xml.click()

                    time.sleep(3)
                    

                    log(f'✅ Download XML iniciado para a chave: {chave}')

                    if chave in chaves_pendentes:
                        chaves_pendentes.remove(chave)

                    #------------BAIXAR DANFE-----------------

                    botao_buscar = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//a[contains(@id,"searchBtn")]')
                    ))
                    botao_buscar.click()

                    time.sleep(1)


                    try:
                        if navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]"):

                            log("⚠️ Esperando Carregamento")
                            divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                            while divLoading:
                                time.sleep(1)
                                divLoading = navegador.find_element(By.XPATH, "//div[contains(@class, 'jloading')]")

                                if not divLoading:
                                    break
                        else:
                            pass
                    except NoSuchElementException:
                        pass



                    parar, erros_seguidos = verificar_erro(navegador, chave, chaves_pendentes, chaves_com_erro, erros_seguidos, limite_erros, usuario, log)

                    if parar:
                        continue
                    
                    else:
                        print("✅ Nenhum erro detectado, continuando com DANFE")
                        
                        botao_baixar_danfe = wait.until(EC.element_to_be_clickable(
                            (By.XPATH, '//span[contains(text(),"Baixar DANFE")]')
                        ))                
                        botao_baixar_danfe.click()

                        time.sleep(tempo_espera)
                        log(f'✅ Download DANFE iniciado para a chave: {chave}\n')

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

    except Exception as e:
        log(f"❌ Erro inesperado no robô: {e}")
        fila_logs.put("SINAL:ERRO")

    finally:
        # Salvar erros
        if chaves_com_erro:
            arquivo_erros = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"
            with open(arquivo_erros, 'w') as erro_file:
                for chave_errada in chaves_com_erro:
                    erro_file.write(chave_errada + '\n')
            log(f"\n❌ As seguintes chaves apresentaram erro e foram salvas em {arquivo_erros}:\n")
            log("\n".join(chaves_com_erro))
        else:
            log("\n✅ Nenhuma chave apresentou erro.\n")

        # FECHAR JANELA AO FINAL DA EXECUÇÃO
        janelas = gw.getWindowsWithTitle('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')
        for janela in janelas:
            try:
                janela.close()
            except Exception as e:
                log(f"Erro ao fechar a janela: {e}")

        log("\n✅ Processamento concluído!")
        log("⚠️ Para iniciar novamente abra o CHROME\n")

        compararPastaLista("pdf_xml", caminhoNotas)

        fila_logs.put("SINAL:CONCLUIDO")


""" 
==============================================================================
                            RELACIONADO A PASTAS
==============================================================================
"""

def caminho_script():
    global caminhoPasta

    caminhoPasta = Path(__file__).resolve().parent
    # print(caminhoPasta)

def resource_path(relative_path):
    """Pega o caminho absoluto, funcionando no .exe e fora dele"""
    try:
        base_path = sys._MEIPASS  # Quando rodando no .exe
    except Exception:
        base_path = os.path.abspath(".")  # Quando rodando em .py
    return os.path.join(base_path, relative_path)

def compararPastaLista(tipoDownload, caminhoChaves, log):
    usuario = os.getlogin()

    # Caminhos
    caminho_lista = caminhoChaves

    if tipoDownload.lower() == 'xml':
        pasta_downloads = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\XML"
    elif tipoDownload.lower() == 'pdf':
        pasta_downloads = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\PDF"
    elif tipoDownload.lower() == 'pdf_xml':
        pasta_downloads = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\PDF_XML"

    arquivo_saida = fr"C:\Users\{usuario}\Desktop\DOWNLOAD'S ROBOS\faltantes.txt"

    def ler_lista_baixar(caminho):
        with open(caminho, 'r', encoding='utf-8') as f:
            return [linha.strip() for linha in f if linha.strip()]

    def listar_arquivos_baixados_filtrados(pasta, lista):
        encontrados = set()
        for nome in os.listdir(pasta):
            if nome.lower().endswith('.xml'):
                nome_base = os.path.splitext(nome)[0].replace("NFE-", "")
                if nome_base in lista:
                    encontrados.add(nome_base)
        return encontrados

    def salvar_faltantes(lista, caminho):
        with open(caminho, 'w', encoding='utf-8') as f:
            for item in lista:
                f.write(item + '\n')
        log(f"\n📝 Arquivo 'faltantes.txt' salvo com {len(lista)} itens.")

    def main():
        lista_baixar = ler_lista_baixar(caminho_lista)
        baixados_filtrados = listar_arquivos_baixados_filtrados(pasta_downloads, lista_baixar)

        faltantes = [item for item in lista_baixar if item not in baixados_filtrados]

        log(f"🔎 Total na lista: {len(lista_baixar)}")
        log(f"✅ Arquivos XML encontrados (da lista): {len(baixados_filtrados)}")
        log(f"⚠️ Arquivos faltantes: {len(faltantes)}")

        if faltantes:
            salvar_faltantes(faltantes, arquivo_saida)
        else:
            log("🎉 Todos os arquivos foram encontrados.")

    main()


""" 
==============================================================================
                            MONITORAR CHROME ABERTO

Se o chrome estiver aberto ele deixa o botão "ABRIR CHROME" desativado,
caso contrário, deixa o botão ativado.

==============================================================================
"""

def monitorarChrome():

    # janelas = consultar_janelas_abertas()
    # meuDanfe = 'Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF'

    # if meuDanfe in janelas:
    #     if botao_chrome.cget('state') != 'disabled':
    #         botao_chrome.configure(state='disabled')
            
    # else:
    #     if botao_chrome.cget('state') != 'normal':
    #         # botao_chrome.configure(state='normal')
    #         botao_iniciar.configure(state='disabled')
    #         # adicionar_log("⚠️ PARA INICIAR ABRA O CHROME")



    root.after(500, monitorarChrome)
    #Monitora a cada 500 milissegundos

""" 
==============================================================================
                            VERIFICAÇÕES E STATUS

Verifica se está tudo OK para iniciar o Robô e atualiza o status na interface
==============================================================================
"""

def verificar_erro(navegador, chave, chaves_pendentes, chaves_com_erro,
                   erros_seguidos, limite_erros, usuario, log):
    try:
        alert_div = navegador.find_element(By.ID, "alertDiv")
        style = alert_div.get_attribute("style")

        if style == "":
            print("🚨 Erro detectado!")
            log(f'❌ Erro ao processar chave {chave}: Erro detectado pela div | Passando para próxima CHAVE \n')


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

            return True, erros_seguidos  # ← sempre retorna True se style == ""
    except:
        pass

    return False, erros_seguidos

def monitorar_pasta():
    if os.path.exists(pasta_arquivos):
        botao_pasta.configure(state="normal")
    else:
        botao_pasta.configure(state="disabled")

    root.after(500, monitorar_pasta)

def atualizar_status_robo(status):
    status_label.configure(text=f"STATUS: {status}")
    if status == "EXECUTANDO":
        status_label.configure(text_color="red")
        botao_iniciar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
        botao_upload.configure(state="disabled")
        botao_vpn.configure(state="disabled")
        # botao_chrome.configure(state="disabled")
        opcaoXml.configure(state="disabled")
        opcaoPdf.configure(state="disabled")
        opcaoPdfXml.configure(state="disabled")
        # botao_resetar.configure(state="disabled")

    elif status in ["PARADO", "PRONTO"]:
        status_label.configure(text_color="red")
        botao_iniciar.configure(state="normal", fg_color="#00CC00", hover_color="#009900")
        botao_upload.configure(state="normal")
        botao_vpn.configure(state="normal")
        opcaoXml.configure(state="normal")
        opcaoPdf.configure(state="normal")
        opcaoPdfXml.configure(state="normal")
        # botao_resetar.configure(state="normal")
        # botao_chrome.configure(state="normal")

    elif status == "CONCLUIDO":
        status_label.configure(text_color="green")
        botao_iniciar.configure(state="normal", fg_color="#00CC00", hover_color="#009900")
        botao_upload.configure(state="normal")
        botao_vpn.configure(state="normal")
        # botao_resetar.configure(state="normal")
        opcaoXml.configure(state="normal")
        opcaoPdf.configure(state="normal")
        opcaoPdfXml.configure(state="normal")
        # botao_chrome.configure(state="normal")

def verificarIniciar():
    tipo_download = var_opcao.get()  

    janelas = consultar_janelas_abertas()
    meuDanfe = 'Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF'


    if tipo_download != "" and caminho_arquivo:
        adicionar_log("✅ PRONTO PARA INICIAR\n")
        atualizar_status_robo("PRONTO")

    elif tipo_download != "" and caminho_arquivo == "":
        adicionar_log("⚠️ FAÇA UPLOAD PARA INICIAR\n")

    elif caminho_arquivo and tipo_download == "":
        adicionar_log("⚠️ SELECIONE UMA OPÇÃO DE DOWNLOAD\n")


    # and meuDanfe in janelas
    if caminho_arquivo and tipo_download in ['PDF', 'XML', 'PDF_XML'] :
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
    import os

    global caminho_arquivo

    tipo_download = var_opcao.get()
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

    if caminho_arquivo:
        adicionar_log(f"Arquivo selecionado: {caminho_arquivo}\n")
        arquivo_label.configure(text=f"Arquivo: {os.path.basename(caminho_arquivo)}")

        # ✅ Limpar linhas em branco do próprio arquivo
        try:
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                linhas = f.readlines()

            # Remove linhas que estão vazias ou só com espaços
            linhas_limpa = [linha for linha in linhas if linha.strip() != ""]

            # Se removeu alguma linha, sobrescreve o arquivo
            if len(linhas) != len(linhas_limpa):
                with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                    f.writelines(linhas_limpa)
                adicionar_log("🧹 Linhas em branco foram removidas do arquivo.\n")

        except Exception as e:
            adicionar_log(f"❌ Erro ao limpar o arquivo: {e}\n")

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

    messagebox.showinfo(
        "Atenção",
        "NÃO MEXA NO COMPUTADOR DURANTE A EXECUÇÃO DO ROBÔ\nClique em OK para continuar"
    )

    if tipo_download == "XML":
        processo_robo = Process(target=robo_xml, args=(caminho_arquivo, fila_logs, chaves_com_erro))
        processo_robo.start()

    elif tipo_download == "PDF":
        processo_robo = Process(target=robo_pdf, args=(caminho_arquivo, fila_logs, chaves_com_erro))
        processo_robo.start()

    elif tipo_download == "PDF_XML":
        processo_robo = Process(target=robo_pdf_xml, args=(caminho_arquivo, fila_logs, chaves_com_erro))
        processo_robo.start()

    else:
        adicionar_log("❌ Tipo de download inválido.")


    try:
        adicionar_log(f"▶️ Robô {tipo_download} iniciado com sucesso!")
        atualizar_status_robo("EXECUTANDO")
        botao_parar.configure(state="normal", fg_color="red", hover_color="#cc0000")

    except Exception as e:
        adicionar_log(f"❌ Erro ao iniciar o robô: {e}")
        atualizar_status_robo("PRONTO")
        botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")

def parar_robo(lista_chaves_com_erro):
    global processo_robo

    tipoDownload = var_opcao.get()

    if processo_robo and processo_robo.is_alive():
        try:
            processo_robo.terminate()
            processo_robo = None

            # Salvar as chaves com erro da lista compartilhada no arquivo
            caminho_arquivo_erros = fr"C:\Users\{getpass.getuser()}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"
            if lista_chaves_com_erro:
                with open(caminho_arquivo_erros, 'w') as f:
                    for chave in lista_chaves_com_erro:
                        f.write(chave + '\n')
                adicionar_log(f"❌ Execução interrompida. Chaves com erro salvas em: {caminho_arquivo_erros}")
            else:
                adicionar_log("⛔ Execução interrompida. Nenhuma chave com erro foi registrada.")

            adicionar_log("⛔ Execução interrompida pelo usuário.")
            atualizar_status_robo("PARADO")
            botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
            botao_iniciar.configure(state="disabled")

            janelas = gw.getWindowsWithTitle('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')

            for janela in janelas:
                try:
                    janela.close()
                except Exception as e:
                    adicionar_log(f"Erro ao fechar a janela: {e}")

        except Exception as e:
            adicionar_log(f"❌ Erro ao tentar parar o robô: {e}")
    else:
        adicionar_log("⚠️ Nenhum robô está em execução.")
    
    verificarIniciar()
    compararPastaLista(tipoDownload, caminho_arquivo, adicionar_log)

def abrir_ajuda():
    try:
        caminho_pdf = fr"{caminhoPasta}\documentacao\v1.0.1\Documentacao_Interface_Robos_Certezza_v1.0.1.pdf"

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
        comando = r'start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Selenium\ChromeTestProfile" --disable-blink-features=AutomationControlled --app=https://meudanfe.com.br/'
        subprocess.Popen(comando, shell=True)

        time.sleep(2)

        maximizarJanela("Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF")
        verificarIniciar()

    # botao_chrome.configure(state='disabled')

def resetar_opcoes():
    global caminho_arquivo

    var_opcao.set("")
    caminho_arquivo = None  
    arquivo_label.configure(text="Arquivo: Nenhum selecionado") 
    adicionar_log("🔄 Configurações Resetadas\n")
    atualizar_status_robo("PARADO")

    verificarIniciar()

def abrir_pasta():
    os.startfile(pasta_arquivos)

def add_tooltip(widget, text, offset_y=15):
    tooltip = ctk.CTkLabel(root, text=text, fg_color="gray20", text_color="white", corner_radius=6)
    tooltip.place_forget()

    def show_tooltip(event):
        # Posição fixa: acima do botão, centralizado
        widget_x = widget.winfo_rootx() - root.winfo_rootx()
        widget_y = widget.winfo_rooty() - root.winfo_rooty()
        widget_w = widget.winfo_width()

        # Coloca tooltip centralizado acima, com pequeno offset
        tooltip_x = widget_x + widget_w/2 - tooltip.winfo_reqwidth()/2
        tooltip_y = widget_y - tooltip.winfo_reqheight() - offset_y

        tooltip.place(x=tooltip_x, y=tooltip_y)

    def hide_tooltip(event):
        tooltip.place_forget()

    widget.bind("<Enter>", show_tooltip)
    widget.bind("<Leave>", hide_tooltip)

def add_aba_tooltip(notebook, tab_index, tooltip_text):
    tooltip = None

    def on_motion(event):
        nonlocal tooltip
        element = notebook.identify(event.x, event.y)
        if element == "label":
            try:
                hovered_index = notebook.index(f"@{event.x},{event.y}")
                if hovered_index == tab_index:
                    if tooltip is None:
                        x = notebook.winfo_rootx() + event.x + 10
                        y = notebook.winfo_rooty() + event.y + 10
                        tooltip = tk.Toplevel()
                        tooltip.wm_overrideredirect(True)
                        tooltip.wm_attributes("-topmost", True)
                        tooltip.geometry(f"+{x}+{y}")
                        label = tk.Label(
                            tooltip,
                            text=tooltip_text,
                            background="#ffffe0",
                            relief="solid",
                            borderwidth=1,
                            font=("Segoe UI", 10)
                        )
                        label.pack(ipadx=4, ipady=2)
                    return
            except tk.TclError:
                pass

        if tooltip:
            tooltip.destroy()
            tooltip = None

    def on_leave(event):
        nonlocal tooltip
        if tooltip:
            tooltip.destroy()
            tooltip = None

    notebook.bind("<Motion>", on_motion)
    notebook.bind("<Leave>", on_leave)


""" 
==============================================================================
                            FUNÇÕES JANELAS
==============================================================================
"""


def consultar_janelas_abertas():
    janelas = gw.getAllTitles()
    # print(janelas)
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
                        CONFIGURAÇÃO DA INTERFACE
==============================================================================
"""


def interfaceXml(master):

    global var_opcao
    global opcaoPdf
    global opcaoXml
    global opcaoPdfXml
    global log_textbox
    global botao_pasta
    global botao_parar
    global botao_iniciar
    global botao_upload
    global botao_vpn
    global status_label
    global arquivo_label
    # global botao_chrome
    # global botao_resetar

    var_opcao = tk.StringVar(value="")
    caminho_script()


    main_frame = ctk.CTkFrame(master=master, fg_color="#D9D9D9", border_color="#F5C400", border_width=2)
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

    opcaoPdf = ctk.CTkRadioButton(opcoes_frame, text="PDF", variable=var_opcao, value="PDF", radiobutton_height=20, radiobutton_width=20, command=selecionar_opcao, hover_color="#ffcc00", fg_color="#ffcc00", state="normal")
    opcaoPdf.pack(side="left")

    opcaoXml = ctk.CTkRadioButton(opcoes_frame, text="XML", variable=var_opcao, value="XML", radiobutton_height=20, radiobutton_width=20, command=selecionar_opcao, hover_color="#ffcc00", fg_color="#ffcc00")
    opcaoXml.pack(side="left")

    opcaoPdfXml = ctk.CTkRadioButton(opcoes_frame, text="PDF E XML", variable=var_opcao, value="PDF_XML", radiobutton_height=20, radiobutton_width=20, command=selecionar_opcao, hover_color="#ffcc00", fg_color="#ffcc00", state="normal")
    opcaoPdfXml.pack(side="left")

    # add_tooltip(opcaoPdfXml, "Em desenvolvimento")

    instrucoes = (
        "1- ESCOLHER O TIPO DE DOWNLOAD (PDF/XML/AMBOS)\n"
        "2- FAZER UPLOAD DE ARQUIVO TXT CONTENDO AS NOTAS UMA EMBAIXO DA OUTRA\n"
        "3- ATIVAR A VPN\n"
        # "4- ABRIR CHROME\n"
        "4- CLICAR NO BOTÃO DE INICIAR"
    )
    ctk.CTkLabel(main_frame, text=instrucoes, justify="left", font=("Arial", 11)).grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="w")

    botoes_linha_frame = ctk.CTkFrame(master=main_frame, fg_color="transparent")
    botoes_linha_frame.grid(row=3, column=0, columnspan=2, padx=(10), pady=(5, 10), sticky="w")

    botao_upload = ctk.CTkButton(botoes_linha_frame, text="1- UPLOAD", command=upload_arquivo, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    botao_upload.grid(row=0, column=0, padx=(10, 10))

    botao_vpn = ctk.CTkButton(botoes_linha_frame, text="2- ABRIR VPN", command=abrir_vpn_thread, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    botao_vpn.grid(row=0, column=1, padx=(0, 10))

    botao_pasta = ctk.CTkButton(botoes_linha_frame, text="ABRIR PASTA", command=abrir_pasta, fg_color="#A9A9A9", hover_color="#808080", text_color="black", width=100)
    botao_pasta.grid(row=0, column=4, padx=(0, 10))

    if os.path.exists(pasta_arquivos):
        botao_pasta.configure(state="normal")
    else:
        botao_pasta.configure(state="disabled")

    botao_iniciar = ctk.CTkButton(botoes_linha_frame, text="3- INICIAR", command=iniciar_robo, fg_color="#00CC00", hover_color="#009900", text_color="black", width=100, state="disabled")
    botao_iniciar.grid(row=0, column=2, padx=(0, 10))

    botao_parar = ctk.CTkButton(botoes_linha_frame, text="PARAR EXECUÇÃO", command=lambda: parar_robo(chaves_com_erro), fg_color="#A9A9A9", hover_color="#808080", text_color="white", width=120, state="disabled")
    botao_parar.grid(row=0, column=3, padx=(0, 10))
    
    log_frame = ctk.CTkFrame(master=master, fg_color="#FFFFFF", border_color="#F5C400", border_width=2)
    log_frame.pack(padx=10, pady=(0, 10), fill="both", expand=True)

    log_inner_frame = tk.Frame(log_frame, bg="#FFFFFF")
    log_inner_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Scrollbar
    scrollbar = tk.Scrollbar(log_inner_frame)
    scrollbar.pack(side="right", fill="y")

    # Textbox com Scrollbar
    log_textbox = tk.Text(log_inner_frame, bg="#EDEDED", font=("Courier New", 10), relief="flat", bd=0, highlightthickness=0, wrap="word", yscrollcommand=scrollbar.set)
    log_textbox.pack(side="left", fill="both", expand=True)

    scrollbar.config(command=log_textbox.yview)


    adicionar_log("⚠️ CASO HOUVER ALGUMA DÚVIDA EM RELAÇÃO AS FUNCIONALIDADES DO SISTEMA, CLIQUE NO BOTÃO 'AJUDA'\n")

    monitorar_logs()
    monitorar_pasta()
    # monitorarChrome()

    return caminho_arquivo

def interfaceEcac(master):
    
    global var_opcao_ecac
    global log_textbox_ecac
    global botao_parar_ecac
    global botao_iniciar_ecac
    global status_label_ecac
    global botao_chrome_ecac
    global inicio_data_ecac
    global fim_data_ecac
    # global botao_upload
    # global botao_vpn
    # global arquivo_label

    var_opcao_ecac = tk.StringVar(value="")
    caminho_script()


    main_frame_ecac = ctk.CTkFrame(master=master, fg_color="#D9D9D9", border_color="#F5C400", border_width=2)
    main_frame_ecac.pack(padx=10, pady=(10, 0), fill="x")

    # Cabeçalho com status e arquivo
    header_frame_ecac = ctk.CTkFrame(master=main_frame_ecac, fg_color="transparent")
    header_frame_ecac.grid(row=0, column=0, columnspan=2, padx=10, pady=(5, 0), sticky="ew")
    header_frame_ecac.grid_columnconfigure(0, weight=0)  # Status
    header_frame_ecac.grid_columnconfigure(1, weight=1)  # Arquivo (vai expandir)
    header_frame_ecac.grid_columnconfigure(2, weight=0)  # Botão Ajuda

    status_label_ecac = ctk.CTkLabel(header_frame_ecac, text="STATUS: PARADO", font=("Arial", 12, "bold"), text_color="red")
    status_label_ecac.grid(row=0, column=0, padx=(0, 10), sticky="w")

    # arquivo_label = ctk.CTkLabel(header_frame_ecac, text="Arquivo: Nenhum arquivo selecionado", font=("Arial", 11), text_color="black")
    # arquivo_label.grid(row=0, column=1, sticky="w")

    botao_ajuda_ecac = ctk.CTkButton(header_frame_ecac, text="AJUDA", command=abrir_ajuda,
                                fg_color="#007B8A", hover_color="#006472", text_color="white", width=100)
    botao_ajuda_ecac.grid(row=0, column=3, sticky="e", padx=(360, 0), pady=(5, 0))


    # Radiobuttons=========================================================================

    opcoes_frame_ecac = ctk.CTkFrame(master=main_frame_ecac, fg_color="transparent")
    opcoes_frame_ecac.grid(row=1, column=0, columnspan=2, pady=(10, 0), padx=10, sticky="w")

    dctf_pdf = ctk.CTkRadioButton(
        opcoes_frame_ecac, text="DCTF PDF", variable=var_opcao_ecac, value="DCTF_PDF",
        radiobutton_height=20, radiobutton_width=20,
        command=selecionar_opcao_ecac, hover_color="#ffcc00", fg_color="#ffcc00"
    )
    dctf_pdf.grid(row=0, column=0, sticky="w", padx=(0, 10))

    inicio_data_ecac = ctk.CTkEntry(opcoes_frame_ecac, placeholder_text="Ano Início YYYY", width=120, state="disabled")
    inicio_data_ecac.grid(row=0, column=1, padx=5)

    fim_data_ecac = ctk.CTkEntry(opcoes_frame_ecac, placeholder_text="Ano Fim YYYY", width=120, state="disabled")
    fim_data_ecac.grid(row=0, column=2, padx=5)

    main_frame_ecac.after(50, lambda: botao_ajuda_ecac.focus())

    # Linha 2: DCTF WEB e DARF
    dctf_web = ctk.CTkRadioButton(
        opcoes_frame_ecac, text="DCTF WEB", variable=var_opcao_ecac, value="DCTF_WEB",
        radiobutton_height=20, radiobutton_width=20,
        command=selecionar_opcao_ecac, hover_color="#ffcc00", fg_color="#ffcc00"
    )
    dctf_web.grid(row=1, column=0, sticky="w", pady=(5, 0))

    darf = ctk.CTkRadioButton(
        opcoes_frame_ecac, text="DARF", variable=var_opcao_ecac, value="DARF",
        radiobutton_height=20, radiobutton_width=20,
        command=selecionar_opcao_ecac, hover_color="#ffcc00", fg_color="#ffcc00"
    )
    darf.grid(row=1, column=1, sticky="w", pady=(5, 0), padx=(5, 0))

    pgfn = ctk.CTkRadioButton(
        opcoes_frame_ecac, text="PGFN", variable=var_opcao_ecac, value="PGFN",
        radiobutton_height=20, radiobutton_width=20,
        command=selecionar_opcao_ecac, hover_color="#ffcc00", fg_color="#ffcc00"
    )
    pgfn.grid(row=1, column=2, sticky="w", pady=(5, 0), padx=(5, 0))


    fnts_pagadoras = ctk.CTkRadioButton(
        opcoes_frame_ecac, text="FONTES PAGADORAS", variable=var_opcao_ecac, value="FNTS_PAGADORAS",
        radiobutton_height=20, radiobutton_width=20,
        command=selecionar_opcao_ecac, hover_color="#ffcc00", fg_color="#ffcc00"
    )
    fnts_pagadoras.grid(row=1, column=3, sticky="w", pady=(5, 0), padx=(0))


    #===================================================================================

    instrucoes = (
        "1- ESCOLHER O TIPO DE DOWNLOAD (DCTF PDF/WEB | DARF | PGFN)\n"
        "2- ABRIR CHROME\n"
        "3- LOGAR COM O CERTIFICADO NO ECAC\n"
        "4- CLICAR NO BOTÃO DE INICIAR"
    )
    ctk.CTkLabel(main_frame_ecac, text=instrucoes, justify="left", font=("Arial", 11)).grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="w")

    botoes_linha_frame_ecac = ctk.CTkFrame(master=main_frame_ecac, fg_color="transparent")
    botoes_linha_frame_ecac.grid(row=3, column=0, columnspan=2, padx=(10), pady=(5, 10), sticky="w")

    # botao_upload = ctk.CTkButton(botoes_linha_frame_ecac, text="1- UPLOAD", command=upload_arquivo, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    # botao_upload.grid(row=0, column=0, padx=(0, 10))

    # botao_vpn = ctk.CTkButton(botoes_linha_frame_ecac, text="2- ABRIR VPN", command=abrir_vpn_thread, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    # botao_vpn.grid(row=0, column=1, padx=(0, 10))

    botao_chrome_ecac = ctk.CTkButton(botoes_linha_frame_ecac, text="1- ABRIR CHROME", command=abrir_chrome, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
    botao_chrome_ecac.grid(row=0, column=2, padx=(0, 10))

    botao_iniciar_ecac = ctk.CTkButton(botoes_linha_frame_ecac, text="2- INICIAR", command=iniciar_robo, fg_color="#00CC00", hover_color="#009900", text_color="black", width=100, state="disabled")
    botao_iniciar_ecac.grid(row=0, column=3, padx=(0, 10))

    botao_parar_ecac = ctk.CTkButton(botoes_linha_frame_ecac, text="PARAR EXECUÇÃO", command=parar_robo, fg_color="#A9A9A9", hover_color="#808080", text_color="white", width=120, state="disabled")
    botao_parar_ecac.grid(row=0, column=4, padx=(0, 10))

    botao_pasta = ctk.CTkButton(botoes_linha_frame_ecac, text="ABRIR PASTA", command=abrir_pasta, fg_color="#A9A9A9", hover_color="#808080", text_color="black", width=100)
    botao_pasta.grid(row=0, column=5, padx=(0, 10))

    if os.path.exists(pasta_arquivos):
        botao_pasta.configure(state="normal")
    else:
        botao_pasta.configure(state="disabled")

    log_frame = ctk.CTkFrame(master=master, fg_color="#FFFFFF", border_color="#F5C400", border_width=2)
    log_frame.pack(padx=10, pady=(0, 10), fill="both", expand=True)

    log_inner_frame = tk.Frame(log_frame, bg="#FFFFFF")
    log_inner_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Scrollbar
    scrollbar = tk.Scrollbar(log_inner_frame)
    scrollbar.pack(side="right", fill="y")

    # Textbox com Scrollbar
    log_textbox_ecac = tk.Text(log_inner_frame, bg="#EDEDED", font=("Courier New", 10), relief="flat", bd=0, highlightthickness=0, wrap="word", yscrollcommand=scrollbar.set)
    log_textbox_ecac.pack(side="left", fill="both", expand=True)

    scrollbar.config(command=log_textbox.yview)

    adicionar_log_ecac("⚠️ CASO HOUVER ALGUMA DÚVIDA EM RELAÇÃO AS FUNCIONALIDADES DO SISTEMA, CLIQUE NO BOTÃO 'AJUDA'\n")

    monitorar_logs()
    monitorarChrome()

    return caminho_arquivo





""" 
|===================================================================================================================|
|                                                    INTERFACE ECAC                                                 |
|===================================================================================================================|
"""


""" 
==============================================================================
                ADICIONAR LOGS NA AREA DE TEXTO DA INTERFACE
==============================================================================
"""

log_textbox_ecac = None
fila_logs_ecac = Queue() 


def monitorar_logs_ecac():
    while not fila_logs_ecac.empty():
        mensagem = fila_logs_ecac.get()

        if mensagem.startswith("SINAL:"):
            comando = mensagem.split("SINAL:")[1]

            if comando == "CONCLUIDO":
                botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
                botao_iniciar.configure(state="disabled")
                # botao_chrome.configure(state='normal')
                botao_upload.configure(state='normal')
                botao_vpn.configure(state='normal')
                opcaoXml.configure(state="normal")
                opcaoPdf.configure(state="normal")
                opcaoPdfXml.configure(state="normal")

                atualizar_status_robo_ecac("CONCLUIDO")

        else:
            adicionar_log_ecac(mensagem)

    root.after(500, monitorar_logs_ecac)

def adicionar_log_ecac(texto):
    global log_textbox_ecac
    if log_textbox_ecac is None:
        print(texto)
        return

    
    log_textbox_ecac.configure(state="normal")
    log_textbox_ecac.insert("end", texto + "\n")
    log_textbox_ecac.configure(state="disabled")
    log_textbox_ecac.see("end")
    


""" 
==============================================================================
                                FUNÇÕES ROBOS
==============================================================================
"""

#---ECAC---

def robo_dctf_pdf():
    mkey = mousekey.MouseKey()

    def iniciar_navegador(com_debugging_remoto=True):
        chrome_driver_path = ChromeDriverManager().install()
        chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
        
        #print(f"ChromeDriver path: {chrome_driver_executable}")
        if not os.path.isfile(chrome_driver_executable):
            raise FileNotFoundError(f"O ChromeDriver não foi encontrado em {chrome_driver_executable}")

        service = Service(executable_path=chrome_driver_executable)
        
        chrome_options = Options()
        if com_debugging_remoto:
            remote_debugging_port = 9222
            chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
        
        navegador = webdriver.Chrome(service=service, options=chrome_options)
        return navegador

    navegador = iniciar_navegador(com_debugging_remoto=True)

    def procurar_imagem(nome_arquivo, confidence=0.8, region=None, maxTentativas=60, horizontal=0, vertical=0, dx=0, dy=0, acao='clicar', clicks=1, ocorrencia=1, delay_tentativa=1):
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

    def ajustar_data(data):
        meses = {
            "Janeiro": "01",
            "Fevereiro": "02",
            "Marco": "03",
            "Março": "03",
            "Abril": "04",
            "Maio": "05",
            "Junho": "06",
            "Julho": "07",
            "Agosto": "08",
            "Setembro": "09",
            "Outubro": "10",
            "Novembro": "11",
            "Dezembro": "12"
        }
        
        for mes_nome, mes_numero in meses.items():
            data = data.replace(mes_nome, mes_numero)
        
        return data

    def periodo(anoInicio, anoFim):
        navegador.switch_to.default_content()
        iframe = navegador.find_element(By.ID, 'frmApp')
        navegador.switch_to.frame(iframe)

        linha = 2
        countIn = 0
        while True:
            try:
                periodo_element = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{linha}]/td[2]')
                periodoTxtIn = periodo_element.text
                #print(f"anoIn: {periodoTxtIn}")
                barraIn = periodoTxtIn.find("/")
                anoIn = periodoTxtIn[barraIn+1:barraIn+5]

                #print(anoIn)
                if anoInicio == anoIn and countIn == 0:
                    linhaAnoInicio = linha
                    #print(linhaAnoInicio)
                    countIn += 1


                periodoFim_element = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{linha+1}]/td[2]')
                periodoTxtFim = periodoFim_element.text
                #print(f"anoFi: {periodoTxtFim}")
                barraFim = periodoTxtFim.find("/")
                anoFi = periodoTxtFim[barraFim+1:barraFim+5]
                
                if anoIn != anoFi:
                    if anoIn == anoFim:
                        linhaAnoFim = linha
                        #print(linhaAnoFim)
                        break

                linha += 1
            except NoSuchElementException:
                if anoFim == "2024":
                    linhaAnoFim = linha
                break

        return int(linhaAnoInicio), int(linhaAnoFim)

    def percorrer_dctf(anoInicio, anoFim, Ativa="S"):
        #procurar_imagem(r'C:\VS_CODE_MAIN\0 - ECAC\DCTF - PDF\prints\chrome.png', maxTentativas=5, ocorrencia=2)

        linhaAnoInicio, linhaAnoFim = periodo(anoInicio, anoFim)
        print(f"{linhaAnoInicio} - {linhaAnoFim}" )


        i = linhaAnoInicio 
        while i <= linhaAnoFim:
            try:
                navegador.switch_to.default_content()

                iframe = navegador.find_element(By.ID, 'frmApp')
                navegador.switch_to.frame(iframe)

                
                data = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[2]')
                dataRecepcao = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[3]')
                status = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[7]')
                btimprimir = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[8]/input[2]')
                TXTdata = data.text
                TXTdata = TXTdata.replace("/","_")

                TXTdataRecep = dataRecepcao.text
                TXTdataRecep = TXTdataRecep.replace("/","_")

                TXTstatus = status.text
                TXTstatus = TXTstatus.replace("/ ","_")

                dataRecepcao


                ano = TXTdata[-4:]
                

                mes = ajustar_data(TXTdata)

                #print(mes, " | ", TXTstatus, " | ", ano, " | ", TXTdataRecep)

                time.sleep(0.5)

                if "Ativa" in TXTstatus:
                    nome_arquivo = f"{mes} - 0 - {TXTdataRecep} - {TXTstatus} - {i}"
                elif "Cancelada" in TXTstatus:
                    nome_arquivo = f"{mes} - 1 - {TXTdataRecep} - {TXTstatus} - {i}"


                print (nome_arquivo)
                        
                match Ativa:
                    case ("S" | "s"):
                        if 'Ativa' in TXTstatus:
                            btimprimir.send_keys(Keys.CONTROL + Keys.RETURN)
                            popup = navegador.switch_to.alert
                            time.sleep(1.5)
                            popup.accept()
                            robo(ano, nome_arquivo, i)

                    case ("N" | "n"):
                            btimprimir.send_keys(Keys.CONTROL + Keys.RETURN)
                            popup = navegador.switch_to.alert
                            time.sleep(1.5)
                            popup.accept()
                            robo(ano, nome_arquivo, i)

                i += 1
            except NoSuchElementException:
                break

    def robo(ano, nome_arquivo, i):
        abas = navegador.window_handles
        navegador.switch_to.window(abas[1])
        
        if procurar_imagem(r'0 - ECAC\DCTF - PDF\prints\ministeriofaze.png'):
            pass
        else:
            navegador.refresh()
            procurar_imagem(r'0 - ECAC\DCTF - PDF\prints\ministeriofaze.png')

        novo_caminho = fr"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\DCTF - PDF\PDFs\{ano}"

        if not os.path.exists(novo_caminho):
            os.makedirs(novo_caminho)
        time.sleep(0.5)
        
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(1.5)
        procurar_imagem(r'0 - ECAC\DCTF - PDF\prints\btnSalvar.png')

        caminho_completo = f"{novo_caminho}\\{nome_arquivo}.pdf"
        pyperclip.copy(caminho_completo)
        time.sleep(1.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(1.5)
        pyautogui.hotkey('ctrl', 'w')
        navegador.switch_to.window(abas[0])
        navegador.switch_to.default_content()

    def calcular_tempo_execucao(func, *args, **kwargs):
        tempoInicio = datetime.now()
        resultado = func(*args, **kwargs)
        tempoFim = datetime.now()
        tempoExecucao = str(tempoFim - tempoInicio)

        larguraTotal = max(len(tempoExecucao) + 8, 30)  # Garante espaço suficiente
        titulo = " TEMPO DE EXECUÇÃO "
        
        linhaSuperior = f"╭{'─' * (larguraTotal - 2)}╮"
        linhaTitulo = f"│{titulo.center(larguraTotal - 2)}│"
        linhaMeio = f"│{tempoExecucao.center(larguraTotal - 2)}│"
        linhaInferior = f"╰{'─' * (larguraTotal - 2)}╯"
        
        print(linhaSuperior)
        print(linhaTitulo)
        print(linhaMeio)
        print(linhaInferior)
        
        return resultado

    print()
    calcular_tempo_execucao(percorrer_dctf, anoInicio="2022", anoFim="2024")

def robo_dctf_web():

    def iniciar_navegador(com_debugging_remoto=True):
        chrome_driver_path = ChromeDriverManager().install()
        chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
        
        #print(f"ChromeDriver path: {chrome_driver_executable}")
        if not os.path.isfile(chrome_driver_executable):
            raise FileNotFoundError(f"O ChromeDriver não foi encontrado em {chrome_driver_executable}")

        service = Service(executable_path=chrome_driver_executable)
        
        chrome_options = Options()
        if com_debugging_remoto:
            remote_debugging_port = 9222
            chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
        
        navegador = webdriver.Chrome(service=service, options=chrome_options)
        return navegador

    navegador = iniciar_navegador(com_debugging_remoto=True)

    #----------------------------------------------------------------------------------------------------------------
    def aplicarFiltros(anoAp, anoTr, filtroAlternativo=False):
        DtIniApuracao = f"01/01/{anoAp}"
        DtFimApuracao = f"31/12/{anoAp}"

        DtIniTransmissao = f"01/01/{anoTr}"

        anoAtual = datetime.now().year
        hoje = datetime.now().strftime("%d/%m/%Y")

        if anoTr == anoAtual:
            DtFimTransmissao = hoje
        else:
            DtFimTransmissao = f"26/12/{anoTr}"

        if filtroAlternativo:
            DtIniTransmissao = f"27/12/{anoTr}"
            DtFimTransmissao = f"31/12/{anoTr}"

        try:
            WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

            navegador.switch_to.default_content()
            iframe = navegador.find_element(By.ID, 'frmApp')
            navegador.switch_to.frame(iframe)

            #Apuracao
            campoDataInicioAp = navegador.find_element(By.XPATH, '//*[@id="txtDataInicio"]')
            campoDataInicioAp.clear()
            campoDataInicioAp.send_keys(DtIniApuracao)

            campoDataFimAp = navegador.find_element(By.XPATH, '//*[@id="txtDataFinal"]')
            campoDataFimAp.clear()
            campoDataFimAp.send_keys(DtFimApuracao)

            #Transmisssao
            campoDataInicioTr = navegador.find_element(By.XPATH, '//*[@id="txtDataTransmissaoInicial"]')
            campoDataInicioTr.clear()
            campoDataInicioTr.send_keys(DtIniTransmissao)

            campoDataFimTr = navegador.find_element(By.XPATH, '//*[@id="txtDataTransmissaoFinal"]')
            campoDataFimTr.clear()
            campoDataFimTr.send_keys(DtFimTransmissao)


            #categoria declaração
            categoriasFiltradas = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div/button/span[1]')
            txtCategoriasFiltrdas = categoriasFiltradas.text

            if "6 selecionados" in txtCategoriasFiltrdas:
                filtroTipoDeclaracao = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div')
                navegador.execute_script("arguments[0].className = 'btn-group bootstrap-select show-tick span11 open';", filtroTipoDeclaracao)
                selecionarNenhum = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div/div/div/div/button[2]')
                time.sleep(0.5)
                navegador.execute_script("arguments[0].click();", selecionarNenhum)

                opcoes = navegador.find_elements(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div/div/ul/li/a')

                for opcao in opcoes:
                    textoOpcao = opcao.text.strip().lower()
                    if "geral" in textoOpcao or "13" in textoOpcao:
                        navegador.execute_script("arguments[0].click();", opcao)
                        time.sleep(0.2)  

            #situação declaração
            situacoesFiltradas = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[4]/div/div/div/button/span[1]')
            txtSituacoesFiltradas = situacoesFiltradas.text

            if 'Em andamento' in txtSituacoesFiltradas:
                filtroSituacaoDeclaracao = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[4]/div/div/div')
                navegador.execute_script("arguments[0].className = 'btn-group bootstrap-select show-tick span10 open';", filtroSituacaoDeclaracao)
                filtroEmAndamento = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[4]/div/div/div/div/ul/li[1]/a')
                time.sleep(0.5)
                navegador.execute_script("arguments[0].click();", filtroEmAndamento)

            botaoPesquisar = navegador.find_element(By.XPATH, '//*[@id="ctl00_cphConteudo_btnFiltar"]')
            time.sleep(0.5)
            navegador.execute_script("arguments[0].click();", botaoPesquisar)

        except (TimeoutException, NoSuchElementException):
            print("erro ao aplicaro o filtro:")
            traceback.print_exc()

    def baixarXML():
        try:
            WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

            navegador.switch_to.default_content()

            iframe = navegador.find_element(By.ID, 'frmApp')
            navegador.switch_to.frame(iframe)

            time.sleep(0.5)
            dropDownRelatorios = WebDriverWait(navegador, 15).until(EC.visibility_of_element_located(
                        (By.XPATH, '//*[@id="dropDown_Relatorios"]')))
            time.sleep(0.5)
            #dropDownRelatorios = navegador.find_element(By.XPATH, '//*[@id="dropDown_Relatorios"]')
            #driver.execute_script("arguments[0].className = 'btn-group bootstrap-select show-tick span11 open';", filtroTipoDeclaracao)

            navegador.execute_script("arguments[0].className = 'class=dropdown open' ", dropDownRelatorios)

            time.sleep(1)
            baixarXMLdeSaida = WebDriverWait(navegador, 15).until(EC.visibility_of_element_located(
                        (By.XPATH, '//*[@id="dropDown_Relatorios_Menu"]/li[3]/a')))
            time.sleep(0.5)
            #baixarXMLdeSaida = navegador.find_element(By.XPATH, '//*[@id="dropDown_Relatorios_Menu"]/li[3]/a')
            baixarXMLdeSaida.click()
            
        except (TimeoutException, Exception):
            print("erro ao baixar xml!")
            traceback.print_exc()

    def percorrer_dctf():
        anoAp, anoTr = 2018, 2018
        d = 2
        filtroAlternativo = False
        anoAtual = datetime.now().year

        registrosBaixados = set()  

        aplicarFiltros(anoAp, anoTr, filtroAlternativo)

        while True:
            try:
                WebDriverWait(navegador, 90).until(
                    lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

                navegador.switch_to.default_content()
                iframe = navegador.find_element(By.ID, 'frmApp')
                navegador.switch_to.frame(iframe)

                Dctf = True
                Visualizar = True
                try:
                    campoPeriodoApu = navegador.find_element(By.XPATH, f'//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"]/tbody/tr[{d}]/td[1]')
                    periodoApuracao = campoPeriodoApu.text.replace("/", "_")

                    campoTipo = navegador.find_element(By.XPATH, f'//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"]/tbody/tr[{d}]/td[5]')
                    tipo = campoTipo.text

                    campoSituacao = navegador.find_element( By.XPATH, f'//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"]/tbody/tr[{d}]/td[6]')
                    situacao = campoSituacao.text

                    try:
                        botaoVisualizar = navegador.find_element( By.XPATH, f'/html/body/form/div[5]/div/div/div[6]/div/div/div[1]/table/tbody/tr[{d}]/td[9]/a[1]')
                    except NoSuchElementException:
                        Visualizar = False

                except NoSuchElementException:
                    Dctf = False

                # === INÍCIO DA LÓGICA PRINCIPAL ===

                if anoTr > anoAtual:
                    filtroAlternativo = False
                    anoAp += 1
                    anoTr = anoAp
                    d = 2  
                    if anoAp > anoAtual:
                        break  
                    aplicarFiltros(anoAp, anoTr, filtroAlternativo)
                    continue 

                if not Dctf:
                    print()
                    d = 2 
                    if not filtroAlternativo:
                        filtroAlternativo = True
                        aplicarFiltros(anoAp, anoTr, filtroAlternativo)
                        continue
                    else:
                        filtroAlternativo = False
                        anoTr += 1
                        if anoTr == 2026:
                            anoAp += 1
                            anoTr = anoAp
                        aplicarFiltros(anoAp, anoTr, filtroAlternativo)
                        continue

                elif Visualizar:
                    chave = f"{periodoApuracao}-{tipo}-{situacao}"
                    if chave in registrosBaixados:
                        print(f"Já baixado: {chave}, pulando...")
                        d += 1
                        continue
                    registrosBaixados.add(chave)

                    print(f'{periodoApuracao} - {tipo} - {situacao}')
                    time.sleep(0.5)
                    botaoVisualizar.click()
                    d += 1
                    baixarXML()
                    navegador.back()
                    continue

                else:
                    print(f"Sem opção visualizar na DCTF da linha {d-1}º, pulando pra próxima")
                    d += 1
                    continue

            except (TimeoutException, Exception):
                print("Erro ao percorrer a lista!")
                traceback.print_exc()
                break

    def redirecionarXml(origem=r'C:\VS_CODE\DOWNLOAD_ARQUIVOS',destino=r'C:\VS_CODE\0 - ECAC\DCTF - WEB\XMLs'):

        for arquivo in os.listdir(origem):
            if "XMLSaida" in arquivo and arquivo.endswith(".xml"):
                caminhoOrigem = os.path.join(origem, arquivo)
                caminhoDestino = os.path.join(destino, arquivo)
                
                try:
                    shutil.move(caminhoOrigem, caminhoDestino)
                    print(f"Arquivo {arquivo} foi movido para {destino}")
                except Exception as e:
                    print(f"Erro ao mover o arquivo {arquivo}: {e}") 


    print("iniciando...")

    time.sleep(1)
    percorrer_dctf()
    #redirecionarXml()


""" 
==============================================================================
                            VERIFICAÇÕES E STATUS

Verifica se está tudo OK para iniciar o Robô e atualiza o status na interface
==============================================================================
"""

def atualizar_status_robo_ecac(status):
    status_label.configure(text=f"STATUS: {status}")
    if status == "EXECUTANDO":
        status_label.configure(text_color="red")
        botao_iniciar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
        botao_upload.configure(state="disabled")
        botao_vpn.configure(state="disabled")
        # botao_chrome.configure(state="disabled")
        opcaoXml.configure(state="disabled")
        opcaoPdf.configure(state="disabled")
        opcaoPdfXml.configure(state="disabled")
        # botao_resetar.configure(state="disabled")

    elif status in ["PARADO", "PRONTO"]:
        status_label.configure(text_color="red")
        botao_iniciar.configure(state="normal", fg_color="#00CC00", hover_color="#009900")
        botao_upload.configure(state="normal")
        botao_vpn.configure(state="normal")
        opcaoXml.configure(state="normal")
        opcaoPdf.configure(state="normal")
        opcaoPdfXml.configure(state="normal")
        # botao_resetar.configure(state="normal")
        # botao_chrome.configure(state="normal")

    elif status == "CONCLUIDO":
        status_label.configure(text_color="green")
        botao_iniciar.configure(state="normal", fg_color="#00CC00", hover_color="#009900")
        botao_upload.configure(state="normal")
        botao_vpn.configure(state="normal")
        # botao_resetar.configure(state="normal")
        opcaoXml.configure(state="normal")
        opcaoPdf.configure(state="normal")
        opcaoPdfXml.configure(state="normal")
        # botao_chrome.configure(state="normal")

def verificarIniciar_ecac():
    tipoDownloadEcac = var_opcao_ecac.get()  

    if tipoDownloadEcac == "DCTF_PDF":
        inicio_data_ecac.configure(state="normal", placeholder_text="Ano Início YYYY")
        fim_data_ecac.configure(state="normal", placeholder_text="Ano Fim YYYY")
    elif tipoDownloadEcac in ["DCTF_WEB", "DARF", "PGFN", "FNTS_PAGADORAS"]:
        inicio_data_ecac.configure(state="disabled", placeholder_text="Ano Início YYYY")
        fim_data_ecac.configure(state="disabled", placeholder_text="Ano Fim YYYY")


    if tipoDownloadEcac != "" and caminho_arquivo:
        adicionar_log_ecac("✅ PRONTO PARA INICIAR\n")
        atualizar_status_robo_ecac("PRONTO")


def selecionar_opcao_ecac():
    verificarIniciar_ecac()


""" 
==============================================================================
                            BOTÕES INTERFACE
==============================================================================
"""

def iniciar_robo_ecac():

    global processo_robo_ecac

    if processo_robo_ecac and processo_robo_ecac.is_alive():
        adicionar_log_ecac("⚠️ Já existe um robô em execução.")
        return
    
    tipoDownloadEcac = var_opcao_ecac.get()

    if tipoDownloadEcac == "":
        adicionar_log_ecac("⚠️ Selecione uma opção de download antes de iniciar.")
        return


    # adicionar_log("Robô iniciado com sucesso!")

    if tipoDownloadEcac == "DCTF_PDF":
        processo_robo_ecac = Process(target=robo_dctf_pdf, args=(caminho_arquivo, fila_logs))
        processo_robo_ecac.start()

    elif tipoDownloadEcac == "DCTF_WEB":
        processo_robo_ecac = Process(target=robo_pdf, args=(caminho_arquivo, fila_logs))
        processo_robo_ecac.start()

    elif tipoDownloadEcac == "FONTES_PAGADORAS":
        processo_robo_ecac = Process(target=robo_pdf_xml, args=(caminho_arquivo, fila_logs))
        processo_robo_ecac.start()
    
    elif tipoDownloadEcac == "PGFN":
        processo_robo_ecac = Process(target=robo_pdf_xml, args=(caminho_arquivo, fila_logs))
        processo_robo_ecac.start()

    else:
        adicionar_log_ecac("❌ Tipo de download inválido.")


    try:
        adicionar_log_ecac(f"▶️ Robô {tipoDownloadEcac} iniciado com sucesso!")
        atualizar_status_robo_ecac("EXECUTANDO")
        botao_parar.configure(state="normal", fg_color="red", hover_color="#cc0000")

    except Exception as e:
        adicionar_log_ecac(f"❌ Erro ao iniciar o robô: {e}")
        atualizar_status_robo_ecac("PRONTO")
        botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")

def parar_robo_ecac(lista_chaves_com_erro):
    global processo_robo

    tipoDownload = var_opcao.get()

    if processo_robo and processo_robo.is_alive():
        try:
            processo_robo.terminate()
            processo_robo = None

            # Salvar as chaves com erro da lista compartilhada no arquivo
            caminho_arquivo_erros = fr"C:\Users\{getpass.getuser()}\Desktop\DOWNLOAD'S ROBOS\erros_download.txt"
            if lista_chaves_com_erro:
                with open(caminho_arquivo_erros, 'w') as f:
                    for chave in lista_chaves_com_erro:
                        f.write(chave + '\n')
                adicionar_log(f"❌ Execução interrompida. Chaves com erro salvas em: {caminho_arquivo_erros}")
            else:
                adicionar_log("⛔ Execução interrompida. Nenhuma chave com erro foi registrada.")

            adicionar_log("⛔ Execução interrompida pelo usuário.")
            atualizar_status_robo("PARADO")
            botao_parar.configure(state="disabled", fg_color="#A9A9A9", hover_color="#808080")
            botao_iniciar.configure(state="disabled")

            janelas = gw.getWindowsWithTitle('Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF')

            for janela in janelas:
                try:
                    janela.close()
                except Exception as e:
                    adicionar_log(f"Erro ao fechar a janela: {e}")

        except Exception as e:
            adicionar_log(f"❌ Erro ao tentar parar o robô: {e}")
    else:
        adicionar_log("⚠️ Nenhum robô está em execução.")
    
    verificarIniciar()
    compararPastaLista(tipoDownload, caminho_arquivo, adicionar_log)

def abrir_chrome():

    opcaoDownload = var_opcao_ecac.get()

    if opcaoDownload == "":
        adicionar_log("⚠️ Selecione uma opção de download antes de iniciar.")
        return

    elif opcaoDownload in ['DCTF_PDF', 'DCTF_WEB', 'FONTES_PAGADORAS']:
        comando = r'start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Selenium\ChromeTestProfile" --disable-blink-features=AutomationControlled --app=https://meudanfe.com.br/'
        subprocess.Popen(comando, shell=True)

        time.sleep(2)

        maximizarJanela("Meu Danfe - Gerar Danfe, Buscar NFe pela chave, Danfe PDF")
        verificarIniciar()

    # botao_chrome.configure(state='disabled')





""" 
==============================================================================
                        CONFIGURAÇÕES TELA PRINCIPAL
==============================================================================
"""


def telaPrincipal():
    global root

    caminho_script()

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Painel de Robôs - Certezza")
    root.geometry("650x600")
    root.configure(fg_color="#EDEDED") 
    root.iconbitmap(resource_path(fr"{caminhoPasta}\img\logoICO.ico"))
    root.resizable(False, False)

    style = Style(theme="flatly")
    style.configure('TNotebook', borderwidth=0, background="white")
    style.configure('TNotebook.Tab', padding=[10, 5])

    container = tk.Frame(master=root)
    container.pack(fill='both', expand=True, padx=10, pady=10)

    notebook = Notebook(container, bootstyle="light") 
    notebook.pack(expand=True, fill="both", padx=0, pady=0)

    notebook.tk.call("ttk::style", "configure", "TNotebook", "-borderwidth", "0")
    notebook.tk.call("ttk::style", "configure", "TNotebook.Tab", "-focuscolor", "white")

    # Criar abas
    aba_xml = tk.Frame(notebook, bg="white")
    aba_ecac = tk.Frame(notebook, bg="white")
    aba_esocial = tk.Frame(notebook, bg="white")

    notebook.add(aba_xml, text="XML")
    notebook.add(aba_ecac, text="eCAC")
    notebook.add(aba_esocial, text="eSocial")

    notebook.tab(1, state="disabled")
    notebook.tab(2, state="disabled")

    add_aba_tooltip(notebook, 2, "Em desenvolvimento: eCAC | eSocial")


    # Chamar as interfaces dentro das abas
    interfaceXml(aba_xml)
    # interfaceEcac(aba_ecac)  # futura
    # interfaceEsocial(aba_esocial)  # futura

    root.mainloop()



if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()


    manager = Manager()
    chaves_com_erro = manager.list()


    root = telaPrincipal()

