@echo off
title Compilando Robô - Certezza

REM === Configurações ===
set "NOME_SCRIPT=[v2] Certezza - Robos.py"
set "NOME_EXE=Certezza - Robos"
set "ICON=img\logoICO.ico"
set "PASTA_DIST=dist\Certezza - Robos"

echo ===========================
echo ==== LIMPANDO ARQUIVOS ====
echo ===========================

:: Apagar pastas antigas - NÃO usar durante testes!
REM rmdir /s /q build
REM rmdir /s /q dist
REM del /q "%NOME_EXE%.spec"

echo ===========================
echo ==== COMPILANDO...  =======
echo ===========================
REM 
pyinstaller --onedir --clean --noconfirm --noconsole ^
--name "%NOME_EXE%" ^
--icon="%ICON%" ^
--upx-dir=C:\Ferramentas\upx ^
--add-data "img;img" ^
--add-data "scripts;scripts" ^
--add-data "documentacao\v1.0.1\Documentacao_Interface_Robos_Certezza_v1.0.1.pdf;." ^
--hidden-import=pdfplumber ^
--hidden-import=comtypes.client ^
--hidden-import=screeninfo ^
--hidden-import=screeninfo.screeninfo ^
--hidden-import=ttkbootstrap ^
"%NOME_SCRIPT%"

IF %ERRORLEVEL% NEQ 0 (
    echo =================================
    echo === ERRO NA COMPILACAO ==========
    echo =================================
    pause
    exit /b
)

echo ===========================
echo ==== ORGANIZANDO DIST =====
echo ===========================

REM Cria pasta de distribuicao
if not exist "%PASTA_DIST%" mkdir "%PASTA_DIST%"

REM Move o EXE
move /Y "dist\%NOME_EXE%\%NOME_EXE%.exe" "%PASTA_DIST%"

REM Copia pastas adicionais (se quiser incluir)
REM xcopy "img" "%PASTA_DIST%\img" /E /I /Y
REM xcopy "scripts" "%PASTA_DIST%\scripts" /E /I /Y

echo ===========================
echo ==== COMPILADO COM SUCESSO!
echo ==== Executável em %PASTA_DIST%\%NOME_EXE%.exe
echo ===========================

pause
