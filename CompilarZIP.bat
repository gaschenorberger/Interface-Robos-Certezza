@echo off
title Compilando Robô - Certezza

REM === Configurações ===
set "NOME_SCRIPT=Certezza - Robos.py"
set "NOME_EXE=Certezza - Robos.exe"
set "ICON=img\logoICO.ico"
set "PASTA_DIST=DISTRIBUICAO"
set "NOME_ZIP=Certezza-Robo.zip"

echo ===========================
echo ==== LIMPANDO ARQUIVOS ====
echo ===========================
rmdir /s /q build
rmdir /s /q dist
del /q *.spec
del /q "%NOME_ZIP%"

echo ===========================
echo ==== COMPILANDO...  =======
echo ===========================

pyinstaller --onefile --noconsole ^
--icon="%ICON%" ^
--add-data "img;img" ^
--add-data "scripts;scripts" ^
--hidden-import=comtypes ^
--hidden-import=comtypes.client ^
--hidden-import=comtypes.stream ^
--hidden-import=comtypes.automation ^
--hidden-import=comtypes.typeinfo ^
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
move /Y "dist\%NOME_EXE%" "%PASTA_DIST%"

REM Copia pastas adicionais (se quiser incluir)
xcopy "img" "%PASTA_DIST%\img" /E /I /Y
xcopy "scripts" "%PASTA_DIST%\scripts" /E /I /Y

echo ===========================
echo ==== GERANDO ZIP ==========
echo ===========================

REM Gera o zip usando PowerShell nativo
powershell Compress-Archive -Path "%PASTA_DIST%\*" -DestinationPath "%NOME_ZIP%"

echo ===========================
echo ==== COMPILADO COM SUCESSO!
echo ==== Executável em %PASTA_DIST%\%NOME_EXE%
echo ==== Arquivo ZIP: %NOME_ZIP%
echo ===========================

pause
