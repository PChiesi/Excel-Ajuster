@echo off
REM Arraste uma ou mais planilhas (ou uma pasta) sobre este arquivo.
cd /d "%~dp0"
python -m excel_ajuster %*
if not "%~1"=="" pause
