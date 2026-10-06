@echo off
:: Navega para a pasta do projeto
cd /d "F:\GUIAS\BorderoGuias"

:: Ativa o ambiente virtual
call venv\Scripts\activate

:: Executa o script Python
python bordero.py

:: Desativa o ambiente virtual ao terminar
deactivate

:: Log de erros
python bordero.py > log_execucao.txt 2>&1