@echo off
echo Ativando o ambiente virtual...
call venv\Scripts\activate
echo Ambiente virtual ativado.

echo Executando o arquivo Python...
python main.py

echo Pressione qualquer tecla para fechar.
pause >nul