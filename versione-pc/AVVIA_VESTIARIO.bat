@echo off
cd /d "%~dp0"
py -c "import openpyxl" 2>nul || py -m pip install openpyxl
start "" pyw "%~dp0vestiario.py"
