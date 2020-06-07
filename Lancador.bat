@echo off
:: Check for Python Installation
python --version 2>NUL
if errorlevel 1 goto errorNoPython
CLS
echo ############################################################################
echo ########################## Python esta instalado ##########################
echo ############################################################################

:verOpenpyxl
pip show openpyxl>NUL
if errorlevel 1 goto errorNoOpenpyxl
echo ############################################################################
echo ##################### Biblioteca OPENPYXL esta instalado ###################
echo ############################################################################


::executa o comando py
echo %~dp0
cls
python "%~dp0main.py"
echo ####### POWERED BY: Bruno B. Linhares ######
pause>nul

:: Reaching here means Python is installed.
:: Execute stuff...

:: Once done, exit the batch file -- skips executing the errorNoPython section
goto:eof

:errorNoPython
echo.
echo ############################################################################
echo ##################### Error^: Python NOT installed #########################
echo ############################################################################
echo .
echo INSTALE O PYTHON E TENTE NOVAMENTE
pause>null
exit

:errorNoOpenpyxl
echo.
echo ############################################################################
echo ####################### Biblioteca nao Instalada ###########################
echo ############################################################################
echo .
echo Instalando...
pip install openpyxl==2.5.12
goto:verOpenpyxl


