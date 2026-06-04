@echo off
title Servidor Reportes GU
echo Instalando dependencias...
pip install flask flask-cors --quiet
echo.
echo Iniciando servidor...
python servidor.py
pause
