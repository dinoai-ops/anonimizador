@echo off
rem Anonimizador Juridico LGPD - app local
rem Abre o app local sem janela de terminal.
cd /d "%~dp0"
start "" pythonw app.py
