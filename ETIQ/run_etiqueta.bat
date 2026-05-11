@echo off
cd /d %~dp0
if exist .venv\Scripts\python.exe (
  .venv\Scripts\python.exe etiqueta_app.py
) else (
  python etiqueta_app.py
)
