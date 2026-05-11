@echo off
cd /d C:\Volume\Volume\ETIQ
if exist .venv\Scripts\python.exe (
  .venv\Scripts\python.exe etiqueta_app.py
) else (
  py etiqueta_app.py
)
