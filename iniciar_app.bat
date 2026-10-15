@echo off
REM Inicia la app de certificados con doble clic.
cd /d "%~dp0"
"%LocalAppData%\Programs\Python\Python312\python.exe" app.py
pause
