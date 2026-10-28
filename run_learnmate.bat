@echo off
title LearnMate AI
cd /d "%~dp0"
echo ======================================
echo        LearnMate AI Audio Tutor
echo ======================================
where py >nul 2>nul
if %errorlevel%==0 (
  set PYTHON=py
) else (
  set PYTHON=python
)
%PYTHON% -m pip install --upgrade pip
%PYTHON% -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Package installation failed. Check your internet connection and Python installation.
  pause
  exit /b 1
)
echo.
echo Starting LearnMate AI in your browser...
%PYTHON% -m streamlit run app.py --server.port 8501
if errorlevel 1 (
  echo If port 8501 is busy, try: python -m streamlit run app.py --server.port 8502
  pause
)
