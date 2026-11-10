@echo off
cd /d "%~dp0"
echo Starting Internal Data Analysis Agent...
call .venv\Scripts\activate
streamlit run app.py
pause
