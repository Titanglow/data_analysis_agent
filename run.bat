@echo off
chcp 65001 > nul
echo [데이터 분석 Agent] 웹 대시보드를 시작합니다...
call .venv\Scripts\activate
streamlit run app.py
pause