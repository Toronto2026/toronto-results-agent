@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo Запуск агента результатів фестивалю TORONTO...
echo Відкрийте браузер: http://localhost:8501
python -m streamlit run app.py
pause
