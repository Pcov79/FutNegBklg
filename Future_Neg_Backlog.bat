@echo off
SETLOCAL

REM Check if virtual environment exists
IF EXIST "venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call venv\Scripts\activate.bat
)

REM Run the Streamlit app
streamlit run Future_Neg_Backlog.py

ENDLOCAL
pause
