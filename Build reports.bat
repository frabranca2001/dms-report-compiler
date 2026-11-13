@echo off
rem Double-click to build the reports without opening an editor.
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" run.py
) else (
    py run.py
)
pause
